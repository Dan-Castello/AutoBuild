#Requires -Version 5.1
# =============================================================================
# lib/Logger.ps1
# AutoBuild v3.0 - Structured JSONL logging engine.
#
# AUDIT RESOLUTIONS:
#   PROBLEMA-ARQUITECTURAL-01 : Get-EngineConfig extracted to Config.ps1.
#                               Logger owns ONLY logging.
#   PROBLEMA-LOG-01 (HIGH)    : New-RunId uses [Guid]::NewGuid() — 4.3B
#                               unique values/second. v1 had 456K (collision-
#                               prone under parallel builds).
#   PROBLEMA-LOG-02 (HIGH)    : Automatic log rotation. Triggered when
#                               registry.jsonl exceeds maxLogSizeBytes.
#   PROBLEMA-LOG-03 (HIGH)    : Detail sanitized before JSON encoding.
#                               Newlines -> ' | '. Control chars stripped.
#                               Guarantees valid single-line JSONL entries.
#   PROBLEMA-LOG-04 (MED)     : Timestamps use ISO 8601 with UTC offset (zzz).
#   PROBLEMA-LOG-05 (LOW)     : FATAL level added above ERROR.
#   PROBLEMA-LOG-06 (MED)     : Log rotation uses move (atomic rename) not
#                               content rewrite.
#   CONC-01/02/03 (CRITICAL)  : Write-LogLine checks WaitOne result.
#                               If mutex not acquired -> entry DISCARDED.
#                               Corrupt JSONL is worse than a missing entry.
#   F3-08 (robustness)        : User and hostname captured in every log entry.
#
# MUTEX DESIGN NOTE:
#   Name: 'Global\AutoBuildLogMutex'
#   Prefix 'Global\' ensures cross-session serialization (Build-Parallel).
#   Timeout: 5000ms — generous. Normal contention resolves in <50ms.
#   On timeout: entry is discarded (fail-safe), not written unprotected.
#   QueueRunner.psm1 uses the same mutex name (ACOPLAMIENTO-02 fix).
# =============================================================================
Set-StrictMode -Version Latest

$Script:LogMutexName  = 'Global\AutoBuildLogMutex'
# FIX R-08: Tracks entries discarded due to mutex unavailability.
# Exposed via Get-LogDroppedCount for health monitoring.
$Script:LogDroppedCount = 0

function Get-LogDroppedCount {
    <#
    .SYNOPSIS Returns the count of log entries discarded due to mutex contention or failure. #>
    return $Script:LogDroppedCount
}

function Reset-LogDroppedCount { $Script:LogDroppedCount = 0 }

# ---------------------------------------------------------------------------
# RUN IDENTIFIER
# ---------------------------------------------------------------------------

function New-RunId {
    <#
    .SYNOPSIS
        Returns a globally unique run identifier.
        Format: yyyyMMdd_HHmmss_<8-char-guid-fragment>
    #>
    $ts   = Get-Date -Format 'yyyyMMdd_HHmmss'
    $guid = ([Guid]::NewGuid().ToString('N')).Substring(0, 8).ToUpper()
    return "${ts}_${guid}"
}

# ---------------------------------------------------------------------------
# LOG WRITERS
# ---------------------------------------------------------------------------

function Write-BuildLog {
    <#
    .SYNOPSIS
        Writes a structured JSONL entry to registry.jsonl atomically.
    .PARAMETER Context
        Task execution context produced by New-TaskContext.
    .PARAMETER Level
        DEBUG | INFO | WARN | ERROR | FATAL
    .PARAMETER Message
        Short human-readable message (single line).
    .PARAMETER Detail
        Optional extended detail. Newlines sanitized to preserve JSONL.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Context,
        [ValidateSet('DEBUG','INFO','WARN','ERROR','FATAL')]
        [string]$Level   = 'INFO',
        [Parameter(Mandatory)][string]$Message,
        [string]$Detail  = ''
    )

    # Level filter.
    $levelMap = @{ DEBUG = 0; INFO = 1; WARN = 2; ERROR = 3; FATAL = 4 }
    $cfgLevel = $Context.Config.engine.logLevel
    if ($levelMap[$Level] -lt $levelMap[$cfgLevel]) { return }

    # ISO 8601 with UTC offset. (LOG-04 fix)
    $ts = Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'

    # Console output.
    $color = switch ($Level) {
        'DEBUG' { 'Gray'    }
        'INFO'  { 'Cyan'    }
        'WARN'  { 'Yellow'  }
        'ERROR' { 'Red'     }
        'FATAL' { 'Magenta' }
    }
    Write-Host "[$ts][$Level] $Message" -ForegroundColor $color

    # Sanitize Detail: no newlines, no control chars in JSONL. (LOG-03 fix)
    $safeDetail = Invoke-SanitizeLogText -Text $Detail

    $entry = [ordered]@{
        ts       = $ts
        level    = $Level
        runId    = $Context.RunId
        task     = $Context.TaskName
        user     = $Context.User
        hostname = $Context.Hostname
        message  = $Message
        detail   = $safeDetail
    }
    $json    = $entry | ConvertTo-Json -Compress
    $regFile = Join-Path $Context.Paths.Logs 'registry.jsonl'

    # Rotate if needed before appending. (LOG-02 fix)
    $maxLogBytes = try { [long]$Context.Config.reports.maxLogSizeBytes } catch { 0 }
    if ($maxLogBytes -le 0) { $maxLogBytes = 10485760 }   # 10 MB default
    Invoke-LogRotationIfNeeded -FilePath $regFile -MaxBytes $maxLogBytes

    Write-LogLine -FilePath $regFile -Line $json
}

function Write-RunResult {
    <#
    .SYNOPSIS
        Writes the final outcome entry for a task run to registry.jsonl.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Context,
        [Parameter(Mandatory)][bool]$Success,
        [string]$ErrorMsg = ''
    )

    $ts      = Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'
    $status  = if ($Success) { 'OK' } else { 'ERROR' }
    $elapsed = [math]::Round(([datetime]::Now - $Context.StartTime).TotalSeconds, 3)
    $safeErr = Invoke-SanitizeLogText -Text $ErrorMsg

    $entry = [ordered]@{
        ts       = $ts
        level    = $status
        runId    = $Context.RunId
        task     = $Context.TaskName
        user     = $Context.User
        hostname = $Context.Hostname
        message  = "Run $status"
        detail   = $safeErr
        elapsed  = $elapsed
    }
    $json    = $entry | ConvertTo-Json -Compress
    $regFile = Join-Path $Context.Paths.Logs 'registry.jsonl'

    Write-LogLine -FilePath $regFile -Line $json
}

# ---------------------------------------------------------------------------
# INTERNAL HELPERS
# ---------------------------------------------------------------------------

function Invoke-SanitizeLogText {
    <#
    .SYNOPSIS
        Internal. Replaces newlines and strips control chars from a string.
        Guarantees valid single-line JSON.
    #>
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $s = $Text -replace "`r`n", ' | ' -replace "`n", ' | ' -replace "`r", ' | '
    return $s -replace '[\x00-\x1F\x7F]', ''
}

function Write-LogLine {
    <#
    .SYNOPSIS
        Internal. Appends $Line to $FilePath under a named Mutex.
    .NOTES
        MUTEX CORRECTNESS (resolves CONC-01, CONC-02, CONC-03):
        If WaitOne returns $false (timeout, mutex not acquired), the write
        is SKIPPED. We never write without the mutex held.
        Rationale: corrupted JSONL is worse than a missing log entry.
        This is an intentional design decision.

        FIX R-08 (AUDIT v3): The previous catch-fallback wrote WITHOUT any mutex
        when the Global\ constructor failed — the exact race condition the mutex
        was designed to prevent.

        CORRECTION: Two-tier mutex strategy:
          Tier 1: 'Global\AutoBuildLogMutex'  — cross-session serialization
                  (required for multi-user concurrent builds)
          Tier 2: 'Local\AutoBuildLogMutex'   — same-session fallback
                  (used when Group Policy restricts Global\ named mutexes,
                   e.g. in restricted Pester test environments)

        The write is ALWAYS under a mutex (Tier 1 or Tier 2).
        The entry is ONLY discarded if both mutex tiers fail AND WaitOne times out.
        An unprotected write is never performed.
    #>
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string]$Line
    )

    # Ensure the parent directory exists before any write attempt. (CRITICAL-ISSUE-1 fix)
    $dir = Split-Path $FilePath -Parent
    if ($dir -and -not (Test-Path $dir)) {
        try { New-Item -ItemType Directory -Path $dir -Force | Out-Null } catch { }
    }

    $mutex  = $null
    $locked = $false
    try {
        # Tier 1: Global\ mutex for cross-session serialization.
        try {
            $mutex = New-Object System.Threading.Mutex($false, $Script:LogMutexName)
        } catch {
            # Global\ mutex restricted (e.g. Group Policy, restricted test runner).
            # FIX R-08: fall back to Local\ — still serialized within the session.
            $localName = $Script:LogMutexName -replace '^Global\\', 'Local\'
            $mutex = New-Object System.Threading.Mutex($false, $localName)
            Write-Verbose "Write-LogLine: Global mutex unavailable, using Local mutex. Error: $_"
        }

        $locked = $mutex.WaitOne(5000)
        if ($locked) {
            Add-Content -Path $FilePath -Value $Line -Encoding ASCII
        } else {
            # WaitOne timed out — entry discarded to preserve JSONL integrity.
            # Never write without the mutex held.
            $Script:LogDroppedCount++
            Write-Verbose "Write-LogLine: mutex timeout, entry discarded."
        }
    } catch {
        # Both mutex tiers failed — discard and track. No unprotected write.
        $Script:LogDroppedCount++
        Write-Verbose "Write-LogLine: mutex unavailable (both tiers), entry discarded. Error: $_"
    } finally {
        if ($locked -and $null -ne $mutex) {
            try { $mutex.ReleaseMutex() } catch { }
        }
        if ($null -ne $mutex) {
            try { $mutex.Dispose() } catch { }
        }
    }
}

# ---------------------------------------------------------------------------
# LOG ROTATION AND PURGE
# ---------------------------------------------------------------------------

function Invoke-LogRotationIfNeeded {
    <#
    .SYNOPSIS
        If $FilePath exceeds $MaxBytes, atomically renames it to a
        timestamped archive. Rotation is a fast size check + Move-Item.
        (Resolves PROBLEMA-LOG-02 and PROBLEMA-LOG-06.)
    .NOTES
        NTFS rename (Move-Item same-volume) is atomic: no data loss if
        interrupted mid-rotation. The active log path is then empty for
        the next write.
    #>
    param(
        [string]$FilePath,
        [long]$MaxBytes = 10485760   # 10 MB default
    )

    if (-not (Test-Path $FilePath)) { return }
    try {
        $fi = [System.IO.FileInfo]$FilePath
        if ($fi.Length -lt $MaxBytes) { return }

        $dir     = $fi.DirectoryName
        $base    = [System.IO.Path]::GetFileNameWithoutExtension($fi.Name)
        $stamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
        $archive = Join-Path $dir "${base}_${stamp}.jsonl"

        # Atomic rename: current log becomes archive, path is vacated for new entries.
        Move-Item -Path $FilePath -Destination $archive -Force -ErrorAction Stop
    } catch {
        # Rotation failure is non-fatal. Logging continues on the oversized file.
    }
}

function Invoke-LogPurge {
    <#
    .SYNOPSIS
        Deletes rotated archive log files older than $RetentionDays.
        Operates only on timestamped archive files (pattern: registry_*.jsonl).
        Never touches the active registry.jsonl.
    .PARAMETER LogsDir
        Path to the logs/ directory.
    .PARAMETER RetentionDays
        Delete archives older than this many days.
    #>
    param(
        [Parameter(Mandatory)][string]$LogsDir,
        [int]$RetentionDays = 30
    )

    $cutoff = [datetime]::Now.AddDays(-$RetentionDays)

    Get-ChildItem -Path $LogsDir -Filter 'registry_*.jsonl' -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt $cutoff } |
        ForEach-Object {
            try { Remove-Item $_.FullName -Force -ErrorAction Stop } catch { }
        }
}
