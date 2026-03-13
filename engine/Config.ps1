#Requires -Version 5.1
# =============================================================================
# lib/Config.ps1
# AutoBuild v3.1 - Configuration loader, merger, validation, and bootstrap.
#
# REMEDIATION v3.1 CHANGES:
#   FIX-CFG-01 (CRITICAL) : Test-EngineConfiguration added.
#     Previously, missing SMTP config caused diag_notify to FAIL the check and
#     throw. SMTP is OPTIONAL. The new validator distinguishes CRITICAL vs
#     OPTIONAL issues and returns SmtpConfigured flag. diag_notify branches
#     on SmtpConfigured instead of counting SMTP absence as a failure.
#
#   FIX-CFG-02 (HIGH)  : All config sections now have explicit defaults for
#     every key. Missing JSON keys never produce $null downstream.
#
#   FIX-CFG-03 (MED)   : Deep merge: partial JSON sections no longer discard
#     built-in defaults for absent keys.
# =============================================================================
Set-StrictMode -Version Latest

$Script:KnownSections = @('engine','sap','excel','reports','notifications','security')

function Get-EngineConfig {
    param([Parameter(Mandatory)][string]$Root)

    $cfg = @{
        engine = @{
            logLevel          = 'INFO'
            maxRetries        = 3
            retryDelaySeconds = 5
            maxConcurrentJobs = 4
            ibVersion         = 'unknown'
        }
        sap = @{
            systemId = 'PRD'
            client   = '800'
            language = 'ES'
            timeout  = 180
        }
        excel = @{
            visible        = $false
            screenUpdating = $false
        }
        reports = @{
            defaultFormat   = 'xlsx'
            retentionDays   = 30
            maxLogSizeBytes = 10485760
        }
        # FIX-CFG-01: Empty smtpServer = OPTIONAL-MISSING, never a CRITICAL failure.
        notifications = @{
            enabled      = $false
            smtpServer   = ''
            smtpPort     = 587
            fromAddr     = ''
            toAddr       = ''
            smtpUser     = ''
            smtpPassword = ''
            useTls       = $false
        }
        security = @{
            adminAdGroup        = ''
            developerAdGroup    = ''
            adminUsers          = ''
            developerUsers      = ''
            roleCacheTTLMinutes = 480
        }
    }

    $cfgFile = Join-Path $Root 'engine.config.json'
    if (Test-Path $cfgFile) {
        try {
            $raw = Get-Content $cfgFile -Raw -ErrorAction Stop | ConvertFrom-Json
            foreach ($section in $Script:KnownSections) {
                $rawSection = $raw.PSObject.Properties[$section]
                if ($null -ne $rawSection -and $null -ne $rawSection.Value) {
                    if (-not $cfg.ContainsKey($section)) { $cfg[$section] = @{} }
                    # FIX-CFG-03: Key-by-key merge preserves defaults for absent JSON keys.
                    foreach ($prop in $rawSection.Value.PSObject.Properties) {
                        $cfg[$section][$prop.Name] = $prop.Value
                    }
                }
            }
        } catch {
            Write-Warning "AutoBuild Config: Cannot read engine.config.json: $_"
        }
    }

    try {
        $ibFile = Join-Path $Root 'tools\InvokeBuild\Invoke-Build.ps1'
        if (Test-Path $ibFile) {
            $ibMatch = Select-String -Path $ibFile -Pattern 'version\s+([\d\.]+)' |
                       Select-Object -First 1
            if ($ibMatch) { $cfg.engine.ibVersion = $ibMatch.Matches.Groups[1].Value }
        }
    } catch { }

    foreach ($sub in @('logs','input','output','reports')) {
        $dir = Join-Path $Root $sub
        if (-not (Test-Path $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
    }

    return $cfg
}

function Get-ConfigSection {
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$Section
    )
    if ($Config.ContainsKey($Section)) {
        return $Config[$Section].Clone()
    }
    return @{}
}

function Test-EngineConfiguration {
    <#
    .SYNOPSIS
        FIX-CFG-01: Validates engine configuration. Classifies issues as
        CRITICAL (engine cannot operate) or OPTIONAL (feature disabled/unconfigured).

    .DESCRIPTION
        Returns a structured result hashtable. Tasks should branch on
        SmtpConfigured / Valid rather than failing when OPTIONAL settings
        are absent.

        CRITICAL: engine section, task dirs, log dir
        OPTIONAL: SMTP, Excel COM, SAP, Security (AD groups)

    .OUTPUTS
        @{
            Valid           = [bool]
            CriticalIssues  = [string[]]
            OptionalIssues  = [string[]]
            SmtpConfigured  = [bool]   # $true only when smtpServer non-empty
            ExcelAvailable  = [bool]   # $true when COM probe succeeded
            Summary         = [string]
        }
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [string]$Root         = '',
        [bool]$ProbeExcel     = $false,
        [int]$ExcelTimeoutSec = 10
    )

    # Use List[string] — safe against PS5.1 op_Addition on typed arrays.
    $critical = [System.Collections.Generic.List[string]]::new()
    $optional  = [System.Collections.Generic.List[string]]::new()

    # --- CRITICAL: engine section ---
    $eng = $Config['engine']
    if ($null -eq $eng) {
        $critical.Add('engine config section is missing entirely')
    } else {
        $validLevels = @('DEBUG','INFO','WARN','ERROR','FATAL')
        $ll = "$($eng['logLevel'])"
        if ([string]::IsNullOrWhiteSpace($ll) -or $validLevels -notcontains $ll) {
            $critical.Add("engine.logLevel invalid (got: '$ll'). Must be: $($validLevels -join ', ')")
        }
        $mr = $eng['maxRetries']
        if ($null -eq $mr -or [int]$mr -lt 0) {
            $critical.Add("engine.maxRetries must be >= 0 (got: '$mr')")
        }
    }

    # --- CRITICAL: directories ---
    if (-not [string]::IsNullOrWhiteSpace($Root)) {
        $tasksDir = Join-Path $Root 'tasks'
        if (-not (Test-Path $tasksDir)) {
            $critical.Add("tasks/ directory not found: $tasksDir")
        }
        $logsDir = Join-Path $Root 'logs'
        if (-not (Test-Path $logsDir)) {
            try {
                New-Item -ItemType Directory -Path $logsDir -Force | Out-Null
            } catch {
                $critical.Add("logs/ directory missing and cannot be created: $_")
            }
        }
    }

    # --- OPTIONAL: SMTP ---
    # FIX-CFG-01 CORE: Empty smtpServer is OPTIONAL-MISSING, not a failure.
    $smtpConfigured = $false
    $notif = $Config['notifications']
    if ($null -eq $notif) {
        $optional.Add('Notifications disabled: notifications section missing from config')
    } elseif ([string]::IsNullOrWhiteSpace("$($notif['smtpServer'])")) {
        $optional.Add('Notifications disabled: smtpServer not configured in engine.config.json')
    } else {
        $smtpConfigured = $true
        $port = try { [int]$notif['smtpPort'] } catch { 0 }
        if ($port -lt 1 -or $port -gt 65535) {
            $optional.Add("Notifications: smtpPort '$($notif['smtpPort'])' invalid. Default 587 will be used.")
        }
        if ([string]::IsNullOrWhiteSpace("$($notif['fromAddr'])")) {
            $optional.Add('Notifications: fromAddr not set. Fallback autobuild@<hostname> will be used.')
        }
        if ([string]::IsNullOrWhiteSpace("$($notif['toAddr'])")) {
            $optional.Add('Notifications: toAddr not set. Notifications cannot be delivered.')
        }
    }

    # --- OPTIONAL: Excel COM probe ---
    $excelAvailable = $false
    if ($ProbeExcel) {
        # Test-ComAvailable is defined in ComHelper.ps1 (loaded before Config in engine)
        $excelAvailable = Test-ComAvailable -ProgId 'Excel.Application' -TimeoutSec $ExcelTimeoutSec
        if (-not $excelAvailable) {
            $optional.Add('Excel COM unavailable. Excel-dependent tasks will be skipped.')
        }
    }

    # --- OPTIONAL: SAP ---
    $sap = $Config['sap']
    if ($null -eq $sap -or [string]::IsNullOrWhiteSpace("$($sap['systemId'])")) {
        $optional.Add('SAP not configured. SAP tasks will fail if executed.')
    }

    # --- OPTIONAL: Security ---
    $sec = $Config['security']
    if ($null -ne $sec) {
        $hasAdGroup  = -not [string]::IsNullOrWhiteSpace("$($sec['adminAdGroup'])")
        $hasFallback = -not [string]::IsNullOrWhiteSpace("$($sec['adminUsers'])")
        if (-not $hasAdGroup -and -not $hasFallback) {
            $optional.Add('Security: adminAdGroup and adminUsers both empty. Running in open-access dev mode.')
        }
    }

    $isValid = ($critical.Count -eq 0)
    $summary = if ($isValid) {
        "Configuration OK. Optional issues: $($optional.Count)."
    } else {
        "Configuration INVALID. Critical: $($critical.Count). Optional: $($optional.Count)."
    }

    return @{
        Valid          = $isValid
        CriticalIssues = $critical.ToArray()
        OptionalIssues = $optional.ToArray()
        SmtpConfigured = $smtpConfigured
        ExcelAvailable = $excelAvailable
        Summary        = $summary
    }
}
