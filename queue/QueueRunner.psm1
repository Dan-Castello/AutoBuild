#Requires -Version 5.1
# =============================================================================
# queue/QueueRunner.psm1  v3.0
# AutoBuild - Sequential task queue execution engine.
#
# COMPLETE REWRITE — resolves all audit findings plus supplies every function
# that QueueGUI.psm1 requires.
#
# AUDIT RESOLUTIONS vs v1:
#   CONC-02 (HIGH)      : Write-QueueLog checks WaitOne — fail-safe discard.
#   ACOPLAMIENTO-02(MED): Same mutex name as Logger.ps1.
#   RUN-01 (HIGH)       : Start-TaskProcess serialises params to JSON -Params.
#   MAINT-01            : Poll-ActiveTask renamed Step-PollActiveTask.
#   SCALE-04 note       : State is process-local. Future: SQLite backend.
#   LOG-03/04           : Detail sanitised; timestamps include UTC offset.
#
# NEW IN v3 vs v2 partial:
#   UseWpfTimer support in Start-QueueRunner (required by QueueGUI).
#   AutoAdvance flag — runner ticks autonomously when true.
#   Invoke-QueueRunnerTick — the WPF DispatcherTimer tick body.
#   Suspend-Queue / Resume-Queue — pause/unpause without stopping.
#   Start-AllQueueTasks / Start-SelectedQueueTasks / Start-NextQueueTask.
#   Skip-QueueTask / Stop-ActiveTask.
#   PendingCount and ActiveTask fields in Get-QueueRunnerState.
#   PollTimer in runner state so QueueGUI can stop it on window close.
# =============================================================================
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$Script:Runner = @{
    IsRunning       = $false
    IsPaused        = $false
    IsStopped       = $true
    AutoAdvance     = $false
    ActiveTaskId    = $null
    ActiveProcess   = $null
    ActiveQueue     = $null
    ActiveBuffer    = $null
    ActiveStarted   = $null
    ActiveOutHandler = $null   # FIX STRUCTURAL-2: stored for explicit unregistration
    ActiveErrHandler = $null   # FIX STRUCTURAL-2: stored for explicit unregistration
    PollTimer       = $null
    EngineRoot      = ''
    RunScript       = ''
    RegistryFile    = ''
}

$Script:QueueEvents = @{
    OnTaskStarted   = $null
    OnTaskCompleted = $null
    OnTaskFailed    = $null
    OnQueueEmpty    = $null
    OnStateChanged  = $null
}

function Write-QueueLog {
    param(
        [string]$TaskName  = 'QUEUE',
        [string]$Level     = 'INFO',
        [string]$Message   = '',
        [string]$Detail    = '',
        [string]$EngineRunId = ''   # FIX OBS: accept correlated RunId from Logger.ps1
    )
    if ([string]::IsNullOrWhiteSpace($Script:Runner.RegistryFile)) { return }
    if (-not (Test-Path (Split-Path $Script:Runner.RegistryFile -Parent))) { return }

    $ts         = Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'
    $safeDetail = $Detail -replace "`r`n",' | ' -replace "`n",' | ' -replace "`r",' | '
    $safeDetail = $safeDetail -replace '[\x00-\x1F\x7F]', ''

    # FIX OBS: use correlated RunId when provided, else generate queue-local id.
    $runId = if (-not [string]::IsNullOrWhiteSpace($EngineRunId)) {
        $EngineRunId
    } else {
        'Q_{0}' -f (Get-Date -Format 'HHmmssfff')
    }

    $entry = [ordered]@{
        ts      = $ts
        runId   = $runId
        task    = $TaskName
        level   = $Level
        message = $Message
        detail  = $safeDetail
    }
    $line = $entry | ConvertTo-Json -Compress

    $mutex  = $null
    $locked = $false
    try {
        # Two-tier mutex: Global\ for cross-session, Local\ fallback for restricted environments.
        try {
            $mutex = New-Object System.Threading.Mutex($false, 'Global\AutoBuildLogMutex')
        } catch {
            $mutex = New-Object System.Threading.Mutex($false, 'Local\AutoBuildLogMutex')
        }
        $locked = $mutex.WaitOne(5000)
        if ($locked) {
            Add-Content -Path $Script:Runner.RegistryFile -Value $line -Encoding ASCII
        }
        # FIX R-08: No fallback unprotected write. Discard on timeout.
    } catch {
        # Both mutex tiers failed — discard. No unprotected write.
        Write-Verbose "Write-QueueLog: mutex unavailable, entry discarded. Error: $_"
    } finally {
        if ($locked -and $null -ne $mutex) { try { $mutex.ReleaseMutex() } catch { } }
        if ($null -ne $mutex) { try { $mutex.Dispose() } catch { } }
    }
}

function Invoke-Event {
    param([string]$Name, [object[]]$Args = @())
    $h = $Script:QueueEvents[$Name]
    if ($null -ne $h) { try { & $h @Args } catch { } }
}

function Clear-ActiveProcess {
    <#
    .SYNOPSIS
        FIX STRUCTURAL-2 (AUDIT v3): Unregisters DataReceived event handlers BEFORE
        disposing the process object. Without explicit removal, handlers remain
        registered in the GC root — accumulating across task cancellations and
        causing unbounded memory growth in long-running UI sessions.
    #>
    $proc = $Script:Runner.ActiveProcess
    if ($null -ne $proc) {
        # Unregister handlers before Dispose() to allow GC collection.
        if ($null -ne $Script:Runner.ActiveOutHandler) {
            try { $proc.remove_OutputDataReceived($Script:Runner.ActiveOutHandler) } catch { }
        }
        if ($null -ne $Script:Runner.ActiveErrHandler) {
            try { $proc.remove_ErrorDataReceived($Script:Runner.ActiveErrHandler) } catch { }
        }
        try { $proc.Dispose() } catch { }
    }
    $Script:Runner.ActiveProcess    = $null
    $Script:Runner.ActiveOutHandler = $null
    $Script:Runner.ActiveErrHandler = $null
    $Script:Runner.ActiveQueue      = $null
    $Script:Runner.ActiveBuffer     = $null
    $Script:Runner.ActiveStarted    = $null
    $Script:Runner.ActiveTaskId     = $null
}

function Start-TaskProcess {
    param([hashtable]$Task)
    if (-not (Test-Path $Script:Runner.RunScript)) {
        Write-Warning "QueueRunner: Run.ps1 not found at '$($Script:Runner.RunScript)'"
        return $false
    }

    $paramsJson = '{}'
    if ($Task.Parameters -and $Task.Parameters.Count -gt 0) {
        $paramsJson = $Task.Parameters | ConvertTo-Json -Compress
    }

    # FIX V-01/R-01 (AUDIT v3 CRITICAL): Previously this code manually escaped
    # the JSON string with -replace '"', '\"' and embedded it in $argString.
    # A parameter value containing \" could break the argstring parsing and
    # inject arbitrary arguments into the child powershell.exe process.
    #
    # FIX NET-COMPAT (CRITICAL): Same as AutoBuild.UI.ps1.
    # ProcessStartInfo.ArgumentList does not exist in .NET Framework 4.x.
    # The comment "available in .NET 4.5+" was factually incorrect —
    # ArgumentList was introduced in .NET Core 2.1+, never in .NET Framework.
    # Using it on PS 5.1 Desktop throws PropertyNotFoundException at runtime.
    #
    # Fix: use ProcessStartInfo.Arguments (string) with Windows argv-quoting.
    $quoteArg = {
        param([string]$t)
        if ($t -match '[ \t"]') { return '"' + ($t -replace '"', '""') + '"' }
        return $t
    }
    $argTokens = [System.Collections.Generic.List[string]]::new()
    $argTokens.Add('-NoProfile')
    $argTokens.Add('-NonInteractive')
    $argTokens.Add('-ExecutionPolicy')
    $argTokens.Add('Bypass')
    $argTokens.Add('-File')
    $argTokens.Add($Script:Runner.RunScript)
    $argTokens.Add('-Task')
    $argTokens.Add($Task.TaskReference)
    $argTokens.Add('-Params')
    $argTokens.Add($paramsJson)

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName               = 'powershell.exe'
    $psi.Arguments              = ($argTokens | ForEach-Object { & $quoteArg $_ }) -join ' '
    $psi.UseShellExecute        = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.CreateNoWindow         = $true
    $psi.WorkingDirectory       = $Script:Runner.EngineRoot
    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo           = $psi
    $proc.EnableRaisingEvents = $true

    $q   = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
    $buf = [System.Text.StringBuilder]::new(4096)

    $outH = [System.Diagnostics.DataReceivedEventHandler]{ param($s,$e); if ($null -ne $e.Data) { $q.Enqueue($e.Data) } }
    $errH = [System.Diagnostics.DataReceivedEventHandler]{ param($s,$e); if ($null -ne $e.Data) { $q.Enqueue("[ERR] $($e.Data)") } }
    $proc.add_OutputDataReceived($outH)
    $proc.add_ErrorDataReceived($errH)

    # Store handlers so they can be explicitly removed on cleanup. (FIX STRUCTURAL-2)
    $Script:Runner.ActiveOutHandler = $outH
    $Script:Runner.ActiveErrHandler = $errH

    try {
        [void]$proc.Start()
        $proc.BeginOutputReadLine()
        $proc.BeginErrorReadLine()
    } catch {
        Write-Warning "QueueRunner: process start failed: $_"
        return $false
    }

    $Script:Runner.ActiveTaskId  = $Task.TaskId
    $Script:Runner.ActiveProcess = $proc
    $Script:Runner.ActiveQueue   = $q
    $Script:Runner.ActiveBuffer  = $buf
    $Script:Runner.ActiveStarted = [datetime]::Now
    return $true
}

function Step-PollActiveTask {
    if ($null -eq $Script:Runner.ActiveTaskId) { return $false }
    $proc = $Script:Runner.ActiveProcess
    if ($null -eq $proc) { return $false }

    $line = $null
    while ($Script:Runner.ActiveQueue.TryDequeue([ref]$line)) {
        [void]$Script:Runner.ActiveBuffer.AppendLine($line)
    }

    $taskItem = Get-QueueTask -TaskId $Script:Runner.ActiveTaskId
    if ($null -ne $taskItem -and [int]$taskItem.TimeoutSeconds -gt 0) {
        $elapsed = ([datetime]::Now - $Script:Runner.ActiveStarted).TotalSeconds
        if ($elapsed -ge [int]$taskItem.TimeoutSeconds) {
            try { if (-not $proc.HasExited) { $proc.Kill() } } catch { }
        }
    }

    if (-not $proc.HasExited) { return $true }

    Start-Sleep -Milliseconds 100
    while ($Script:Runner.ActiveQueue.TryDequeue([ref]$line)) {
        [void]$Script:Runner.ActiveBuffer.AppendLine($line)
    }

    $exitCode = $proc.ExitCode
    $output   = $Script:Runner.ActiveBuffer.ToString()
    $tid      = $Script:Runner.ActiveTaskId
    $taskRef  = Get-QueueTask -TaskId $tid

    # FIX STRUCTURAL-2: Unregister handlers before Dispose via helper.
    Clear-ActiveProcess

    if ($exitCode -eq 0) {
        Set-QueueTaskStatus -TaskId $tid -Status 'Completed' -Result $output
        Write-QueueLog -TaskName $taskRef.Name -Level 'OK' -Message 'Task completed successfully.'
        Invoke-Event -Name 'OnTaskCompleted' -Args @((Get-QueueTask -TaskId $tid))
    } else {
        $errMsg = "ExitCode=$exitCode"
        if ($null -ne $taskRef -and [int]$taskRef.RetryCount -lt [int]$taskRef.MaxRetries) {
            Set-QueueTaskStatus -TaskId $tid -Status 'Queued' -IncrementRetry
            Write-QueueLog -TaskName $taskRef.Name -Level 'WARN' `
                -Message "Retry $([int]$taskRef.RetryCount + 1)/$($taskRef.MaxRetries)."
        } else {
            Set-QueueTaskStatus -TaskId $tid -Status 'Failed' -Result $output -ErrorMessage $errMsg
            Write-QueueLog -TaskName $taskRef.Name -Level 'ERROR' `
                -Message 'Task failed after all retries.' -Detail $errMsg
            Invoke-Event -Name 'OnTaskFailed' -Args @((Get-QueueTask -TaskId $tid), $errMsg)
        }
    }
    return $false
}

function Step-Queue {
    if ($Script:Runner.IsPaused -or $Script:Runner.IsStopped) { return }
    if ($null -ne $Script:Runner.ActiveTaskId) { return }

    $next = Get-NextPendingTask
    if ($null -eq $next) {
        $pending = @(Get-QueueSnapshot | Where-Object { $_.Status -in @('Pending','Queued') })
        if ($pending.Count -eq 0) { Invoke-Event -Name 'OnQueueEmpty' }
        return
    }

    Set-QueueTaskStatus -TaskId $next.TaskId -Status 'Running'
    Write-QueueLog -TaskName $next.Name -Level 'INFO' `
        -Message "Starting (Order=$($next.Order) Priority=$($next.Priority))"
    Invoke-Event -Name 'OnTaskStarted' -Args @((Get-QueueTask -TaskId $next.TaskId))

    $ok = Start-TaskProcess -Task $next
    if (-not $ok) {
        Set-QueueTaskStatus -TaskId $next.TaskId -Status 'Failed' `
            -ErrorMessage 'Process could not start.'
        Invoke-Event -Name 'OnTaskFailed' -Args @((Get-QueueTask -TaskId $next.TaskId), 'Process not started')
    }
}

# ============================================================================
# PUBLIC API
# ============================================================================

function Set-QueueRunnerConfig {
    param(
        [Parameter(Mandatory)][string]$EngineRoot,
        [hashtable]$EventHandlers = @{}
    )
    $Script:Runner.EngineRoot   = $EngineRoot
    $Script:Runner.RunScript    = Join-Path $EngineRoot 'Run.ps1'
    $Script:Runner.RegistryFile = Join-Path $EngineRoot 'logs\registry.jsonl'
    foreach ($evt in $EventHandlers.Keys) {
        if ($Script:QueueEvents.ContainsKey($evt)) { $Script:QueueEvents[$evt] = $EventHandlers[$evt] }
    }
}

function Start-QueueRunner {
    param(
        [switch]$UseWpfTimer,
        [int]$PollIntervalMs = 800,
        [bool]$AutoAdvance   = $true
    )
    $Script:Runner.IsStopped   = $false
    $Script:Runner.IsPaused    = $false
    $Script:Runner.IsRunning   = $true
    $Script:Runner.AutoAdvance = $AutoAdvance
    Invoke-Event -Name 'OnStateChanged' -Args @($Script:Runner)

    if ($UseWpfTimer) {
        if ($null -ne $Script:Runner.PollTimer) { $Script:Runner.PollTimer.Stop(); $Script:Runner.PollTimer = $null }
        $t = New-Object System.Windows.Threading.DispatcherTimer
        $t.Interval = [TimeSpan]::FromMilliseconds($PollIntervalMs)
        $t.Add_Tick({ Invoke-QueueRunnerTick })
        $t.Start()
        $Script:Runner.PollTimer = $t
    }
}

function Stop-QueueRunner {
    if ($null -ne $Script:Runner.PollTimer) { $Script:Runner.PollTimer.Stop(); $Script:Runner.PollTimer = $null }
    $Script:Runner.IsStopped = $true
    $Script:Runner.IsRunning = $false
    $Script:Runner.IsPaused  = $false

    if ($null -ne $Script:Runner.ActiveTaskId) {
        $tid = $Script:Runner.ActiveTaskId
        if ($null -ne $Script:Runner.ActiveProcess) {
            try {
                if (-not $Script:Runner.ActiveProcess.HasExited) { $Script:Runner.ActiveProcess.Kill() }
            } catch { }
        }
        Set-QueueTaskStatus -TaskId $tid -Status 'Canceled' -ErrorMessage 'Stop-QueueRunner called.'
        # FIX STRUCTURAL-2: unregister handlers before Dispose().
        Clear-ActiveProcess
    }
    Invoke-Event -Name 'OnStateChanged' -Args @($Script:Runner)
}

function Suspend-Queue {
    $Script:Runner.IsPaused = $true
    Invoke-Event -Name 'OnStateChanged' -Args @($Script:Runner)
}

function Resume-Queue {
    if ($Script:Runner.IsStopped) { Write-Warning 'QueueRunner: use Start-QueueRunner first.'; return }
    $Script:Runner.IsPaused = $false
    Invoke-Event -Name 'OnStateChanged' -Args @($Script:Runner)
    Step-Queue
}

function Invoke-QueueRunnerTick {
    if ($Script:Runner.IsStopped) { return }
    $still = $false
    if ($null -ne $Script:Runner.ActiveTaskId) { $still = Step-PollActiveTask }
    if (-not $still -and $Script:Runner.AutoAdvance -and -not $Script:Runner.IsPaused) { Step-Queue }
}

function Start-AllQueueTasks {
    $snap = Get-QueueSnapshot
    foreach ($item in ($snap | Where-Object { $_.Status -eq 'Pending' })) {
        Set-QueueTaskStatus -TaskId $item.TaskId -Status 'Queued'
    }
    if ($Script:Runner.IsStopped) { Start-QueueRunner -UseWpfTimer -AutoAdvance $true }
    Resume-Queue
}

function Start-SelectedQueueTasks {
    param([Parameter(Mandatory)][string[]]$TaskIds)
    $n = 0
    foreach ($id in $TaskIds) {
        $item = Get-QueueTask -TaskId $id
        if ($null -ne $item -and $item.Status -eq 'Pending') { Set-QueueTaskStatus -TaskId $id -Status 'Queued'; $n++ }
    }
    if ($n -gt 0) {
        if ($Script:Runner.IsStopped) { Start-QueueRunner -UseWpfTimer -AutoAdvance $true }
        Resume-Queue
    }
}

function Start-NextQueueTask {
    if ($null -ne $Script:Runner.ActiveTaskId) { Write-Warning 'QueueRunner: task already running.'; return $false }
    $next = Get-NextPendingTask
    if ($null -eq $next) { return $false }
    $wasPaused = $Script:Runner.IsPaused
    $Script:Runner.IsPaused = $false
    Step-Queue
    if ($wasPaused) { $Script:Runner.IsPaused = $true }
    return $true
}

function Skip-QueueTask {
    param([Parameter(Mandatory)][string]$TaskId)
    $item = Get-QueueTask -TaskId $TaskId
    if ($null -eq $item) { return $false }
    if ($item.Status -notin @('Pending','Queued')) { Write-Warning 'Can only skip Pending or Queued tasks.'; return $false }
    Set-QueueTaskStatus -TaskId $TaskId -Status 'Skipped'
    return $true
}

function Stop-ActiveTask {
    if ($null -eq $Script:Runner.ActiveTaskId) { return $false }
    $tid = $Script:Runner.ActiveTaskId
    if ($null -ne $Script:Runner.ActiveProcess) {
        try {
            if (-not $Script:Runner.ActiveProcess.HasExited) { $Script:Runner.ActiveProcess.Kill() }
        } catch { }
    }
    # FIX STRUCTURAL-2: use Clear-ActiveProcess to unregister event handlers before Dispose().
    Clear-ActiveProcess
    Set-QueueTaskStatus -TaskId $tid -Status 'Canceled' -ErrorMessage 'Canceled by user.'
    return $true
}

function Get-QueueRunnerState {
    $active = $null
    if ($null -ne $Script:Runner.ActiveTaskId) {
        $t = Get-QueueTask -TaskId $Script:Runner.ActiveTaskId
        if ($null -ne $t) {
            $elapsed = if ($null -ne $Script:Runner.ActiveStarted) {
                [math]::Round(([datetime]::Now - $Script:Runner.ActiveStarted).TotalSeconds, 1)
            } else { 0 }
            $line = $null
            while ($null -ne $Script:Runner.ActiveQueue -and $Script:Runner.ActiveQueue.TryDequeue([ref]$line)) {
                [void]$Script:Runner.ActiveBuffer.AppendLine($line)
            }
            $active = [PSCustomObject]@{
                TaskId  = $t.TaskId
                Name    = $t.Name
                Elapsed = $elapsed
                Output  = $Script:Runner.ActiveBuffer.ToString()
            }
        }
    }
    return [PSCustomObject]@{
        IsRunning    = $Script:Runner.IsRunning
        IsPaused     = $Script:Runner.IsPaused
        IsStopped    = $Script:Runner.IsStopped
        AutoAdvance  = $Script:Runner.AutoAdvance
        ActiveTask   = $active
        PendingCount = @(Get-QueueSnapshot | Where-Object { $_.Status -in @('Pending','Queued') }).Count
        PollTimer    = $Script:Runner.PollTimer
    }
}

Export-ModuleMember -Function @(
    'Set-QueueRunnerConfig',
    'Start-QueueRunner',    'Stop-QueueRunner',
    'Suspend-Queue',        'Resume-Queue',
    'Invoke-QueueRunnerTick',
    'Start-AllQueueTasks',  'Start-SelectedQueueTasks', 'Start-NextQueueTask',
    'Skip-QueueTask',       'Stop-ActiveTask',
    'Get-QueueRunnerState',
    'Step-Queue',           'Step-PollActiveTask',
    'Write-QueueLog'
)
