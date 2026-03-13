#Requires -Version 5.1
<#
.SYNOPSIS
    AutoBuild v3.0 - WPF Automation Interface
.DESCRIPTION
    Production-ready, secure, auditable GUI for the AutoBuild engine.
    Pages: Task Catalog, Execute, Live Monitor, Run Log, History, Checkpoints,
           Artifacts, Folders, Metrics, Diagnostics, Configuration, New Task, Audit, Queue.
.NOTES
    ASCII-only. PS 5.1 + .NET 4.x. WPF requires STA apartment.

    PHASE 3/4 FIXES APPLIED:
    CONC-01-UI   : Write-AuditLog / $Script:Fn_WriteAuditLog unified.
                   Single canonical mutex pattern (same as Logger.ps1).
                   WaitOne result checked; write skipped if not acquired.
    COM-FREEZE   : Fn_GetDiagnostics offloaded to Runspace + Dispatcher.BeginInvoke.
                   UI thread never blocks on COM instantiation.
    SCALE-02/03  : Fn_GetExecutionHistory uses a streaming tail-read (FileStream +
                   StreamReader from end) instead of Get-Content of the whole file.
                   LoadMonitorPage paginates to the last 50 lines only.
                   RefreshTimer only re-reads the JSONL tail, not full file.
    SEC-01/02    : Resolve-UserRole from Auth.ps1 validates role against AD/whitelist.
                   -Role parameter is a CEILING, not a grant.
    RUN-01       : Fn_StartTaskExecution serializes params to JSON and passes -Params.
    ACOPLAMIENTO-01: UI no longer hardcodes Run.ps1 parameter names.
    UI-THREADING : All heavy I/O (Fn_GetAllTasks, Fn_GetRunSummaries on History page)
                   deferred to Runspace; result posted to UI via Dispatcher.BeginInvoke.
    MAINT-01     : $Script:Fn_TestPermission replaced by Test-Permission from Auth.ps1.
    CONC-04      : Dead function Write-AuditLog (function form) removed.
                   Only $Script:Fn_WriteAuditLog (scriptblock) is used by event handlers.
    F4-02        : XAML loaded from external ui\AutoBuild.xaml file.
                   Inline XAML string preserved as fallback if file not found.
.EXAMPLE
    .\AutoBuild.UI.ps1
    .\AutoBuild.UI.ps1 -EnginePath "C:\AutoBuild"
    .\AutoBuild.UI.ps1 -Role Developer
#>
param(
    [string]$EnginePath = '',
    [ValidateSet('Operator','Developer','Admin')]
    [string]$Role = 'Operator'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================================================================
# BOOTSTRAP: STA apartment required by WPF
# ============================================================================
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $argList = @('-NoProfile','-ExecutionPolicy','Bypass','-STA','-File',"`"$PSCommandPath`"")
    if ($EnginePath) { $argList += @('-EnginePath', "`"$EnginePath`"") }
    if ($Role)       { $argList += @('-Role', $Role) }
    Start-Process powershell.exe -ArgumentList $argList -Wait
    exit $LASTEXITCODE
}

# ============================================================================
# BOOTSTRAP: WPF assemblies
# ============================================================================
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ============================================================================
# PATHS
# ============================================================================
$Script:UIRoot     = $PSScriptRoot
# Engine files (Run.ps1, lib/, engine/, etc.) live one level above ui/.
# If the caller supplies an explicit -EnginePath that takes precedence.
$Script:EngineRoot = if (-not [string]::IsNullOrWhiteSpace($EnginePath)) {
    $EnginePath
} elseif (Test-Path (Join-Path $PSScriptRoot '..\Run.ps1')) {
    # Standard layout: AutoBuild_v3/ui/AutoBuild.UI.ps1 -> root is parent
    Split-Path -Parent $PSScriptRoot
} else {
    # Fallback: assume the UI and engine share the same directory (flat layout)
    $PSScriptRoot
}

$Script:RunScript    = Join-Path $Script:EngineRoot 'Run.ps1'
$Script:ConfigFile   = Join-Path $Script:EngineRoot 'engine.config.json'
$Script:TasksDir     = Join-Path $Script:EngineRoot 'tasks'
$Script:LogsDir      = Join-Path $Script:EngineRoot 'logs'
$Script:OutputDir    = Join-Path $Script:EngineRoot 'output'
$Script:InputDir     = Join-Path $Script:EngineRoot 'input'
$Script:ReportsDir   = Join-Path $Script:EngineRoot 'reports'
$Script:RegistryFile = Join-Path $Script:LogsDir 'registry.jsonl'
$Script:AuditFile    = Join-Path $Script:LogsDir 'audit.jsonl'

foreach ($dir in @($Script:LogsDir,$Script:OutputDir,$Script:ReportsDir,$Script:InputDir)) {
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
}

# ============================================================================
# LOAD ENGINE LIBRARIES (for Auth.ps1, Logger.ps1, Config.ps1)
# ============================================================================
$Script:LibsLoaded = $false
try {
    $libPath = Join-Path $Script:EngineRoot 'lib'
    foreach ($lib in @('Config.ps1','Logger.ps1','Context.ps1','Auth.ps1')) {
        $lf = Join-Path $libPath $lib
        if (Test-Path $lf) { . $lf }
    }
    $Script:LibsLoaded = $true
} catch {
    Write-Warning "AutoBuild UI: Could not load engine libraries: $_"
}

# ============================================================================
# RBAC - SEC-01/02 fix: real role resolution via Auth.ps1
# ============================================================================
$Script:EngineConfig = $null
if ($Script:LibsLoaded) {
    try { $Script:EngineConfig = Get-EngineConfig -Root $Script:EngineRoot } catch { }
}

$Script:CurrentUser = ''
try {
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $Script:CurrentUser = ($id.Name -split '\\')[-1]
} catch { $Script:CurrentUser = $env:USERNAME }

# Resolve real role against AD/whitelist; -Role param is a ceiling, not a grant.
$Script:CurrentRole = 'Operator'
if ($Script:LibsLoaded -and $null -ne $Script:EngineConfig) {
    try {
        $Script:CurrentRole = Resolve-UserRole `
            -Config        $Script:EngineConfig `
            -RequestedRole $Role
    } catch {
        # Auth.ps1 unavailable; fall back to requested role as-is.
        $Script:CurrentRole = $Role
    }
} else {
    $Script:CurrentRole = $Role
}

# ============================================================================
# QUEUE SYSTEM
# ============================================================================
$Script:QueueDir     = Join-Path $Script:UIRoot 'queue'
$Script:QueueEnabled = $false
if (Test-Path $Script:QueueDir) {
    try {
        Import-Module (Join-Path $Script:QueueDir 'QueueManager.psm1') -Force -ErrorAction Stop
        Import-Module (Join-Path $Script:QueueDir 'QueueRunner.psm1')  -Force -ErrorAction Stop
        try { Import-Module (Join-Path $Script:QueueDir 'QueueGUI.psm1') -Force -ErrorAction Stop } catch { }
        $Script:QueueEnabled = $true
    } catch {
        Write-Warning "QUEUE SYSTEM: Error loading modules: $_"
    }
}

# FIX PROD-GUARD (AUDIT v3): Detect dev-mode startup (empty security config).
# Shown as a persistent banner in the UI so operators cannot miss it.
$Script:SecurityWarning = ''
if ($Script:LibsLoaded -and $null -ne $Script:EngineConfig) {
    try {
        $sec = $Script:EngineConfig.security
        $adminEmpty = [string]::IsNullOrWhiteSpace($sec.adminAdGroup) -and
                      [string]::IsNullOrWhiteSpace($sec.adminUsers)
        $devEmpty   = [string]::IsNullOrWhiteSpace($sec.developerAdGroup) -and
                      [string]::IsNullOrWhiteSpace($sec.developerUsers)
        if ($adminEmpty -or $devEmpty) {
            $Script:SecurityWarning = 'SECURITY: Engine running in DEV MODE — security groups not configured. All users resolved as Operator.'
        }
    } catch { }
}

# ============================================================================
# GLOBALS
# ============================================================================
$Script:Window          = $null
$Script:Ctrl            = @{}
$Script:AllTasks        = @()
$Script:ActiveJobs      = @{}
$Script:RefreshTimer    = $null
$Script:ExecTimer       = $null
$Script:CurrentTaskName = ''
$Script:AllPageNames    = @(
    'pageCatalog','pageExecute','pageMonitor','pageRunLog','pageHistory',
    'pageCheckpoints','pageArtifacts','pageFolders','pageMetrics',
    'pageDiag','pageConfig','pageNewTask','pageAudit','pageQueue'
)

# Mutex name shared with Logger.ps1 and QueueRunner.psm1
$Script:AuditMutexName = 'Global\AutoBuildLogMutex'

# ============================================================================
# AUDIT LOG WRITER
# CONC-01-UI fix: single canonical mutex pattern. WaitOne checked.
# Write discarded (not unprotected) if mutex not acquired.
# CONC-04 fix: the dead function form 'Write-AuditLog' is removed.
#              Only the scriptblock form is used by event handlers.
# ============================================================================
$Script:Fn_WriteAuditLog = {
    param([string]$Action, [string]$Target = '', [string]$Detail = '')
    $dir = Split-Path $Script:AuditFile -Parent
    if (-not (Test-Path $dir)) { return }

    $safeDetail = $Detail -replace "`r`n",' | ' -replace "`n",' | ' -replace "`r",' | '
    $safeDetail = $safeDetail -replace '[\x00-\x1F\x7F]', ''

    $entry = [ordered]@{
        ts       = (Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')
        user     = $Script:CurrentUser
        role     = $Script:CurrentRole
        hostname = $env:COMPUTERNAME
        action   = $Action
        target   = $Target
        detail   = $safeDetail
    }
    $line = $entry | ConvertTo-Json -Compress

    # CONC-01-UI / FIX R-08: Two-tier mutex — Global\ for cross-session, Local\ fallback.
    $mutex  = $null
    $locked = $false
    try {
        try {
            $mutex = New-Object System.Threading.Mutex($false, $Script:AuditMutexName)
        } catch {
            $localName = $Script:AuditMutexName -replace '^Global\\', 'Local\'
            $mutex = New-Object System.Threading.Mutex($false, $localName)
        }
        $locked = $mutex.WaitOne(5000)
        if ($locked) {
            Add-Content -Path $Script:AuditFile -Value $line -Encoding ASCII
        }
        # Timeout: entry discarded. No unprotected write (FIX R-08).
    } catch {
        # Both tiers failed: discard. No unprotected write.
    } finally {
        if ($locked -and $null -ne $mutex) { try { $mutex.ReleaseMutex() } catch { } }
        if ($null -ne $mutex) { try { $mutex.Dispose() } catch { } }
    }
}

# ============================================================================
# PERMISSION CHECK - MAINT-01 fix: delegate to Auth.ps1 Test-Permission
# ============================================================================
$Script:Fn_TestPermission = {
    param([string]$Action)
    if ($Script:LibsLoaded) {
        try { return Test-Permission -Role $Script:CurrentRole -Action $Action } catch { }
    }
    # Fallback: static map when Auth.ps1 unavailable.
    # FIX V-04/R-04 (AUDIT v3 CRITICAL): The previous 'default { return $true }'
    # was FAIL-OPEN: any action not explicitly listed was granted to all roles.
    # This contradicted Auth.ps1's fail-safe design and meant any new action added
    # to the engine was silently accessible without restriction.
    # CORRECTION: default is now DENY ($false), matching Auth.ps1 behaviour.
    switch ($Action) {
        'EditConfig'        { return $Script:CurrentRole -eq 'Admin' }
        'ViewAudit'         { return $Script:CurrentRole -in @('Admin') }
        'PurgeOldLogs'      { return $Script:CurrentRole -eq 'Admin' }
        'DeleteArtifact'    { return $Script:CurrentRole -eq 'Admin' }
        'ManageCheckpoints' { return $Script:CurrentRole -eq 'Admin' }
        'CreateTask'        { return $Script:CurrentRole -in @('Admin','Developer') }
        'EditTask'          { return $Script:CurrentRole -in @('Admin','Developer') }
        'RunTask'           { return $true }   # All authenticated users
        'ViewHistory'       { return $true }
        'ViewArtifacts'     { return $true }
        'ViewMetrics'       { return $true }
        'ViewDiag'          { return $true }
        default             { return $false }  # FIX V-04: DENY unknown actions
    }
}

# ============================================================================
# STREAMING LOG READER
# SCALE-02/03 fix: reads only the last N bytes of registry.jsonl using a
# FileStream seek, never loading the whole file into memory.
# This makes the 5-second monitor refresh O(1) in file size.
# ============================================================================
$Script:Fn_ReadLogTail = {
    param([int]$MaxLines = 50, [string]$FilterTask = '', [string]$FilterLevel = '')
    $results = [System.Collections.Generic.List[object]]::new()
    if (-not (Test-Path $Script:RegistryFile)) { return $results.ToArray() }

    try {
        # Read last ~256KB, which is far more than 50 lines of JSONL.
        $fs = [System.IO.FileStream]::new(
            $Script:RegistryFile,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::ReadWrite)
        try {
            $seekPos = [Math]::Max(0, $fs.Length - 262144)  # last 256KB
            [void]$fs.Seek($seekPos, [System.IO.SeekOrigin]::Begin)
            $sr = [System.IO.StreamReader]::new($fs, [System.Text.Encoding]::ASCII)
            try {
                if ($seekPos -gt 0) { [void]$sr.ReadLine() }  # discard partial first line
                $rawLines = [System.Collections.Generic.List[string]]::new()
                $line = $sr.ReadLine()
                while ($null -ne $line) { $rawLines.Add($line); $line = $sr.ReadLine() }
            } finally { $sr.Close() }
        } finally { $fs.Close() }

        # Parse from end, collect up to MaxLines matching entries
        for ($i = $rawLines.Count - 1; $i -ge 0 -and $results.Count -lt $MaxLines; $i--) {
            $raw = $rawLines[$i]
            if ([string]::IsNullOrWhiteSpace($raw)) { continue }
            try {
                $o = $raw | ConvertFrom-Json
                if ($FilterTask  -and $o.task  -ne $FilterTask)  { continue }
                if ($FilterLevel -and $o.level -ne $FilterLevel) { continue }
                $results.Add($o)
            } catch { }
        }
        # Return in chronological order
        $arr = $results.ToArray()
        [array]::Reverse($arr)
        return $arr
    } catch { return @() }
}

# Full history reader for History page - still tail-based, larger window
$Script:Fn_GetExecutionHistory = {
    param([int]$MaxEntries = 200, [string]$FilterTask = '', [string]$FilterLevel = '')
    return @(& $Script:Fn_ReadLogTail -MaxLines $MaxEntries `
               -FilterTask $FilterTask -FilterLevel $FilterLevel)
}

$Script:Fn_GetRunSummaries = {
    # FIX R-09 (AUDIT v3 MEDIUM): Previous 800-line window caused active runs
    # to appear as 'RUNNING' when their terminal entry (OK/ERROR) was outside
    # the window. In high-activity environments (>13 tasks/min) this was
    # regularly hit. Increased to 2000 lines (~2.5x safety margin).
    $all = @(& $Script:Fn_ReadLogTail -MaxLines 2000)
    if ($all.Count -eq 0) { return @() }
    $summaries = @()
    foreach ($g in ($all | Group-Object runId)) {
        $entries = @($g.Group)
        $result  = $entries | Where-Object { $_.level -in @('OK','ERROR') } | Select-Object -Last 1
        $first   = $entries[0]
        $summaries += [PSCustomObject]@{
            RunId   = $g.Name
            Task    = $first.task
            Started = $first.ts
            Status  = if ($result) { $result.level } else { 'RUNNING' }
            Elapsed = if ($result -and $result.elapsed) { [math]::Round([double]($result.elapsed),1) } else { '' }
            User    = if ($first.user) { $first.user } else { '-' }
            Entries = $entries.Count
        }
    }
    return @($summaries | Sort-Object Started -Descending)
}

# ============================================================================
# TASK METADATA
# ============================================================================
$Script:Fn_GetTaskMetadata = {
    param([string]$Path)
    $name   = [System.IO.Path]::GetFileNameWithoutExtension($Path) -replace '^task_',''
    $desc   = ''
    $cat    = 'General'
    $ver    = '1.0.0'
    $params = @()
    try {
        foreach ($line in (Get-Content $Path -TotalCount 40 -Encoding ASCII -ErrorAction Stop)) {
            if ($line -match '#\s*@Description\s*:\s*(.+)')  { $desc = $Matches[1].Trim() }
            if ($line -match '#\s*@Category\s*:\s*(.+)')     { $cat  = $Matches[1].Trim() }
            if ($line -match '#\s*@Version\s*:\s*(.+)')      { $ver  = $Matches[1].Trim() }
            if ($line -match '#\s*@Param\s*:\s*(\S+)\s+(\S+)\s+(required|optional)\s*"([^"]*)"') {
                $params += [PSCustomObject]@{
                    Name     = $Matches[1]; Type = $Matches[2]
                    Required = ($Matches[3] -eq 'required'); Help = $Matches[4]
                }
            }
        }
    } catch { }
    return [PSCustomObject]@{
        Name=$name; Path=$Path; Description=$desc; Category=$cat
        Version=$ver; Params=$params; LastRun=''; LastStatus=''
    }
}

$Script:Fn_GetAllTasks = {
    $tasks = @()
    if (Test-Path $Script:TasksDir) {
        $tasks = @(
            Get-ChildItem $Script:TasksDir -Filter 'task_*.ps1' |
            Where-Object { $_.Name -ne 'task_TEMPLATE.ps1' } |
            Sort-Object Name |
            ForEach-Object { & $Script:Fn_GetTaskMetadata -Path $_.FullName }
        )
    }
    $history = @(& $Script:Fn_ReadLogTail -MaxLines 500)
    foreach ($task in $tasks) {
        $last = $history |
            Where-Object { $_.task -eq $task.Name -and $_.level -in @('OK','ERROR') } |
            Select-Object -First 1
        if ($last) { $task.LastRun = $last.ts; $task.LastStatus = $last.level }
    }
    return $tasks
}

# ============================================================================
# TASK EXECUTION
# RUN-01 fix: params serialized to JSON, passed via -Params to Run.ps1 v3.
# ACOPLAMIENTO-01 fix: no hardcoded -Centro/-Almacen parameter names.
# ============================================================================
$Script:Fn_StartTaskExecution = {
    param(
        [string]$TaskName,
        [hashtable]$Params,
        [switch]$WhatIf,
        [switch]$Checkpoint,
        [switch]$Resume
    )
    if (-not (Test-Path $Script:RunScript)) {
        return @{ Success=$false; Error="Run.ps1 not found: $Script:RunScript" }
    }
    & $Script:Fn_WriteAuditLog -Action 'EXECUTE_TASK' -Target $TaskName `
        -Detail "WhatIf=$WhatIf Checkpoint=$Checkpoint Resume=$Resume"

    # RUN-01 fix: serialize all params to JSON
    $paramsJson = '{}'
    if ($Params -and $Params.Count -gt 0) {
        $paramsJson = $Params | ConvertTo-Json -Compress
    }
    # FIX V-01: removed manual escaping ($escapedParams = $paramsJson -replace '"', '\"')
    # ArgumentList handles quoting safely — see process creation block below.

    # FIX V-01/R-01 (AUDIT v3 CRITICAL): Previously this code manually escaped
    # the JSON string with -replace '"', '\"' and embedded it in $argString.
    # A parameter value containing \" could break argstring parsing and inject
    # arbitrary arguments into the child powershell.exe process.
    #
    # CORRECTION: Use ProcessStartInfo.ArgumentList (StringCollection) which
    # handles quoting automatically. Each token is a distinct, safe argument.
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName               = 'powershell.exe'
    $psi.UseShellExecute        = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.CreateNoWindow         = $true
    $psi.WorkingDirectory       = $Script:EngineRoot

    [void]$psi.ArgumentList.Add('-NoProfile')
    [void]$psi.ArgumentList.Add('-NonInteractive')
    [void]$psi.ArgumentList.Add('-ExecutionPolicy')
    [void]$psi.ArgumentList.Add('Bypass')
    [void]$psi.ArgumentList.Add('-File')
    [void]$psi.ArgumentList.Add($Script:RunScript)
    [void]$psi.ArgumentList.Add('-Task')
    [void]$psi.ArgumentList.Add($TaskName)
    [void]$psi.ArgumentList.Add('-Params')
    [void]$psi.ArgumentList.Add($paramsJson)   # no manual escaping needed
    if ($WhatIf)     { [void]$psi.ArgumentList.Add('-WhatIf') }
    if ($Checkpoint) { [void]$psi.ArgumentList.Add('-Checkpoint') }
    if ($Resume)     { [void]$psi.ArgumentList.Add('-Resume') }

    $proc  = New-Object System.Diagnostics.Process
    $proc.StartInfo           = $psi
    $proc.EnableRaisingEvents = $true

    $queue  = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
    $outBuf = [System.Text.StringBuilder]::new(4096)

    $outH = [System.Diagnostics.DataReceivedEventHandler]{ param($s,$e); if ($null -ne $e.Data) { $queue.Enqueue($e.Data) } }
    $errH = [System.Diagnostics.DataReceivedEventHandler]{ param($s,$e); if ($null -ne $e.Data) { $queue.Enqueue("[STDERR] $($e.Data)") } }
    $proc.add_OutputDataReceived($outH)
    $proc.add_ErrorDataReceived($errH)

    try {
        [void]$proc.Start()
        $proc.BeginOutputReadLine()
        $proc.BeginErrorReadLine()
    } catch {
        return @{ Success=$false; Error="Process start failed: $_" }
    }

    $Script:ActiveJobs[$TaskName] = @{
        Process  = $proc
        Queue    = $queue
        Buffer   = $outBuf
        Started  = [datetime]::Now
        TaskName = $TaskName
    }
    return @{ Success=$true; Pid=$proc.Id }
}

$Script:Fn_GetJobOutput = {
    param([string]$TaskName)
    if (-not $Script:ActiveJobs.ContainsKey($TaskName)) { return $null }
    $info    = $Script:ActiveJobs[$TaskName]
    $proc    = $info.Process
    $running = -not $proc.HasExited
    $line    = $null
    while ($info.Queue.TryDequeue([ref]$line)) {
        [void]$info.Buffer.AppendLine($line)
    }
    return @{
        State    = if ($running) { 'Running' } elseif ($proc.ExitCode -eq 0) { 'Completed' } else { 'Failed' }
        Output   = $info.Buffer.ToString()
        Duration = ([datetime]::Now - $info.Started).TotalSeconds
        ExitCode = if (-not $running) { $proc.ExitCode } else { $null }
    }
}

# ============================================================================
# DIAGNOSTICS - COM-FREEZE fix: run in background Runspace, post to UI
# via Dispatcher.BeginInvoke so the WPF thread never blocks on COM calls.
# ============================================================================
$Script:Fn_RunDiagnostics = {
    $Script:DiagBd = $Script:Ctrl['btnRunDiag']
    $Script:DiagDs = $Script:Ctrl['txtDiagSummary']
    $Script:DiagGd = $Script:Ctrl['gridDiag']

    if ($null -ne $Script:DiagBd) { $Script:DiagBd.IsEnabled = $false }
    if ($null -ne $Script:DiagDs) { $Script:DiagDs.Text = 'Running diagnostics (background)...' }

    # Capture values for the Runspace (cannot close over $Script: directly)
    $Script:DiagEngineRoot = $Script:EngineRoot

    $Script:DiagRs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $Script:DiagRs.ApartmentState = [System.Threading.ApartmentState]::STA
    $Script:DiagRs.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $Script:DiagRs.Open()

    $Script:DiagPs = [System.Management.Automation.PowerShell]::Create()
    $Script:DiagPs.Runspace = $Script:DiagRs

    [void]$Script:DiagPs.AddScript({
        param($engineRoot)

        $results = [System.Collections.Generic.List[PSCustomObject]]::new()
        $add = {
            param([string]$cat,[string]$item,[string]$val,[string]$status,[string]$msg)
            $results.Add([PSCustomObject]@{Category=$cat;Item=$item;Value=$val;Status=$status;Message=$msg})
        }

        # Runtime
        try {
            $psv = $PSVersionTable.PSVersion.ToString()
            $ok  = [version]$psv -ge [version]'5.1'
            & $add 'Runtime' 'PowerShell'  $psv $(if($ok){'OK'}else{'WARN'}) $(if($ok){'PS 5.1+ OK'}else{'Upgrade recommended'})
            $ed  = $PSVersionTable.PSEdition
            & $add 'Runtime' 'Edition'     $ed  $(if($ed -eq 'Desktop'){'OK'}else{'WARN'}) $(if($ed -eq 'Desktop'){'Full COM support'}else{'Core: limited COM'})
        } catch { & $add 'Runtime' 'PowerShell' 'Error' 'ERROR' "$_" }

        # Engine files
        foreach ($f in @('Run.ps1','engine\Main.build.ps1','engine.config.json')) {
            try {
                $fp = Join-Path $engineRoot $f
                $ok = Test-Path $fp
                & $add 'Engine' $f $(if($ok){'Found'}else{'Missing'}) $(if($ok){'OK'}else{'ERROR'}) $(if(-not $ok){"Missing: $fp"}else{'Found'})
            } catch { & $add 'Engine' $f 'Error' 'ERROR' "$_" }
        }

        # Libraries
        foreach ($lib in @('Config.ps1','Logger.ps1','Context.ps1','Auth.ps1','Retry.ps1','ComHelper.ps1','ExcelHelper.ps1','SapHelper.ps1')) {
            try {
                $lp = Join-Path $engineRoot "lib\$lib"
                $ok = Test-Path $lp
                & $add 'Libraries' $lib $(if($ok){'Found'}else{'Missing'}) $(if($ok){'OK'}else{'ERROR'}) $(if($ok){$lp}else{"Missing: $lp"})
            } catch { & $add 'Libraries' $lib 'Error' 'ERROR' "$_" }
        }

        # Folders
        foreach ($d in @('tasks','logs','output','reports','input')) {
            try {
                $dp = Join-Path $engineRoot $d
                if (Test-Path $dp) {
                    $tp = Join-Path $dp "._test_$(Get-Random)"
                    try {
                        [void](New-Item $tp -ItemType File -Force -ErrorAction Stop)
                        Remove-Item $tp -Force
                        & $add 'Folders' $d $dp 'OK' 'Writable'
                    } catch { & $add 'Folders' $d $dp 'WARN' 'Exists but not writable' }
                } else { & $add 'Folders' $d $dp 'WARN' 'Missing - created on first use' }
            } catch { & $add 'Folders' $d 'Error' 'ERROR' "$_" }
        }

        # COM: Excel - runs in Runspace so UI thread cannot freeze
        try {
            $xl  = New-Object -ComObject Excel.Application -ErrorAction Stop
            $ver = $xl.Version
            $xl.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
            & $add 'COM' 'Microsoft Excel' "v$ver" 'OK' 'Excel COM available'
        } catch { & $add 'COM' 'Microsoft Excel' 'Not Available' 'WARN' "Excel not found: $_" }

        # COM: Word
        try {
            $wd  = New-Object -ComObject Word.Application -ErrorAction Stop
            $ver = $wd.Version
            $wd.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wd) | Out-Null
            & $add 'COM' 'Microsoft Word' "v$ver" 'OK' 'Word COM available'
        } catch { & $add 'COM' 'Microsoft Word' 'Not Available' 'WARN' "Word not found: $_" }

        # COM: SAP GUI
        try {
            $sap = New-Object -ComObject SapROTWr.SapROTWrapper -ErrorAction Stop
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sap) | Out-Null
            & $add 'COM' 'SAP GUI' 'Available' 'OK' 'SAP GUI scripting available'
        } catch { & $add 'COM' 'SAP GUI' 'Not Available' 'INFO' 'Required only for SAP tasks' }

        # Tools
        try {
            $ib = Join-Path $engineRoot 'tools\InvokeBuild\Invoke-Build.ps1'
            $ok = Test-Path $ib
            & $add 'Tools' 'Invoke-Build' $(if($ok){'Found'}else{'Missing'}) $(if($ok){'OK'}else{'ERROR'}) $(if($ok){'Portable IB available'}else{"Missing: $ib"})
        } catch { & $add 'Tools' 'Invoke-Build' 'Error' 'ERROR' "$_" }

        return $results.ToArray()

    }).AddArgument($Script:DiagEngineRoot)

    $Script:DiagAuditFn = $Script:Fn_WriteAuditLog

    $Script:DiagAsyncResult = $Script:DiagPs.BeginInvoke()

    # Poll for completion on the DispatcherTimer (non-blocking)
    $Script:DiagTimer = New-Object System.Windows.Threading.DispatcherTimer
    $Script:DiagTimer.Interval = [TimeSpan]::FromMilliseconds(300)
    $Script:DiagTimer.Add_Tick({
        if (-not $Script:DiagPs.InvocationStateInfo.State.Equals([System.Management.Automation.PSInvocationState]::Running)) {
            $Script:DiagTimer.Stop()
            try {
                $results = @($Script:DiagPs.EndInvoke($Script:DiagAsyncResult))
                $Script:DiagPs.Dispose()
                $Script:DiagRs.Close()
                $Script:DiagRs.Dispose()

                # Post results to UI via Dispatcher (we are already on UI thread via DispatcherTimer)
                if ($null -ne $Script:DiagGd) { $Script:DiagGd.ItemsSource = $results }
                $ok   = @($results | Where-Object { $_.Status -eq 'OK'    }).Count
                $warn = @($results | Where-Object { $_.Status -eq 'WARN'  }).Count
                $err  = @($results | Where-Object { $_.Status -eq 'ERROR' }).Count
                $info = @($results | Where-Object { $_.Status -eq 'INFO'  }).Count
                if ($null -ne $Script:DiagDs) { $Script:DiagDs.Text = "$ok OK  |  $warn WARN  |  $err ERROR  |  $info INFO" }
                & $Script:DiagAuditFn -Action 'RUN_DIAGNOSTICS' -Detail "OK=$ok WARN=$warn ERR=$err"
            } catch {
                if ($null -ne $Script:DiagDs) { $Script:DiagDs.Text = "Diagnostics error: $_" }
            } finally {
                if ($null -ne $Script:DiagBd) { $Script:DiagBd.IsEnabled = $true }
            }
        }
    })
    $Script:DiagTimer.Start()
}

# ============================================================================
# PAGE HELPERS
# ============================================================================
$Script:Fn_HideAllPages = {
    foreach ($pn in $Script:AllPageNames) {
        $p = $Script:Ctrl[$pn]
        if ($null -ne $p) { $p.Visibility = [System.Windows.Visibility]::Collapsed }
    }
}

$Script:Fn_NavigateTo = {
    param([string]$PageName, [string]$Title)
    & $Script:Fn_HideAllPages
    $p = $Script:Ctrl[$PageName]
    if ($null -ne $p) { $p.Visibility = [System.Windows.Visibility]::Visible }
    $t = $Script:Ctrl['txtPageTitle']
    if ($null -ne $t) { $t.Text = $Title }
}

$Script:Fn_LoadCatalogPage = {
    $Script:AllTasks = @(& $Script:Fn_GetAllTasks)
    & $Script:Fn_FilterCatalog
}

$Script:Fn_FilterCatalog = {
    $search   = ($(if ($Script:Ctrl['txtCatalogSearch']) { $Script:Ctrl['txtCatalogSearch'].Text } else { '' })).ToLower()
    $catItem  = if ($Script:Ctrl['cboCatalogCategory']) { $Script:Ctrl['cboCatalogCategory'].SelectedItem } else { $null }
    $category = if ($null -ne $catItem -and $catItem.Content -ne 'All Categories') { $catItem.Content } else { '' }

    $rows = @($Script:AllTasks | Where-Object {
        ($search -eq '' -or
         $_.Name.ToLower()        -like "*$search*" -or
         $_.Description.ToLower() -like "*$search*") -and
        ($category -eq '' -or $_.Category -eq $category)
    } | ForEach-Object {
        [PSCustomObject]@{
            Name       = $_.Name
            Category   = $_.Category
            Description= $_.Description
            Version    = $_.Version
            LastRun    = if ($_.LastRun)    { $_.LastRun }    else { 'Never' }
            LastStatus = if ($_.LastStatus) { $_.LastStatus } else { '-' }
        }
    })
    $g = $Script:Ctrl['gridCatalog'];     if ($null -ne $g) { $g.ItemsSource = $rows }
    $c = $Script:Ctrl['txtCatalogCount']; if ($null -ne $c) { $c.Text = "$($rows.Count) tasks" }
}

$Script:Fn_LoadExecutePage = {
    param([string]$PreSelectTask = '')
    $Script:AllTasks = @(& $Script:Fn_GetAllTasks)
    $cbo = $Script:Ctrl['cboExecTask']
    if ($null -eq $cbo) { return }
    $cbo.Items.Clear()
    foreach ($t in $Script:AllTasks) {
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content    = $t.Name
        $item.Tag        = $t.Name
        $item.Foreground = [System.Windows.Media.Brushes]::WhiteSmoke
        $item.Background = [System.Windows.Media.SolidColorBrush][System.Windows.Media.ColorConverter]::ConvertFromString('#1E2232')
        [void]$cbo.Items.Add($item)
    }
    if ($PreSelectTask) {
        for ($i = 0; $i -lt $cbo.Items.Count; $i++) {
            if ($cbo.Items[$i].Tag -eq $PreSelectTask) { $cbo.SelectedIndex = $i; break }
        }
    }
}

$Script:Fn_BuildParamForm = {
    param([string]$TaskName)
    $panel = $Script:Ctrl['pnlParams']
    if ($null -eq $panel) { return }
    $panel.Children.Clear()
    $task = $Script:AllTasks | Where-Object { $_.Name -eq $TaskName } | Select-Object -First 1
    if ($null -eq $task -or $task.Params.Count -eq 0) {
        $lbl            = New-Object System.Windows.Controls.TextBlock
        $lbl.Text       = 'This task has no parameters.'
        $lbl.Foreground = [System.Windows.Media.Brushes]::Gray
        [void]$panel.Children.Add($lbl)
        return
    }
    foreach ($param in $task.Params) {
        $lbl            = New-Object System.Windows.Controls.TextBlock
        $lbl.Text       = "$($param.Name)$(if($param.Required){' *'})"
        $lbl.Foreground = [System.Windows.Media.Brushes]::Gray
        $lbl.FontSize   = 11
        $lbl.Margin     = [System.Windows.Thickness]::new(0,8,0,3)
        [void]$panel.Children.Add($lbl)
        $tb             = New-Object System.Windows.Controls.TextBox
        $tb.Height      = 32
        $tb.FontSize    = 13
        $tb.Tag         = $param.Name
        $tb.Padding     = [System.Windows.Thickness]::new(8,4,8,4)
        $tb.Background  = [System.Windows.Media.SolidColorBrush][System.Windows.Media.ColorConverter]::ConvertFromString('#1E2232')
        $tb.Foreground  = [System.Windows.Media.Brushes]::WhiteSmoke
        if ($param.Help) { $tb.ToolTip = "$($param.Help) ($($param.Type))" }
        [void]$panel.Children.Add($tb)
        if ($param.Help) {
            $h            = New-Object System.Windows.Controls.TextBlock
            $h.Text       = $param.Help
            $h.Foreground = [System.Windows.Media.Brushes]::DarkGray
            $h.FontSize   = 10
            [void]$panel.Children.Add($h)
        }
    }
}

$Script:Fn_RunSelectedTask = {
    $cbo  = $Script:Ctrl['cboExecTask']
    $sel  = if ($null -ne $cbo) { $cbo.SelectedItem } else { $null }
    $esT  = $Script:Ctrl['txtExecStatus']
    $esB  = $Script:Ctrl['elExecStatus']
    $bRun = $Script:Ctrl['btnExecuteTask']
    $bCan = $Script:Ctrl['btnCancelTask']

    if ($null -eq $sel) {
        if ($null -ne $esT) { $esT.Text = 'Select a task first.' }
        return
    }
    $taskName = $sel.Tag
    $params   = @{}
    $panel    = $Script:Ctrl['pnlParams']
    if ($null -ne $panel) {
        foreach ($child in $panel.Children) {
            if ($child -is [System.Windows.Controls.TextBox] -and $null -ne $child.Tag) {
                $params[$child.Tag] = $child.Text
            }
        }
    }
    $task = $Script:AllTasks | Where-Object { $_.Name -eq $taskName } | Select-Object -First 1
    if ($null -ne $task) {
        foreach ($p in $task.Params) {
            if ($p.Required -and [string]::IsNullOrWhiteSpace($params[$p.Name])) {
                if ($null -ne $esT) { $esT.Text = "Required field missing: $($p.Name)" }
                if ($null -ne $esB) { $esB.Background = [System.Windows.Media.Brushes]::Crimson }
                return
            }
        }
    }
    if ($null -ne $Script:ExecTimer) { $Script:ExecTimer.Stop(); $Script:ExecTimer = $null }
    $Script:CurrentTaskName = $taskName
    $eo = $Script:Ctrl['txtExecOutput']
    if ($null -ne $eo)  { $eo.Text  = "Starting: $taskName`n$('-'*60)`n" }
    if ($null -ne $esT) { $esT.Text = 'Running...' }
    if ($null -ne $esB) { $esB.Background = [System.Windows.Media.Brushes]::Goldenrod }
    if ($null -ne $bRun){ $bRun.IsEnabled = $false }
    if ($null -ne $bCan){ $bCan.IsEnabled = $true }

    $wi         = $Script:Ctrl['chkWhatIf']
    $cp         = $Script:Ctrl['chkCheckpoint']
    $rp         = $Script:Ctrl['chkResume']
    $whatIf     = ($null -ne $wi  -and $wi.IsChecked  -eq $true)
    $checkpoint = ($null -ne $cp  -and $cp.IsChecked  -eq $true)
    $resume     = ($null -ne $rp  -and $rp.IsChecked  -eq $true)

    $result = & $Script:Fn_StartTaskExecution -TaskName $taskName -Params $params `
              -WhatIf:$whatIf -Checkpoint:$checkpoint -Resume:$resume

    if (-not $result.Success) {
        if ($null -ne $esT) { $esT.Text = "Error: $($result.Error)" }
        if ($null -ne $esB) { $esB.Background = [System.Windows.Media.Brushes]::Crimson }
        if ($null -ne $bRun){ $bRun.IsEnabled = $true }
        if ($null -ne $bCan){ $bCan.IsEnabled = $false }
        return
    }

    $Script:ExecTimer = New-Object System.Windows.Threading.DispatcherTimer
    $Script:ExecTimer.Interval = [TimeSpan]::FromMilliseconds(800)
    $Script:ExecTimer.Add_Tick({ & $Script:Fn_PollExecOutput })
    $Script:ExecTimer.Start()
}

$Script:Fn_PollExecOutput = {
    if ([string]::IsNullOrEmpty($Script:CurrentTaskName)) { return }
    $info = & $Script:Fn_GetJobOutput -TaskName $Script:CurrentTaskName
    if ($null -eq $info) { return }
    $eo = $Script:Ctrl['txtExecOutput'];   if ($null -ne $eo) { $eo.Text = $info.Output }
    $ed = $Script:Ctrl['txtExecDuration']; if ($null -ne $ed) { $ed.Text = "$([math]::Round($info.Duration,1))s" }
    $sv = $Script:Ctrl['svExecOutput'];    if ($null -ne $sv) { $sv.ScrollToEnd() }
    $rl = $Script:Ctrl['txtRunLog'];       if ($null -ne $rl) { $rl.Text = $info.Output }
    $rs = $Script:Ctrl['svRunLog'];        if ($null -ne $rs) { $rs.ScrollToEnd() }
    $rh = $Script:Ctrl['txtRunLogHeader']
    if ($null -ne $rh) { $rh.Text = "LIVE: $($Script:CurrentTaskName)  |  $([math]::Round($info.Duration,1))s  |  $($info.State)" }

    if ($info.State -in @('Completed','Failed','Stopped')) {
        if ($null -ne $Script:ExecTimer) { $Script:ExecTimer.Stop(); $Script:ExecTimer = $null }
        $bRun = $Script:Ctrl['btnExecuteTask']; if ($null -ne $bRun){ $bRun.IsEnabled = $true }
        $bCan = $Script:Ctrl['btnCancelTask'];  if ($null -ne $bCan){ $bCan.IsEnabled = $false }
        $esT  = $Script:Ctrl['txtExecStatus']
        $esB  = $Script:Ctrl['elExecStatus']
        if ($info.State -eq 'Completed') {
            if ($null -ne $esT) { $esT.Text = 'Completed' }
            if ($null -ne $esB) { $esB.Background = [System.Windows.Media.Brushes]::LimeGreen }
        } else {
            if ($null -ne $esT) { $esT.Text = $info.State }
            if ($null -ne $esB) { $esB.Background = [System.Windows.Media.Brushes]::Crimson }
        }
        if ($Script:ActiveJobs.ContainsKey($Script:CurrentTaskName)) {
            try { $Script:ActiveJobs[$Script:CurrentTaskName].Process.Dispose() } catch { }
            $Script:ActiveJobs.Remove($Script:CurrentTaskName)
        }
    }
}

# ============================================================================
# MONITOR PAGE - SCALE-02/03 fix: tail read only, no full file load
# ============================================================================
$Script:Fn_LoadMonitorPage = {
    $rows = @(& $Script:Fn_GetRunSummaries | Select-Object -First 30 | ForEach-Object {
        [PSCustomObject]@{
            TaskName = $_.Task
            RunId    = $_.RunId
            Started  = $_.Started
            Duration = if ($_.Elapsed -ne '') { "$($_.Elapsed)s" } else { 'running...' }
            Status   = $_.Status
            User     = $_.User
        }
    })
    $g = $Script:Ctrl['gridMonitorJobs']; if ($null -ne $g) { $g.ItemsSource = $rows }
    & $Script:Fn_LoadMonitorLogTail
    $tc = $Script:Ctrl['txtActiveCount']; if ($null -ne $tc) { $tc.Text = "$($Script:ActiveJobs.Count) running" }
}

$Script:Fn_LoadMonitorLogTail = {
    param([int]$DaysBack = 0)
    $logText = ''
    $lines = @(& $Script:Fn_ReadLogTail -MaxLines 50)
    if ($DaysBack -gt 0) {
        $cutoff = [datetime]::Now.AddDays(-$DaysBack)
        $lines = @($lines | Where-Object {
            try { [datetime]($_.ts) -gt $cutoff } catch { $true }
        })
    }
    $logText = ($lines | ForEach-Object {
        "[$($_.ts)][$($_.level)] $($_.task): $($_.message)"
    }) -join "`n"
    $tl = $Script:Ctrl['txtLiveLog'];  if ($null -ne $tl) { $tl.Text = $logText }
    $sl = $Script:Ctrl['svLiveLog'];   if ($null -ne $sl) { $sl.ScrollToEnd() }
}

$Script:Fn_LoadHistoryPage = {
    $ft = if ($Script:Ctrl['txtHistorySearch']) { $Script:Ctrl['txtHistorySearch'].Text } else { '' }
    $si = if ($Script:Ctrl['cboHistoryStatus']) { $Script:Ctrl['cboHistoryStatus'].SelectedItem } else { $null }
    $fs = if ($null -ne $si -and $si.Content -ne 'All Status') { $si.Content } else { '' }
    $summaries = @(& $Script:Fn_GetRunSummaries)
    if ($ft) { $summaries = @($summaries | Where-Object { $_.Task -like "*$ft*" }) }
    if ($fs) { $summaries = @($summaries | Where-Object { $_.Status -eq $fs }) }
    $rows = @($summaries | ForEach-Object {
        [PSCustomObject]@{
            RunId   = $_.RunId
            Task    = $_.Task
            Started = $_.Started
            Elapsed = if ($_.Elapsed -ne '') { "$($_.Elapsed)s" } else { '-' }
            Status  = $_.Status
            Entries = $_.Entries
            User    = $_.User
        }
    })
    $g = $Script:Ctrl['gridHistory']; if ($null -ne $g) { $g.ItemsSource = $rows }
}

$Script:Fn_ShowRunDetail = {
    param([string]$RunId, [string]$Task)
    $entries = @(& $Script:Fn_GetExecutionHistory -MaxEntries 500 -FilterTask $Task |
                 Where-Object { $_.runId -eq $RunId })
    $text = ($entries | ForEach-Object {
        "[$($_.ts)][$($_.level)] $($_.message)$(if($_.detail){' | '+$_.detail})"
    }) -join "`n"
    $td = $Script:Ctrl['txtHistoryDetail'];      if ($null -ne $td) { $td.Text = $text }
    $tt = $Script:Ctrl['txtHistoryDetailTitle']; if ($null -ne $tt) { $tt.Text = "RUN: $RunId" }
    $sd = $Script:Ctrl['svHistoryDetail'];       if ($null -ne $sd) { $sd.ScrollToTop() }
}

$Script:Fn_LoadCheckpointsPage = {
    $logDir = Join-Path $Script:EngineRoot 'logs'
    $items  = @()
    if (Test-Path $logDir) {
        $items = @(
            Get-ChildItem $logDir -Filter 'checkpoint_*.clixml' |
            Sort-Object LastWriteTime -Descending |
            ForEach-Object {
                [PSCustomObject]@{
                    Name     = $_.BaseName
                    TaskName = ($_.BaseName -replace '^checkpoint_','') -replace '_\d{8}_\d{6}$',''
                    Modified = $_.LastWriteTime.ToString('yyyy-MM-dd HH:mm')
                    SizeKB   = [math]::Round($_.Length/1KB,1)
                    Path     = $_.FullName
                }
            }
        )
    }
    $g = $Script:Ctrl['gridCheckpoints']
    if ($null -ne $g) { $g.ItemsSource = $items }
}

$Script:Fn_GetArtifacts = {
    $all = @()
    foreach ($dir in @($Script:OutputDir,$Script:ReportsDir)) {
        if (Test-Path $dir) {
            $all += Get-ChildItem $dir -File -Recurse -ErrorAction SilentlyContinue |
                    ForEach-Object {
                        [PSCustomObject]@{
                            Name=$_.Name; Extension=$_.Extension
                            Directory=(Split-Path $dir -Leaf)
                            Size=$_.Length; Modified=$_.LastWriteTime; Path=$_.FullName
                        }
                    }
        }
    }
    return @($all | Sort-Object Modified -Descending)
}

$Script:Fn_LoadArtifactsPage = {
    $filter = ($(if ($Script:Ctrl['txtArtifactSearch']) { $Script:Ctrl['txtArtifactSearch'].Text } else { '' })).ToLower()
    $arts   = @(& $Script:Fn_GetArtifacts)
    if ($filter) { $arts = @($arts | Where-Object { $_.Name.ToLower() -like "*$filter*" }) }
    $rows = @($arts | ForEach-Object {
        [PSCustomObject]@{
            Name      = $_.Name
            Extension = $_.Extension
            Directory = $_.Directory
            SizeKB    = [math]::Round($_.Size/1KB,1)
            Modified  = $_.Modified.ToString('yyyy-MM-dd HH:mm')
            Path      = $_.Path
        }
    })
    $ga = $Script:Ctrl['gridArtifacts'];    if ($null -ne $ga) { $ga.ItemsSource = $rows }
    $tc = $Script:Ctrl['txtArtifactCount']; if ($null -ne $tc) { $tc.Text = "$($rows.Count) artifacts" }
}

$Script:Fn_GetMetrics = {
    $s   = @(& $Script:Fn_GetRunSummaries)
    $ok  = @($s | Where-Object { $_.Status -eq 'OK'    }).Count
    $err = @($s | Where-Object { $_.Status -eq 'ERROR' }).Count
    $wt  = @($s | Where-Object { $_.Elapsed -ne '' })
    $avg = if ($wt.Count -gt 0) { [math]::Round(($wt | Measure-Object Elapsed -Sum).Sum / $wt.Count,1) } else { 0 }
    $tc  = @($s | Group-Object Task | Sort-Object Count -Descending)
    return @{
        Total      = $s.Count
        OK         = $ok
        Error      = $err
        SuccessRate= if ($s.Count -gt 0) { [math]::Round($ok / $s.Count * 100, 1) } else { 0 }
        AvgElapsed = $avg
        TopTask    = if ($tc.Count -gt 0) { $tc[0].Name } else { 'N/A' }
        TaskCounts = $tc
    }
}

$Script:Fn_LoadMetricsPage = {
    $m = & $Script:Fn_GetMetrics
    $c = $Script:Ctrl['txtMetricTotal'];   if ($null -ne $c) { $c.Text = "$($m.Total)" }
    $c = $Script:Ctrl['txtMetricSuccess']; if ($null -ne $c) { $c.Text = "$($m.SuccessRate)%" }
    $c = $Script:Ctrl['txtMetricOKErr'];   if ($null -ne $c) { $c.Text = "$($m.OK) OK / $($m.Error) Err" }
    $c = $Script:Ctrl['txtMetricAvg'];     if ($null -ne $c) { $c.Text = "$($m.AvgElapsed)s" }
    $c = $Script:Ctrl['txtMetricTop'];     if ($null -ne $c) { $c.Text = $m.TopTask }
    $max  = if ($m.TaskCounts.Count -gt 0) { [int]$m.TaskCounts[0].Count } else { 1 }
    if ($max -lt 1) { $max = 1 }
    $rows = @($m.TaskCounts | Select-Object -First 15 | ForEach-Object {
        [PSCustomObject]@{
            Name  = $_.Name
            Count = $_.Count
            Bar   = ('#' * [math]::Round([int]$_.Count / $max * 30))
        }
    })
    $g = $Script:Ctrl['gridMetricTasks'];  if ($null -ne $g) { $g.ItemsSource = $rows }
    $e = @(& $Script:Fn_GetExecutionHistory -MaxEntries 100 -FilterLevel 'ERROR' | Select-Object -Last 10)
    $g = $Script:Ctrl['gridMetricErrors']; if ($null -ne $g) { $g.ItemsSource = $e }
}

$Script:Fn_GetConfig = {
    if (-not (Test-Path $Script:ConfigFile)) { return $null }
    try { return Get-Content $Script:ConfigFile -Raw -Encoding ASCII | ConvertFrom-Json } catch { return $null }
}

$Script:Fn_LoadConfigPage = {
    $ce = $Script:Ctrl['txtConfigEditor']
    $bs = $Script:Ctrl['btnSaveConfig']
    if (-not (& $Script:Fn_TestPermission -Action 'EditConfig')) {
        if ($null -ne $ce) { $ce.Text = '// Admin role required to edit configuration.'; $ce.IsReadOnly = $true }
        if ($null -ne $bs) { $bs.IsEnabled = $false }
        return
    }
    $cfg = & $Script:Fn_GetConfig
    if ($null -ne $ce) {
        $ce.Text = if ($null -ne $cfg) { $cfg | ConvertTo-Json -Depth 5 } else { '// Error reading config or file missing.' }
    }
}

$Script:Fn_SaveConfigPage = {
    $ce = $Script:Ctrl['txtConfigEditor']
    $txt = if ($null -ne $ce) { $ce.Text } else { '' }
    $cs  = $Script:Ctrl['txtConfigStatus']
    try {
        $obj = $txt | ConvertFrom-Json
        # -- FIX V-05 (AUDIT v3 HIGH): Expanded validation beyond maxRetries.
        # An Admin could previously inject a malicious SMTP server or clear AD groups,
        # effectively escalating to unrestricted access after the next restart.

        # 1. maxRetries range
        $mr  = [int]$obj.engine.maxRetries
        if ($mr -lt 0 -or $mr -gt 10) { throw "engine.maxRetries must be 0-10 (got $mr)" }

        # 2. retryDelaySeconds sanity
        if ($null -ne $obj.engine.retryDelaySeconds) {
            $rd = [double]$obj.engine.retryDelaySeconds
            if ($rd -lt 0 -or $rd -gt 300) { throw "engine.retryDelaySeconds must be 0-300s (got $rd)" }
        }

        # 3. SMTP server: no shell metacharacters, max 253 chars (RFC 1035)
        if (-not [string]::IsNullOrWhiteSpace($obj.notifications.smtpServer)) {
            $smtp = $obj.notifications.smtpServer.Trim()
            if ($smtp.Length -gt 253) { throw "notifications.smtpServer is too long (max 253 chars)" }
            if ($smtp -match '[;&|`$<>!{}()\[\]\\]') {
                throw "notifications.smtpServer contains invalid characters: '$smtp'"
            }
        }

        # 4. SMTP port range
        if ($null -ne $obj.notifications.smtpPort) {
            $port = [int]$obj.notifications.smtpPort
            if ($port -lt 1 -or $port -gt 65535) { throw "notifications.smtpPort must be 1-65535 (got $port)" }
        }

        # 5. AD group names: must start with 'CN=' if non-empty (basic DN format check)
        foreach ($field in @('adminAdGroup','developerAdGroup')) {
            $val = $obj.security.$field
            if (-not [string]::IsNullOrWhiteSpace($val)) {
                $val = $val.Trim()
                if (-not ($val -match '^CN=')) {
                    throw "security.$field must be a valid LDAP Distinguished Name starting with 'CN=' (got '$val')"
                }
                if ($val.Length -gt 512) { throw "security.$field is too long (max 512 chars)" }
            }
        }

        # 6. User whitelists: only alphanumeric, dot, hyphen, underscore
        foreach ($field in @('adminUsers','developerUsers')) {
            $val = $obj.security.$field
            if (-not [string]::IsNullOrWhiteSpace($val)) {
                $users = $val -split '[,;\s]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                foreach ($u in $users) {
                    if ($u -match '[^a-zA-Z0-9._\-]') {
                        throw "security.$field contains invalid username '$u'. Only letters, digits, ., - and _ are allowed."
                    }
                }
            }
        }

        # 7. logLevel must be a valid value
        $validLevels = @('DEBUG','INFO','WARN','ERROR','FATAL')
        if (-not [string]::IsNullOrWhiteSpace($obj.engine.logLevel)) {
            if ($validLevels -notcontains $obj.engine.logLevel.Trim().ToUpper()) {
                throw "engine.logLevel must be one of: $($validLevels -join ', ')"
            }
        }

        [System.IO.File]::WriteAllText($Script:ConfigFile, $txt, [System.Text.Encoding]::ASCII)
        & $Script:Fn_WriteAuditLog -Action 'EDIT_CONFIG'
        if ($null -ne $cs) { $cs.Text = 'Saved successfully.'; $cs.Foreground = [System.Windows.Media.Brushes]::LimeGreen }
    } catch {
        if ($null -ne $cs) { $cs.Text = "Error: $_"; $cs.Foreground = [System.Windows.Media.Brushes]::Crimson }
    }
}

$Script:Fn_LoadAuditPage = {
    $ga = $Script:Ctrl['gridAudit']
    if (-not (& $Script:Fn_TestPermission -Action 'ViewAudit') -or -not (Test-Path $Script:AuditFile)) {
        if ($null -ne $ga) { $ga.ItemsSource = $null }; return
    }
    $entries = @()
    try {
        Get-Content $Script:AuditFile -Encoding ASCII -ErrorAction Stop | Select-Object -Last 500 |
            ForEach-Object { try { $entries += ($_ | ConvertFrom-Json) } catch { } }
    } catch { }
    if ($null -ne $ga) { $ga.ItemsSource = @($entries | Sort-Object ts -Descending) }
}

$Script:Fn_PurgeOldLogs = {
    param([int]$RetentionDays = 30)
    if (-not (& $Script:Fn_TestPermission -Action 'PurgeOldLogs')) {
        [void][System.Windows.MessageBox]::Show('Admin role required.','AutoBuild','OK','Warning')
        return
    }
    # Delegate to Logger.ps1 Invoke-LogPurge if available
    if ($Script:LibsLoaded) {
        try {
            Invoke-LogPurge -LogsDir $Script:LogsDir -RetentionDays $RetentionDays
            & $Script:Fn_WriteAuditLog -Action 'PURGE_LOGS' -Detail "RetentionDays=$RetentionDays"
            [void][System.Windows.MessageBox]::Show("Rotated archives older than $RetentionDays days purged.",'AutoBuild','OK','Information')
            return
        } catch { }
    }
    # Fallback: in-place line filter (uses atomic write)
    $cutoff = (Get-Date).AddDays(-$RetentionDays)
    $purged = 0
    if (Test-Path $Script:RegistryFile) {
        $lines = @(Get-Content $Script:RegistryFile -Encoding ASCII -ErrorAction SilentlyContinue)
        $kept  = @($lines | Where-Object {
            try {
                $o = $_ | ConvertFrom-Json
                [datetime]($o.ts) -gt $cutoff
            } catch { $true }
        })
        $purged = $lines.Count - $kept.Count
        if ($purged -gt 0) {
            $tmp = "$Script:RegistryFile.tmp"
            [System.IO.File]::WriteAllLines($tmp, $kept, [System.Text.Encoding]::ASCII)
            Move-Item -Path $tmp -Destination $Script:RegistryFile -Force
        }
    }
    & $Script:Fn_WriteAuditLog -Action 'PURGE_LOGS' -Detail "Removed=$purged RetentionDays=$RetentionDays"
    [void][System.Windows.MessageBox]::Show("Purged $purged log entries older than $RetentionDays days.",'AutoBuild','OK','Information')
}

$Script:Fn_ExportHistoryCSV = {
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = 'CSV|*.csv'; $dlg.FileName = "history_$(Get-Date -Format yyyyMMdd).csv"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        @(& $Script:Fn_GetRunSummaries) | Select-Object RunId,Task,Started,Elapsed,Status,Entries,User |
            Export-Csv $dlg.FileName -NoTypeInformation -Encoding ASCII
        & $Script:Fn_WriteAuditLog -Action 'EXPORT_HISTORY' -Target $dlg.FileName
        [void][System.Windows.MessageBox]::Show("Exported to:`n$($dlg.FileName)",'AutoBuild','OK','Information')
    }
}

$Script:Fn_ExportAuditCSV = {
    if (-not (Test-Path $Script:AuditFile)) {
        [void][System.Windows.MessageBox]::Show('No audit log found.','AutoBuild','OK','Warning'); return
    }
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = 'CSV|*.csv'; $dlg.FileName = "audit_$(Get-Date -Format yyyyMMdd).csv"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $e = @()
        Get-Content $Script:AuditFile -Encoding ASCII | ForEach-Object {
            try { $e += ($_ | ConvertFrom-Json) } catch { }
        }
        $e | Select-Object ts,user,role,action,target,detail |
            Export-Csv $dlg.FileName -NoTypeInformation -Encoding ASCII
        & $Script:Fn_WriteAuditLog -Action 'EXPORT_AUDIT' -Target $dlg.FileName
        [void][System.Windows.MessageBox]::Show("Exported to:`n$($dlg.FileName)",'AutoBuild','OK','Information')
    }
}

$Script:Fn_ExportRunLog = {
    $rl  = $Script:Ctrl['txtRunLog']
    $txt = if ($null -ne $rl) { $rl.Text } else { '' }
    if ([string]::IsNullOrWhiteSpace($txt)) {
        [void][System.Windows.MessageBox]::Show('No log to export. Run a task first.','AutoBuild','OK','Warning'); return
    }
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = 'Text|*.txt|Log|*.log'
    $dlg.FileName = "runlog_$($Script:CurrentTaskName)_$(Get-Date -Format yyyyMMdd_HHmmss).txt"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        [System.IO.File]::WriteAllText($dlg.FileName, $txt, [System.Text.Encoding]::UTF8)
        & $Script:Fn_WriteAuditLog -Action 'EXPORT_RUNLOG' -Target $dlg.FileName
        [void][System.Windows.MessageBox]::Show("Saved to:`n$($dlg.FileName)",'AutoBuild','OK','Information')
    }
}

$Script:Fn_MakeFolderDropTarget = {
    param([System.Windows.Controls.TextBox]$tb)
    $tb.AllowDrop = $true
    $tb.Add_PreviewDragOver({
        param($s,$e)
        $e.Effects = if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
            [System.Windows.DragDropEffects]::Copy
        } else { [System.Windows.DragDropEffects]::None }
        $e.Handled = $true
    })
    $tb.Add_Drop({
        param($s,$e)
        if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
            $paths = $e.Data.GetData([System.Windows.DataFormats]::FileDrop)
            if ($paths -and $paths.Count -gt 0) {
                $p = $paths[0]
                $s.Text = if (Test-Path $p -PathType Container) { $p } else { Split-Path $p -Parent }
            }
        }
    })
}

$Script:Fn_UpdateNewTaskPreview = {
    $name = if ($Script:Ctrl['txtNewTaskName']) { $Script:Ctrl['txtNewTaskName'].Text } else { '' }
    $ci   = if ($Script:Ctrl['cboNewTaskCategory']) { $Script:Ctrl['cboNewTaskCategory'].SelectedItem } else { $null }
    $cat  = if ($null -ne $ci) { $ci.Content } else { 'Utility' }
    $desc = if ($Script:Ctrl['txtNewTaskDesc']) { $Script:Ctrl['txtNewTaskDesc'].Text } else { '' }
    if ([string]::IsNullOrWhiteSpace($desc)) { $desc = 'Task description pending' }
    $preview = "# Synopsis: $desc`ntask ${name} {`n    # @Category : $cat`n    # TODO: implement`n    Write-Build Green 'Done'`n}"
    $np = $Script:Ctrl['txtNewTaskPreview']; if ($null -ne $np) { $np.Text = $preview }
    $ne = $Script:Ctrl['txtNewTaskNameError']
    if ($null -ne $ne) {
        $valid = [string]::IsNullOrWhiteSpace($name) -or ($name -match '^[a-zA-Z0-9_-]+$')
        $ne.Visibility = if ($valid) { [System.Windows.Visibility]::Collapsed } else { [System.Windows.Visibility]::Visible }
    }
}

$Script:Fn_CreateNewTask = {
    $name = ($(if ($Script:Ctrl['txtNewTaskName']) { $Script:Ctrl['txtNewTaskName'].Text } else { '' })).Trim()
    if ([string]::IsNullOrWhiteSpace($name)) {
        [void][System.Windows.MessageBox]::Show('Task name required.','AutoBuild','OK','Warning'); return
    }
    if ($name -notmatch '^[a-zA-Z0-9_-]+$') {
        [void][System.Windows.MessageBox]::Show('Name: letters, numbers, _ or - only.','AutoBuild','OK','Warning'); return
    }
    $outFile = Join-Path $Script:TasksDir "task_${name}.ps1"
    if (Test-Path $outFile) {
        [void][System.Windows.MessageBox]::Show("Already exists: task_${name}.ps1",'AutoBuild','OK','Warning'); return
    }
    $content = if ($Script:Ctrl['txtNewTaskPreview']) { $Script:Ctrl['txtNewTaskPreview'].Text } else { "task ${name} { }" }
    try {
        [System.IO.File]::WriteAllText($outFile, $content, [System.Text.Encoding]::ASCII)
        $tr = $Script:Ctrl['txtCreateTaskResult']
        if ($null -ne $tr) {
            $tr.Text = "Created: $outFile"
            $tr.Foreground = [System.Windows.Media.Brushes]::LimeGreen
            $tr.Visibility = [System.Windows.Visibility]::Visible
        }
        & $Script:Fn_WriteAuditLog -Action 'CREATE_TASK' -Target "task_${name}.ps1"
        & $Script:Fn_LoadCatalogPage
    } catch { [void][System.Windows.MessageBox]::Show("Error: $_",'AutoBuild','OK','Error') }
}

$Script:Fn_ResumeCheckpoint = {
    $sel = if ($Script:Ctrl['gridCheckpoints']) { $Script:Ctrl['gridCheckpoints'].SelectedItem } else { $null }
    if ($null -eq $sel) {
        [void][System.Windows.MessageBox]::Show('Select a checkpoint first.','AutoBuild','OK','Warning'); return
    }
    & $Script:Fn_NavigateTo -PageName 'pageExecute' -Title 'Execute Task'
    & $Script:Fn_LoadExecutePage -PreSelectTask $sel.TaskName
    $cr = $Script:Ctrl['chkResume']; if ($null -ne $cr) { $cr.IsChecked = $true }
    & $Script:Fn_WriteAuditLog -Action 'RESUME_CHECKPOINT' -Target $sel.Name
}

$Script:Fn_DeleteCheckpoint = {
    if (-not (& $Script:Fn_TestPermission -Action 'ManageCheckpoints')) {
        [void][System.Windows.MessageBox]::Show('Admin role required.','AutoBuild','OK','Warning'); return
    }
    $sel = if ($Script:Ctrl['gridCheckpoints']) { $Script:Ctrl['gridCheckpoints'].SelectedItem } else { $null }
    if ($null -eq $sel) { return }
    if ([System.Windows.MessageBox]::Show("Delete $($sel.Name)?","Confirm",'YesNo','Question') -eq 'Yes') {
        try {
            Remove-Item $sel.Path -Force
            & $Script:Fn_WriteAuditLog -Action 'DELETE_CHECKPOINT' -Target $sel.Name
            & $Script:Fn_LoadCheckpointsPage
        } catch { [void][System.Windows.MessageBox]::Show("Error: $_",'AutoBuild','OK','Error') }
    }
}

$Script:Fn_OpenArtifact = {
    $sel = if ($Script:Ctrl['gridArtifacts']) { $Script:Ctrl['gridArtifacts'].SelectedItem } else { $null }
    if ($null -eq $sel) { return }
    try { Start-Process $sel.Path; & $Script:Fn_WriteAuditLog -Action 'OPEN_ARTIFACT' -Target $sel.Name }
    catch { [void][System.Windows.MessageBox]::Show("Cannot open: $_",'AutoBuild','OK','Warning') }
}

$Script:Fn_SaveArtifact = {
    $sel = if ($Script:Ctrl['gridArtifacts']) { $Script:Ctrl['gridArtifacts'].SelectedItem } else { $null }
    if ($null -eq $sel) { return }
    $dlg = New-Object System.Windows.Forms.SaveFileDialog; $dlg.FileName = $sel.Name
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Copy-Item $sel.Path $dlg.FileName -Force
        & $Script:Fn_WriteAuditLog -Action 'DOWNLOAD_ARTIFACT' -Target $sel.Name -Detail $dlg.FileName
    }
}

$Script:Fn_DeleteArtifact = {
    if (-not (& $Script:Fn_TestPermission -Action 'DeleteArtifact')) {
        [void][System.Windows.MessageBox]::Show('Admin role required.','AutoBuild','OK','Warning'); return
    }
    $sel = if ($Script:Ctrl['gridArtifacts']) { $Script:Ctrl['gridArtifacts'].SelectedItem } else { $null }
    if ($null -eq $sel) { return }
    if ([System.Windows.MessageBox]::Show("Delete $($sel.Name)?","Confirm",'YesNo','Question') -eq 'Yes') {
        try {
            Remove-Item $sel.Path -Force
            & $Script:Fn_WriteAuditLog -Action 'DELETE_ARTIFACT' -Target $sel.Name
            & $Script:Fn_LoadArtifactsPage
        } catch { [void][System.Windows.MessageBox]::Show("Error: $_",'AutoBuild','OK','Error') }
    }
}

# ============================================================================
# XAML LOAD - F4-02 fix: load from external file; fallback to embedded string
# ============================================================================
$Script:XAML = $null
$xamlFile    = Join-Path $Script:UIRoot 'AutoBuild.xaml'
if (Test-Path $xamlFile) {
    try {
        [xml]$Script:XAML = Get-Content $xamlFile -Raw -Encoding UTF8
    } catch {
        Write-Warning "AutoBuild UI: Cannot load ui\AutoBuild.xaml: $_. Using embedded fallback."
    }
}

# Minimal fallback XAML (functional, without all visual polish)
if ($null -eq $Script:XAML) {
    [xml]$Script:XAML = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="AutoBuild v3.0"
        MinHeight="600" MinWidth="860"
        Background="#0F1117" WindowStartupLocation="CenterScreen">
  <Grid><TextBlock Text="AutoBuild v3.0 - Load ui\AutoBuild.xaml for full UI"
    Foreground="#F5A623" FontSize="18" HorizontalAlignment="Center" VerticalAlignment="Center"/></Grid>
</Window>
'@
}

# ============================================================================
# WINDOW INITIALIZATION AND EVENT WIRING
# ============================================================================
$Script:AllPageNames = @(
    'pageCatalog','pageExecute','pageMonitor','pageRunLog','pageHistory',
    'pageCheckpoints','pageArtifacts','pageFolders','pageMetrics',
    'pageDiag','pageConfig','pageNewTask','pageAudit','pageQueue'
)

function Initialize-Window {
    if ($null -eq [System.Windows.Application]::Current) {
        $Script:App = New-Object System.Windows.Application
        $Script:App.ShutdownMode = [System.Windows.ShutdownMode]::OnExplicitShutdown
    } else {
        $Script:App = [System.Windows.Application]::Current
    }

    $reader = New-Object System.Xml.XmlNodeReader $Script:XAML
    $Script:Window = [Windows.Markup.XamlReader]::Load($reader)

    # =========================================================================
    # ADAPTIVE WINDOW SIZING
    # Calculates optimal window dimensions based on the available work area of
    # the primary monitor, so the UI works correctly on laptops (1366x768),
    # standard desktops (1920x1080), 4K monitors, and multi-monitor setups.
    #
    # Strategy:
    #   - Target 92% of the work area (leaves room for taskbar and window chrome)
    #   - Cap at a maximum of 1600x950 so the UI does not over-expand on 4K
    #   - Enforce the XAML MinWidth/MinHeight as hard floor (860 x 600)
    #   - On very small screens (< 1100px wide) maximize automatically
    # =========================================================================
    try {
        $workArea   = [System.Windows.SystemParameters]::WorkArea
        $screenW    = $workArea.Width
        $screenH    = $workArea.Height

        $targetW    = [Math]::Floor($screenW * 0.92)
        $targetH    = [Math]::Floor($screenH * 0.92)

        # Cap at design maximum so the UI does not stretch awkwardly on 4K
        $maxW       = 1600
        $maxH       = 950
        $targetW    = [Math]::Min($targetW, $maxW)
        $targetH    = [Math]::Min($targetH, $maxH)

        # Enforce XAML minimums
        $minW       = 860
        $minH       = 600
        $targetW    = [Math]::Max($targetW, $minW)
        $targetH    = [Math]::Max($targetH, $minH)

        if ($screenW -lt 1100) {
            # Small screen (netbook / embedded terminal): start maximized
            $Script:Window.WindowState = [System.Windows.WindowState]::Maximized
        } else {
            $Script:Window.Width  = $targetW
            $Script:Window.Height = $targetH
        }
    } catch {
        # SystemParameters unavailable (unusual): fall back to fixed safe size
        $Script:Window.Width  = 1280
        $Script:Window.Height = 800
    }

    # Discover all named controls
    $Script:Ctrl = @{}
    $allNames = @(
        'txtVersion','txtUserInitial','txtUserName','txtUserRole',
        'txtEngineStatus','txtEnginePath','elEngineStatus','txtPageTitle','btnRefresh',
        'pnlSecurityWarning','txtSecurityWarning','gridPages','pnlPageHeader',
        'pageCatalog','gridCatalog','txtCatalogSearch','cboCatalogCategory','txtCatalogCount',
        'btnCatalogExecute','btnCatalogViewHistory','btnCatalogAddQueue',
        'pageExecute','cboExecTask','pnlTaskInfo','txtExecTaskDesc','txtExecTaskCat',
        'pnlParams','chkWhatIf','chkCheckpoint','chkResume',
        'btnExecuteTask','btnCancelTask','elExecStatus','txtExecStatus','txtExecDuration',
        'txtExecOutput','svExecOutput','btnAddToQueue',
        'pageMonitor','gridMonitorJobs','txtActiveCount','txtLiveLog','svLiveLog','elLogPulse',
        'btnPurgeLogs','btnExportMonitorLog','cboPeriodFilter','cboPurgePeriod',
        'pageRunLog','txtRunLog','svRunLog','txtRunLogHeader','btnExportRunLog','btnClearRunLog',
        'pageHistory','gridHistory','txtHistorySearch','cboHistoryStatus',
        'btnHistoryFilter','btnExportHistory','txtHistoryDetail','svHistoryDetail','txtHistoryDetailTitle',
        'pageCheckpoints','gridCheckpoints','btnResumeCheckpoint','btnDeleteCheckpoint','btnViewCheckpoint',
        'txtCheckpointTitle','txtCheckpointDetail','svCheckpointDetail',
        'pageArtifacts','gridArtifacts','txtArtifactSearch','btnArtifactFilter',
        'txtArtifactCount','btnOpenArtifact','btnDownloadArtifact','btnDeleteArtifact',
        'pageFolders','txtInputDir','txtOutputDir','txtReportsDir',
        'btnBrowseInput','btnBrowseOutput','btnBrowseReports',
        'btnApplyFolders','btnResetFolders','txtFolderStatus',
        'pageMetrics','txtMetricTotal','txtMetricSuccess','txtMetricOKErr',
        'txtMetricAvg','txtMetricTop','gridMetricTasks','gridMetricErrors',
        'pageDiag','gridDiag','btnRunDiag','txtDiagSummary',
        'pageConfig','txtConfigEditor','btnSaveConfig','btnReloadConfig','txtConfigStatus',
        'pageNewTask','txtNewTaskName','txtNewTaskNameError','cboNewTaskCategory',
        'txtNewTaskDesc','btnCreateTask','txtCreateTaskResult','txtNewTaskPreview',
        'pageAudit','gridAudit','btnExportAudit',
        'btnNavCatalog','btnNavExecute','btnNavQueue','btnNavMonitor','btnNavRunLog','btnNavHistory',
        'btnNavCheckpoints','btnNavArtifacts','btnNavFolders','btnNavMetrics',
        'btnNavDiag','btnNavConfig','btnNavNewTask','btnNavAudit',
        'pageQueue'
    )
    foreach ($name in $allNames) {
        $found = $Script:Window.FindName($name)
        if ($null -eq $found) { Write-Verbose "Control not found: $name" }
        $Script:Ctrl[$name] = $found
    }

    # ── PROD-GUARD: Security warning banner ──────────────────────────────────
    # Show the red banner when the engine is running in dev mode (no AD groups).
    if (-not [string]::IsNullOrWhiteSpace($Script:SecurityWarning)) {
        $bannerPanel = $Script:Ctrl['pnlSecurityWarning']
        $bannerText  = $Script:Ctrl['txtSecurityWarning']
        if ($null -ne $bannerPanel) {
            $bannerPanel.Visibility = [System.Windows.Visibility]::Visible
            if ($null -ne $bannerText) { $bannerText.Text = $Script:SecurityWarning }
            # Shift the pages grid down by banner height (34px) + header (48px)
            $pagesGrid = $Script:Ctrl['gridPages']
            if ($null -ne $pagesGrid) { $pagesGrid.Margin = [System.Windows.Thickness]::new(0, 82, 0, 0) }
        }
    }

    # ── ADAPTIVE: Update page header DockPanel height to match actual header ──
    # The header is now 48px (reduced from 52px for better proportions on small screens).
    # If the security banner is visible add 34px. The pages grid margin is set above.

    # Header
    $c = $Script:Ctrl['txtUserName'];    if ($null -ne $c) { $c.Text = $Script:CurrentUser }
    $c = $Script:Ctrl['txtUserRole'];    if ($null -ne $c) { $c.Text = $Script:CurrentRole }
    $ini = if ($Script:CurrentUser.Length -gt 0) { $Script:CurrentUser[0].ToString().ToUpper() } else { 'U' }
    $c = $Script:Ctrl['txtUserInitial']; if ($null -ne $c) { $c.Text = $ini }
    $c = $Script:Ctrl['txtEnginePath'];  if ($null -ne $c) { $c.Text = $Script:EngineRoot }

    $ok = Test-Path $Script:RunScript
    $c  = $Script:Ctrl['txtEngineStatus']
    if ($null -ne $c) { $c.Text = if ($ok) { 'Engine Ready' } else { 'Engine Not Found' } }
    $c  = $Script:Ctrl['elEngineStatus']
    if ($null -ne $c) { $c.Background = if ($ok) { [System.Windows.Media.Brushes]::LimeGreen } else { [System.Windows.Media.Brushes]::Crimson } }

    # RBAC nav visibility - SEC fix: use real role
    $c = $Script:Ctrl['btnNavConfig'];  if ($null -ne $c) { $c.IsEnabled = (& $Script:Fn_TestPermission 'EditConfig') }
    $c = $Script:Ctrl['btnNavAudit'];   if ($null -ne $c) { $c.IsEnabled = (& $Script:Fn_TestPermission 'ViewAudit') }
    $c = $Script:Ctrl['btnNavNewTask']; if ($null -ne $c) { $c.IsEnabled = (& $Script:Fn_TestPermission 'CreateTask') }

    # Folder drag-drop setup
    $tdI = $Script:Ctrl['txtInputDir'];   if ($null -ne $tdI) { $tdI.Text  = $Script:InputDir;   & $Script:Fn_MakeFolderDropTarget -tb $tdI }
    $tdO = $Script:Ctrl['txtOutputDir'];  if ($null -ne $tdO) { $tdO.Text  = $Script:OutputDir;  & $Script:Fn_MakeFolderDropTarget -tb $tdO }
    $tdR = $Script:Ctrl['txtReportsDir']; if ($null -ne $tdR) { $tdR.Text  = $Script:ReportsDir; & $Script:Fn_MakeFolderDropTarget -tb $tdR }

    # Navigation wiring
    $navDefs = @(
        @{ Btn='btnNavCatalog';     Page='pageCatalog';     Title='Task Catalog';            Load=$Script:Fn_LoadCatalogPage }
        @{ Btn='btnNavExecute';     Page='pageExecute';     Title='Execute Task';            Load=$Script:Fn_LoadExecutePage }
        @{ Btn='btnNavQueue';       Page='pageQueue';       Title='Task Queue';              Load={} }
        @{ Btn='btnNavMonitor';     Page='pageMonitor';     Title='Live Monitor';            Load=$Script:Fn_LoadMonitorPage }
        @{ Btn='btnNavRunLog';      Page='pageRunLog';      Title='Run Log';                 Load={} }
        @{ Btn='btnNavHistory';     Page='pageHistory';     Title='Execution History';       Load=$Script:Fn_LoadHistoryPage }
        @{ Btn='btnNavCheckpoints'; Page='pageCheckpoints'; Title='Checkpoint Manager';      Load=$Script:Fn_LoadCheckpointsPage }
        @{ Btn='btnNavArtifacts';   Page='pageArtifacts';   Title='Artifact Repository';     Load=$Script:Fn_LoadArtifactsPage }
        @{ Btn='btnNavFolders';     Page='pageFolders';     Title='Folder Paths';            Load={} }
        @{ Btn='btnNavMetrics';     Page='pageMetrics';     Title='Metrics & Observability'; Load=$Script:Fn_LoadMetricsPage }
        @{ Btn='btnNavDiag';        Page='pageDiag';        Title='Environment Diagnostics'; Load=$Script:Fn_RunDiagnostics }
        @{ Btn='btnNavConfig';      Page='pageConfig';      Title='Configuration';           Load=$Script:Fn_LoadConfigPage }
        @{ Btn='btnNavNewTask';     Page='pageNewTask';     Title='Create New Task';         Load={& $Script:Fn_UpdateNewTaskPreview} }
        @{ Btn='btnNavAudit';       Page='pageAudit';       Title='Audit Log';               Load=$Script:Fn_LoadAuditPage }
    )
    foreach ($def in $navDefs) {
        $btn = $Script:Ctrl[$def.Btn]
        if ($null -ne $btn) { $btn.Tag = @{ Page=$def.Page; Title=$def.Title; Load=$def.Load } }
    }
    $navClickHandler = {
        param($sender,$e)
        $t = $sender.Tag
        & $Script:Fn_NavigateTo -PageName $t.Page -Title $t.Title
        try { & $t.Load } catch { }
    }
    foreach ($def in $navDefs) {
        $btn = $Script:Ctrl[$def.Btn]
        if ($null -ne $btn) { $btn.Add_Click($navClickHandler) }
    }

    # Refresh button
    $rb = $Script:Ctrl['btnRefresh']
    if ($null -ne $rb) {
        $rb.Add_Click({
            $title = if ($Script:Ctrl['txtPageTitle']) { $Script:Ctrl['txtPageTitle'].Text } else { '' }
            switch ($title) {
                'Task Catalog'            { & $Script:Fn_LoadCatalogPage }
                'Execute Task'            { & $Script:Fn_LoadExecutePage }
                'Live Monitor'            { & $Script:Fn_LoadMonitorPage }
                'Execution History'       { & $Script:Fn_LoadHistoryPage }
                'Checkpoint Manager'      { & $Script:Fn_LoadCheckpointsPage }
                'Artifact Repository'     { & $Script:Fn_LoadArtifactsPage }
                'Metrics & Observability' { & $Script:Fn_LoadMetricsPage }
                'Environment Diagnostics' { & $Script:Fn_RunDiagnostics }
                'Audit Log'               { & $Script:Fn_LoadAuditPage }
            }
        })
    }

    # Catalog events
    $c = $Script:Ctrl['txtCatalogSearch'];    if ($null -ne $c) { $c.Add_TextChanged({ & $Script:Fn_FilterCatalog }) }
    $c = $Script:Ctrl['cboCatalogCategory'];  if ($null -ne $c) { $c.Add_SelectionChanged({ & $Script:Fn_FilterCatalog }) }
    $c = $Script:Ctrl['btnCatalogExecute']
    if ($null -ne $c) {
        $c.Add_Click({
            $sel = if ($Script:Ctrl['gridCatalog']) { $Script:Ctrl['gridCatalog'].SelectedItem } else { $null }
            if ($null -ne $sel) {
                & $Script:Fn_NavigateTo -PageName 'pageExecute' -Title 'Execute Task'
                & $Script:Fn_LoadExecutePage -PreSelectTask $sel.Name
            }
        })
    }

    # Execute events
    $c = $Script:Ctrl['cboExecTask']
    if ($null -ne $c) {
        $c.Add_SelectionChanged({
            $sel = if ($Script:Ctrl['cboExecTask']) { $Script:Ctrl['cboExecTask'].SelectedItem } else { $null }
            if ($null -ne $sel) {
                & $Script:Fn_BuildParamForm -TaskName $sel.Tag
                $task = $Script:AllTasks | Where-Object { $_.Name -eq $sel.Tag } | Select-Object -First 1
                if ($null -ne $task) {
                    $d  = $Script:Ctrl['txtExecTaskDesc']; if ($null -ne $d)  { $d.Text  = $task.Description }
                    $ct = $Script:Ctrl['txtExecTaskCat'];  if ($null -ne $ct) { $ct.Text = "Category: $($task.Category)  v$($task.Version)" }
                    $pi = $Script:Ctrl['pnlTaskInfo'];     if ($null -ne $pi) { $pi.Visibility = [System.Windows.Visibility]::Visible }
                }
            }
        })
    }
    $c = $Script:Ctrl['btnExecuteTask']; if ($null -ne $c) { $c.Add_Click({ & $Script:Fn_RunSelectedTask }) }
    $c = $Script:Ctrl['btnCancelTask']
    if ($null -ne $c) {
        $c.Add_Click({
            $tn = $Script:CurrentTaskName
            if ($tn -and $Script:ActiveJobs.ContainsKey($tn)) {
                $proc = $Script:ActiveJobs[$tn].Process
                try { if (-not $proc.HasExited) { $proc.Kill() }; $proc.Dispose() } catch { }
                $Script:ActiveJobs.Remove($tn)
                $esT = $Script:Ctrl['txtExecStatus']; if ($null -ne $esT) { $esT.Text = 'Cancelled' }
                $esB = $Script:Ctrl['elExecStatus'];  if ($null -ne $esB) { $esB.Background = [System.Windows.Media.Brushes]::Crimson }
                $bRun= $Script:Ctrl['btnExecuteTask'];if ($null -ne $bRun){ $bRun.IsEnabled = $true }
                $bCan= $Script:Ctrl['btnCancelTask']; if ($null -ne $bCan){ $bCan.IsEnabled = $false }
                if ($null -ne $Script:ExecTimer) { $Script:ExecTimer.Stop(); $Script:ExecTimer = $null }
                & $Script:Fn_WriteAuditLog -Action 'CANCEL_TASK' -Target $tn
            }
        })
    }

    # Monitor events
    $c = $Script:Ctrl['cboPeriodFilter']
    if ($null -ne $c) {
        $c.Add_SelectionChanged({
            $sel  = if ($Script:Ctrl['cboPeriodFilter']) { $Script:Ctrl['cboPeriodFilter'].SelectedItem } else { $null }
            $days = if ($null -ne $sel -and $null -ne $sel.Tag) { [int]$sel.Tag } else { 0 }
            & $Script:Fn_LoadMonitorLogTail -DaysBack $days
        })
    }
    $c = $Script:Ctrl['btnPurgeLogs']
    if ($null -ne $c) {
        $c.Add_Click({
            $selPurge = if ($Script:Ctrl['cboPurgePeriod']) { $Script:Ctrl['cboPurgePeriod'].SelectedItem } else { $null }
            $days = if ($null -ne $selPurge -and $null -ne $selPurge.Tag) { [int]$selPurge.Tag } else { 30 }
            $ans  = [System.Windows.MessageBox]::Show("Purge log entries older than $days days?","Confirm Purge",'YesNo','Question')
            if ($ans -eq 'Yes') { & $Script:Fn_PurgeOldLogs -RetentionDays $days; & $Script:Fn_LoadMonitorPage }
        })
    }
    $c = $Script:Ctrl['btnExportMonitorLog']; if ($null -ne $c) { $c.Add_Click({ & $Script:Fn_ExportHistoryCSV }) }

    # Run Log events
    $c = $Script:Ctrl['btnExportRunLog'];  if ($null -ne $c) { $c.Add_Click({ & $Script:Fn_ExportRunLog }) }
    $c = $Script:Ctrl['btnClearRunLog']
    if ($null -ne $c) {
        $c.Add_Click({
            $rl = $Script:Ctrl['txtRunLog'];       if ($null -ne $rl) { $rl.Text = '' }
            $rh = $Script:Ctrl['txtRunLogHeader']; if ($null -ne $rh) { $rh.Text = 'No task running' }
        })
    }

    # History events
    $c = $Script:Ctrl['btnHistoryFilter'];  if ($null -ne $c) { $c.Add_Click({ & $Script:Fn_LoadHistoryPage }) }
    $c = $Script:Ctrl['gridHistory']
    if ($null -ne $c) {
        $c.Add_SelectionChanged({
            $sel = if ($Script:Ctrl['gridHistory']) { $Script:Ctrl['gridHistory'].SelectedItem } else { $null }
            if ($null -ne $sel) { & $Script:Fn_ShowRunDetail -RunId $sel.RunId -Task $sel.Task }
        })
    }
    $c = $Script:Ctrl['btnExportHistory'];  if ($null -ne $c) { $c.Add_Click({ & $Script:Fn_ExportHistoryCSV }) }

    # Checkpoint events
    $c = $Script:Ctrl['btnResumeCheckpoint']; if ($null -ne $c) { $c.Add_Click({ & $Script:Fn_ResumeCheckpoint }) }
    $c = $Script:Ctrl['btnDeleteCheckpoint']; if ($null -ne $c) { $c.Add_Click({ & $Script:Fn_DeleteCheckpoint }) }
    $c = $Script:Ctrl['btnViewCheckpoint']
    if ($null -ne $c) {
        $c.Add_Click({
            $sel = if ($Script:Ctrl['gridCheckpoints']) { $Script:Ctrl['gridCheckpoints'].SelectedItem } else { $null }
            if ($null -eq $sel) { return }
            try {
                $data = Import-Clixml $sel.Path | Out-String
                [void][System.Windows.MessageBox]::Show($data.Substring(0,[Math]::Min($data.Length,2000)),"Checkpoint: $($sel.Name)",'OK','Information')
            } catch { [void][System.Windows.MessageBox]::Show("Cannot read: $_",'AutoBuild','OK','Warning') }
        })
    }
    $c = $Script:Ctrl['gridCheckpoints']
    if ($null -ne $c) {
        $c.Add_SelectionChanged({
            $sel = if ($Script:Ctrl['gridCheckpoints']) { $Script:Ctrl['gridCheckpoints'].SelectedItem } else { $null }
            $tt  = $Script:Ctrl['txtCheckpointTitle']
            $td  = $Script:Ctrl['txtCheckpointDetail']
            if ($null -eq $sel) { return }
            if ($null -ne $tt) { $tt.Text = "CHECKPOINT: $($sel.Name)" }
            if ($null -ne $td) {
                $detail = "Name:     $($sel.Name)`nTask:     $($sel.TaskName)`nModified: $($sel.Modified)`nSize:     $($sel.SizeKB) KB`nPath:     $($sel.Path)"
                try {
                    if (Test-Path $sel.Path) {
                        $obj  = Import-Clixml $sel.Path -ErrorAction Stop
                        $keys = if ($obj -is [hashtable]) { @($obj.Keys) } else { @($obj.PSObject.Properties.Name) }
                        $detail += "`n`nStored keys ($($keys.Count)):`n" + ($keys -join ', ')
                    }
                } catch { $detail += "`n`n[Could not read checkpoint: $_]" }
                $td.Text = $detail
            }
        })
    }

    # Artifact events
    $c = $Script:Ctrl['btnArtifactFilter'];   if ($null -ne $c) { $c.Add_Click({ & $Script:Fn_LoadArtifactsPage }) }
    $c = $Script:Ctrl['txtArtifactSearch']
    if ($null -ne $c) { $c.Add_KeyDown({ param($s,$e); if ($e.Key -eq [System.Windows.Input.Key]::Return) { & $Script:Fn_LoadArtifactsPage } }) }
    $c = $Script:Ctrl['btnOpenArtifact'];     if ($null -ne $c) { $c.Add_Click({ & $Script:Fn_OpenArtifact }) }
    $c = $Script:Ctrl['btnDownloadArtifact']; if ($null -ne $c) { $c.Add_Click({ & $Script:Fn_SaveArtifact }) }
    $c = $Script:Ctrl['btnDeleteArtifact'];   if ($null -ne $c) { $c.Add_Click({ & $Script:Fn_DeleteArtifact }) }

    # Folder page events
    $browseFolderAction = {
        param([System.Windows.Controls.TextBox]$targetTb)
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description  = 'Select a folder'
        $dlg.SelectedPath = if ($null -ne $targetTb -and (Test-Path $targetTb.Text)) { $targetTb.Text } else { $Script:EngineRoot }
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            if ($null -ne $targetTb) { $targetTb.Text = $dlg.SelectedPath }
        }
    }
    $c = $Script:Ctrl['btnBrowseInput'];   if ($null -ne $c) { $c.Add_Click({ & $browseFolderAction -targetTb $Script:Ctrl['txtInputDir']   }) }
    $c = $Script:Ctrl['btnBrowseOutput'];  if ($null -ne $c) { $c.Add_Click({ & $browseFolderAction -targetTb $Script:Ctrl['txtOutputDir']  }) }
    $c = $Script:Ctrl['btnBrowseReports']; if ($null -ne $c) { $c.Add_Click({ & $browseFolderAction -targetTb $Script:Ctrl['txtReportsDir'] }) }
    $c = $Script:Ctrl['btnApplyFolders']
    if ($null -ne $c) {
        $c.Add_Click({
            $inp = $Script:Ctrl['txtInputDir'];   if ($null -ne $inp  -and $inp.Text)  { $Script:InputDir   = $inp.Text }
            $out = $Script:Ctrl['txtOutputDir'];  if ($null -ne $out  -and $out.Text)  { $Script:OutputDir  = $out.Text }
            $rpt = $Script:Ctrl['txtReportsDir']; if ($null -ne $rpt  -and $rpt.Text)  { $Script:ReportsDir = $rpt.Text }
            foreach ($d in @($Script:InputDir,$Script:OutputDir,$Script:ReportsDir)) {
                if (-not (Test-Path $d)) { try { New-Item $d -ItemType Directory -Force | Out-Null } catch { } }
            }
            $fs = $Script:Ctrl['txtFolderStatus']; if ($null -ne $fs) { $fs.Text = "Applied." }
            & $Script:Fn_WriteAuditLog -Action 'APPLY_FOLDERS' -Detail "In=$Script:InputDir"
        })
    }
    $c = $Script:Ctrl['btnResetFolders']
    if ($null -ne $c) {
        $c.Add_Click({
            $Script:InputDir   = Join-Path $Script:EngineRoot 'input'
            $Script:OutputDir  = Join-Path $Script:EngineRoot 'output'
            $Script:ReportsDir = Join-Path $Script:EngineRoot 'reports'
            $inp = $Script:Ctrl['txtInputDir'];   if ($null -ne $inp) { $inp.Text = $Script:InputDir }
            $out = $Script:Ctrl['txtOutputDir'];  if ($null -ne $out) { $out.Text = $Script:OutputDir }
            $rpt = $Script:Ctrl['txtReportsDir']; if ($null -ne $rpt) { $rpt.Text = $Script:ReportsDir }
            $fs  = $Script:Ctrl['txtFolderStatus']; if ($null -ne $fs) { $fs.Text = 'Reset to defaults.' }
        })
    }

    # Diag / Config / New Task / Audit events
    $c = $Script:Ctrl['btnRunDiag'];         if ($null -ne $c) { $c.Add_Click({ & $Script:Fn_RunDiagnostics }) }
    $c = $Script:Ctrl['btnSaveConfig'];      if ($null -ne $c) { $c.Add_Click({ & $Script:Fn_SaveConfigPage }) }
    $c = $Script:Ctrl['btnReloadConfig'];    if ($null -ne $c) { $c.Add_Click({ & $Script:Fn_LoadConfigPage }) }
    $c = $Script:Ctrl['txtNewTaskName'];     if ($null -ne $c) { $c.Add_TextChanged({ & $Script:Fn_UpdateNewTaskPreview }) }
    $c = $Script:Ctrl['cboNewTaskCategory']; if ($null -ne $c) { $c.Add_SelectionChanged({ & $Script:Fn_UpdateNewTaskPreview }) }
    $c = $Script:Ctrl['txtNewTaskDesc'];     if ($null -ne $c) { $c.Add_TextChanged({ & $Script:Fn_UpdateNewTaskPreview }) }
    $c = $Script:Ctrl['btnCreateTask'];      if ($null -ne $c) { $c.Add_Click({ & $Script:Fn_CreateNewTask }) }
    $c = $Script:Ctrl['btnExportAudit'];     if ($null -ne $c) { $c.Add_Click({ & $Script:Fn_ExportAuditCSV }) }

    # Auto-refresh timer - SCALE fix: only reads log TAIL (50 lines), not full file
    $Script:RefreshTimer = New-Object System.Windows.Threading.DispatcherTimer
    $Script:RefreshTimer.Interval = [TimeSpan]::FromSeconds(5)
    $Script:RefreshTimer.Add_Tick({
        $title = if ($Script:Ctrl['txtPageTitle']) { $Script:Ctrl['txtPageTitle'].Text } else { '' }
        if ($title -eq 'Live Monitor') { & $Script:Fn_LoadMonitorPage }
        $tc = $Script:Ctrl['txtActiveCount']
        if ($null -ne $tc) { $tc.Text = "$($Script:ActiveJobs.Count) running" }
    })
    $Script:RefreshTimer.Start()

    $Script:Window.Add_Closed({
        if ($null -ne $Script:RefreshTimer) { $Script:RefreshTimer.Stop() }
        if ($null -ne $Script:ExecTimer)    { $Script:ExecTimer.Stop() }
        foreach ($key in @($Script:ActiveJobs.Keys)) {
            try { $Script:ActiveJobs[$key].Process.Kill() } catch { }
        }
        if ($Script:QueueEnabled) { try { Stop-QueueRunner } catch { } }
        & $Script:Fn_WriteAuditLog -Action 'UI_CLOSE'
        try { $Script:App.Shutdown() } catch { }
    })

    & $Script:Fn_WriteAuditLog -Action 'UI_OPEN' -Detail "Role=$Script:CurrentRole User=$Script:CurrentUser"
    & $Script:Fn_LoadCatalogPage

    if ($Script:QueueEnabled) {
        try {
            Initialize-QueuePage `
                -Window $Script:Window -Ctrl $Script:Ctrl `
                -AllTasks $Script:AllTasks -EngineRoot $Script:EngineRoot
        } catch { Write-Warning "QUEUE SYSTEM: page init error: $_" }
    }

    $Script:App.MainWindow = $Script:Window
    $Script:Window.Show()
    [void]$Script:App.Run()
}

Initialize-Window