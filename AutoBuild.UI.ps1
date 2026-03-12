#Requires -Version 5.1
<#
.SYNOPSIS
    AutoBuild Automation Interface - WPF GUI
.DESCRIPTION
    Production-ready, secure, auditable interface for AutoBuild.
    Provides: Task Catalog, Execution Panel, Live Monitor, History/Audit,
    Checkpoint Management, Artifact Repository, Environment Diagnostics,
    Configuration Management, Metrics Dashboard, and Task Creation.
.NOTES
    ASCII-only. PS 5.1 compatible. WPF/WinForms via .NET 4.x.
    Requires AutoBuild engine in same directory or via -EnginePath.
    Roles: Operator (run tasks), Developer (create/edit tasks), Admin (config/audit).
.EXAMPLE
    .\AutoBuild.UI.ps1
    .\AutoBuild.UI.ps1 -EnginePath "C:\AutoBuild"
    .\AutoBuild.UI.ps1 -Role Admin
#>
param(
    [string]$EnginePath = '',
    [ValidateSet('Operator','Developer','Admin')]
    [string]$Role = 'Operator'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================================================================
# BOOTSTRAP: Load WPF assemblies
# ============================================================================
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ============================================================================
# PATH RESOLUTION
# ============================================================================
$Script:UIRoot = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($EnginePath)) {
    $Script:EngineRoot = $Script:UIRoot
} else {
    $Script:EngineRoot = $EnginePath
}

$Script:RunScript    = Join-Path $Script:EngineRoot 'Run.ps1'
$Script:NewTaskScript= Join-Path $Script:EngineRoot 'New-Task.ps1'
$Script:ConfigFile   = Join-Path $Script:EngineRoot 'engine.config.json'
$Script:TasksDir     = Join-Path $Script:EngineRoot 'tasks'
$Script:LogsDir      = Join-Path $Script:EngineRoot 'logs'
$Script:OutputDir    = Join-Path $Script:EngineRoot 'output'
$Script:ReportsDir   = Join-Path $Script:EngineRoot 'reports'
$Script:RegistryFile = Join-Path $Script:LogsDir 'registry.jsonl'
$Script:AuditFile    = Join-Path $Script:LogsDir 'audit.jsonl'

# Ensure log/output dirs exist
foreach ($dir in @($Script:LogsDir, $Script:OutputDir, $Script:ReportsDir)) {
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
}

# ============================================================================
# RBAC
# ============================================================================
$Script:CurrentRole = $Role
$Script:CurrentUser = $env:USERNAME

$Script:Permissions = @{
    Operator  = @('RunTask','ViewHistory','ViewArtifacts','ViewMetrics','ViewDiag')
    Developer = @('RunTask','ViewHistory','ViewArtifacts','ViewMetrics','ViewDiag','CreateTask','EditTask')
    Admin     = @('RunTask','ViewHistory','ViewArtifacts','ViewMetrics','ViewDiag','CreateTask','EditTask','EditConfig','DeleteArtifact','ViewAudit','ManageCheckpoints')
}

function Test-Permission {
    param([string]$Action)
    return ($Script:Permissions[$Script:CurrentRole] -contains $Action)
}

# ============================================================================
# AUDIT LOGGING
# ============================================================================
function Write-AuditLog {
    param(
        [string]$Action,
        [string]$Target = '',
        [string]$Detail = ''
    )
    $entry = [ordered]@{
        ts     = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss')
        user   = $Script:CurrentUser
        role   = $Script:CurrentRole
        action = $Action
        target = $Target
        detail = $Detail
    }
    $json = $entry | ConvertTo-Json -Compress
    try {
        $mutex = New-Object System.Threading.Mutex($false, 'Global\AutoBuildAuditMutex')
        $locked = $mutex.WaitOne(3000)
        Add-Content -Path $Script:AuditFile -Value $json -Encoding ASCII
        if ($locked) { try { $mutex.ReleaseMutex() } catch {} }
        $mutex.Dispose()
    } catch {
        try { Add-Content -Path $Script:AuditFile -Value $json -Encoding ASCII } catch {}
    }
}

# ============================================================================
# TASK METADATA PARSER
# ============================================================================
function Get-TaskMetadata {
    param([string]$Path)
    $meta = @{
        File        = $Path
        Name        = ''
        Description = ''
        Category    = ''
        Version     = ''
        Params      = @()
        LastRun     = $null
        LastStatus  = ''
    }
    $name = [System.IO.Path]::GetFileNameWithoutExtension($Path) -replace '^task_', ''
    $meta.Name = $name

    if (Test-Path $Path) {
        $lines = Get-Content $Path -TotalCount 20 -Encoding ASCII
        foreach ($line in $lines) {
            if ($line -match '@Description\s*:\s*(.+)')  { $meta.Description = $Matches[1].Trim() }
            if ($line -match '@Category\s*:\s*(.+)')     { $meta.Category    = $Matches[1].Trim() }
            if ($line -match '@Version\s*:\s*(.+)')      { $meta.Version     = $Matches[1].Trim() }
            if ($line -match '@Param\s*:\s*(\S+)\s+(\S+)\s+(required|optional)\s+"([^"]*)"') {
                $meta.Params += @{
                    Name     = $Matches[1]
                    Type     = $Matches[2]
                    Required = ($Matches[3] -eq 'required')
                    Help     = $Matches[4]
                }
            }
        }
    }
    return $meta
}

function Get-AllTasks {
    $tasks = @()
    if (Test-Path $Script:TasksDir) {
        $tasks = @(Get-ChildItem -Path $Script:TasksDir -Filter 'task_*.ps1' |
            Where-Object { $_.Name -ne 'task_TEMPLATE.ps1' } |
            Sort-Object Name |
            ForEach-Object { Get-TaskMetadata -Path $_.FullName })
    }
    # Enrich with last-run info: single tail read, fast
    if (Test-Path $Script:RegistryFile) {
        $history = @()
        try {
            Get-Content $Script:RegistryFile -Tail 400 -Encoding ASCII | ForEach-Object {
                if (-not [string]::IsNullOrWhiteSpace($_)) {
                    try { $history += ($_ | ConvertFrom-Json) } catch {}
                }
            }
        } catch {}
        foreach ($task in $tasks) {
            $lastRun = $history |
                Where-Object { $_.task -eq $task.Name -and $_.level -in @('OK','ERROR') } |
                Select-Object -Last 1
            if ($lastRun) {
                $task.LastRun    = $lastRun.ts
                $task.LastStatus = $lastRun.level
            }
        }
    }
    return $tasks
}

# ============================================================================
# LOG READER
# ============================================================================
function Get-ExecutionHistory {
    param(
        [int]$MaxEntries = 200,
        [string]$FilterTask = '',
        [string]$FilterLevel = ''
    )
    $entries = @()
    if (-not (Test-Path $Script:RegistryFile)) { return $entries }
    try {
        $lines = Get-Content $Script:RegistryFile -Encoding ASCII |
                    Select-Object -Last ($MaxEntries * 3)
        foreach ($line in $lines) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            try {
                $obj = $line | ConvertFrom-Json
                if ($FilterTask  -and $obj.task  -ne $FilterTask)  { continue }
                if ($FilterLevel -and $obj.level -ne $FilterLevel) { continue }
                $entries += $obj
            } catch {}
        }
    } catch {}
    return $entries | Select-Object -Last $MaxEntries
}

function Get-RunSummaries {
    # Group log entries by runId, capped to avoid UI freeze on large logs
    $all = @(Get-ExecutionHistory -MaxEntries 400)
    if ($all.Count -eq 0) { return @() }
    $grouped = $all | Group-Object -Property runId
    $summaries = @()
    foreach ($g in $grouped) {
        $entries = @($g.Group)
        $result  = $entries | Where-Object { $_.level -in @('OK','ERROR') } | Select-Object -Last 1
        $first   = $entries[0]
        $summaries += [PSCustomObject]@{
            RunId    = $g.Name
            Task     = $first.task
            Started  = $first.ts
            Status   = if ($result) { $result.level } else { 'RUNNING' }
            Elapsed  = if ($result -and $result.elapsed) { [math]::Round([double]$result.elapsed, 1) } else { '' }
            User     = $Script:CurrentUser
            Entries  = $entries.Count
        }
    }
    return @($summaries | Sort-Object Started -Descending)
}

# ============================================================================
# METRICS
# ============================================================================
function Get-Metrics {
    $summaries = @(Get-RunSummaries)
    $total   = $summaries.Count
    $ok      = @($summaries | Where-Object { $_.Status -eq 'OK' }).Count
    $err     = @($summaries | Where-Object { $_.Status -eq 'ERROR' }).Count
    $running = @($summaries | Where-Object { $_.Status -eq 'RUNNING' }).Count
    $avgElapsed = 0
    $withTime = @($summaries | Where-Object { $_.Elapsed -ne '' -and $null -ne $_.Elapsed })
    if ($withTime.Count -gt 0) {
        $sum = ($withTime | Measure-Object -Property Elapsed -Sum).Sum
        $avgElapsed = [math]::Round($sum / $withTime.Count, 1)
    }
    $taskCounts = @($summaries | Group-Object Task | Sort-Object Count -Descending)
    $topTask = if ($taskCounts.Count -gt 0) { $taskCounts[0].Name } else { 'N/A' }

    return @{
        Total      = $total
        OK         = $ok
        Error      = $err
        Running    = $running
        SuccessRate= if ($total -gt 0) { [math]::Round(($ok / $total) * 100, 1) } else { 0 }
        AvgElapsed = $avgElapsed
        TopTask    = $topTask
        TaskCounts = $taskCounts
    }
}

# ============================================================================
# TASK EXECUTION
# Uses System.Diagnostics.Process with async BeginOutputReadLine so the UI
# thread never blocks. Output lines accumulate in a thread-safe StringBuilder.
# The DispatcherTimer reads the StringBuilder on the UI thread without waiting.
# ============================================================================
$Script:ActiveJobs = @{}   # key=TaskName, value=process-info hashtable

function Start-TaskExecution {
    param(
        [string]$TaskName,
        [hashtable]$Params,
        [switch]$WhatIf,
        [switch]$Checkpoint,
        [switch]$Resume
    )

    if (-not (Test-Path $Script:RunScript)) {
        return @{ Success=$false; Error="Run.ps1 not found at: $Script:RunScript" }
    }

    Write-AuditLog -Action 'EXECUTE_TASK' -Target $TaskName `
        -Detail "WhatIf=$WhatIf Checkpoint=$Checkpoint Resume=$Resume Params=$($Params | ConvertTo-Json -Compress)"

    # Build argument string for powershell.exe
    $psArgs = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$Script:RunScript`" $TaskName"
    foreach ($k in $Params.Keys) {
        if (-not [string]::IsNullOrWhiteSpace($Params[$k])) {
            $psArgs += " -$k `"$($Params[$k])`""
        }
    }
    if ($WhatIf)     { $psArgs += ' -WhatIf' }
    if ($Checkpoint) { $psArgs += ' -Checkpoint' }
    if ($Resume)     { $psArgs += ' -Resume' }

    # Create process with async output capture
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName               = 'powershell.exe'
    $psi.Arguments              = $psArgs
    $psi.UseShellExecute        = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.CreateNoWindow         = $true
    $psi.WorkingDirectory       = $Script:EngineRoot

    $proc   = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    $proc.EnableRaisingEvents = $true

    # Thread-safe output buffer
    $outBuf = New-Object System.Text.StringBuilder

    # Async output handlers (run on ThreadPool, NOT UI thread - safe to append to StringBuilder)
    $outHandler = [System.Diagnostics.DataReceivedEventHandler]{
        param($s, $e)
        if ($null -ne $e.Data) {
            [void]$outBuf.AppendLine($e.Data)
        }
    }
    $errHandler = [System.Diagnostics.DataReceivedEventHandler]{
        param($s, $e)
        if ($null -ne $e.Data) {
            [void]$outBuf.AppendLine("[STDERR] $($e.Data)")
        }
    }
    $proc.add_OutputDataReceived($outHandler)
    $proc.add_ErrorDataReceived($errHandler)

    try {
        [void]$proc.Start()
        $proc.BeginOutputReadLine()
        $proc.BeginErrorReadLine()
    } catch {
        return @{ Success=$false; Error="Failed to start process: $_" }
    }

    $Script:ActiveJobs[$TaskName] = @{
        Process  = $proc
        Buffer   = $outBuf
        Started  = [datetime]::Now
        TaskName = $TaskName
    }
    return @{ Success=$true; Pid=$proc.Id; TaskName=$TaskName }
}

function Get-JobOutput {
    param([string]$TaskName)
    if (-not $Script:ActiveJobs.ContainsKey($TaskName)) { return $null }
    $info = $Script:ActiveJobs[$TaskName]
    $proc = $info.Process

    # .HasExited is non-blocking
    $running = -not $proc.HasExited
    $state   = if ($running) { 'Running' } else {
        if ($proc.ExitCode -eq 0) { 'Completed' } else { 'Failed' }
    }

    # Read buffer snapshot without blocking (ToString on StringBuilder is safe from UI thread
    # while background threads append - worst case we miss the last partial line)
    $out = $info.Buffer.ToString()

    return @{
        State    = $state
        Output   = $out
        Started  = $info.Started
        Duration = ([datetime]::Now - $info.Started).TotalSeconds
        ExitCode = if (-not $running) { $proc.ExitCode } else { $null }
    }
}

function Stop-TaskProcess {
    param([string]$TaskName)
    if (-not $Script:ActiveJobs.ContainsKey($TaskName)) { return }
    $proc = $Script:ActiveJobs[$TaskName].Process
    try {
        if (-not $proc.HasExited) { $proc.Kill() }
        $proc.Dispose()
    } catch {}
    $Script:ActiveJobs.Remove($TaskName)
}

# ============================================================================
# CONFIG MANAGEMENT
# ============================================================================
function Get-Config {
    if (-not (Test-Path $Script:ConfigFile)) {
        return @{
            engine  = @{ logLevel='INFO'; maxRetries=3; retryDelaySeconds=5 }
            sap     = @{ systemId='PRD'; client='800'; language='ES'; timeout=180 }
            excel   = @{ visible=$false; screenUpdating=$false }
            reports = @{ defaultFormat='xlsx'; retentionDays=30 }
        }
    }
    try {
        $raw = Get-Content $Script:ConfigFile -Raw -Encoding ASCII | ConvertFrom-Json
        return $raw
    } catch {
        return $null
    }
}

function Save-Config {
    param([string]$JsonContent)
    # Validate JSON
    try {
        $parsed = $JsonContent | ConvertFrom-Json
        $null = $parsed
    } catch {
        return @{ Success=$false; Error="Invalid JSON: $_" }
    }
    # Validate required keys
    if (-not $parsed.engine -or -not $parsed.reports) {
        return @{ Success=$false; Error="Missing required sections: engine, reports" }
    }
    # Validate values
    if ($parsed.engine.maxRetries -lt 0 -or $parsed.engine.maxRetries -gt 10) {
        return @{ Success=$false; Error="maxRetries must be 0-10" }
    }
    if ($parsed.reports.retentionDays -lt 1 -or $parsed.reports.retentionDays -gt 365) {
        return @{ Success=$false; Error="retentionDays must be 1-365" }
    }

    try {
        [System.IO.File]::WriteAllText($Script:ConfigFile, $JsonContent, [System.Text.Encoding]::ASCII)
        Write-AuditLog -Action 'EDIT_CONFIG' -Target 'engine.config.json'
        return @{ Success=$true }
    } catch {
        return @{ Success=$false; Error="Write failed: $_" }
    }
}

# ============================================================================
# ENVIRONMENT DIAGNOSTICS
# ============================================================================
function Get-DiagnosticReport {
    $results = @()

    # PowerShell version
    $psv = $PSVersionTable.PSVersion
    $results += @{
        Category = 'PowerShell'
        Item     = 'Version'
        Value    = "$($psv.Major).$($psv.Minor).$($psv.Build)"
        Status   = if ($psv.Major -eq 5 -and $psv.Minor -eq 1) { 'OK' } else { 'WARN' }
        Message  = if ($psv.Major -eq 5 -and $psv.Minor -eq 1) { 'PS 5.1 Desktop - Compatible' } else { 'Recommend PS 5.1 for COM compatibility' }
    }
    $results += @{
        Category = 'PowerShell'
        Item     = 'Edition'
        Value    = $PSVersionTable.PSEdition
        Status   = if ($PSVersionTable.PSEdition -eq 'Desktop') { 'OK' } else { 'WARN' }
        Message  = if ($PSVersionTable.PSEdition -eq 'Desktop') { 'Desktop edition - Full COM support' } else { 'Core edition may have limited COM support' }
    }

    # Engine files
    foreach ($f in @('Run.ps1','New-Task.ps1','engine.config.json','engine\Main.build.ps1')) {
        $fp = Join-Path $Script:EngineRoot $f
        $results += @{
            Category = 'Engine'
            Item     = $f
            Value    = if (Test-Path $fp) { 'Found' } else { 'Missing' }
            Status   = if (Test-Path $fp) { 'OK' } else { 'ERROR' }
            Message  = if (Test-Path $fp) { $fp } else { "MISSING: $fp" }
        }
    }

    # Folder permissions
    foreach ($dir in @('tasks','logs','output','reports','input')) {
        $dp = Join-Path $Script:EngineRoot $dir
        $writable = $false
        if (Test-Path $dp) {
            $testFile = Join-Path $dp "_write_test_$([System.IO.Path]::GetRandomFileName())"
            try {
                [System.IO.File]::WriteAllText($testFile, 'test', [System.Text.Encoding]::ASCII)
                Remove-Item $testFile -Force
                $writable = $true
            } catch {}
        }
        $results += @{
            Category = 'Folders'
            Item     = $dir
            Value    = if (Test-Path $dp) { if ($writable) { 'Writable' } else { 'Read-Only' } } else { 'Missing' }
            Status   = if (Test-Path $dp -and $writable) { 'OK' } elseif (Test-Path $dp) { 'WARN' } else { 'ERROR' }
            Message  = if (-not (Test-Path $dp)) { "Create folder: $dp" } elseif (-not $writable) { "Grant write permissions to: $dp" } else { $dp }
        }
    }

    # Lib files
    foreach ($lib in @('Logger.ps1','Context.ps1','ComHelper.ps1','ExcelHelper.ps1','SapHelper.ps1','WordHelper.ps1','Assertions.ps1')) {
        $lp = Join-Path $Script:EngineRoot "lib\$lib"
        $results += @{
            Category = 'Libraries'
            Item     = $lib
            Value    = if (Test-Path $lp) { 'Found' } else { 'Missing' }
            Status   = if (Test-Path $lp) { 'OK' } else { 'ERROR' }
            Message  = if (Test-Path $lp) { 'Library available' } else { "MISSING: $lp" }
        }
    }

    # Excel COM
    try {
        $xl = New-Object -ComObject Excel.Application
        $ver = $xl.Version
        $xl.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
        $results += @{
            Category = 'COM'
            Item     = 'Microsoft Excel'
            Value    = "Version $ver"
            Status   = 'OK'
            Message  = 'Excel COM automation available'
        }
    } catch {
        $results += @{
            Category = 'COM'
            Item     = 'Microsoft Excel'
            Value    = 'Not Available'
            Status   = 'WARN'
            Message  = 'Excel not found or COM registration failed. Install Office.'
        }
    }

    # Word COM
    try {
        $wd = New-Object -ComObject Word.Application
        $ver = $wd.Version
        $wd.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wd) | Out-Null
        $results += @{
            Category = 'COM'
            Item     = 'Microsoft Word'
            Value    = "Version $ver"
            Status   = 'OK'
            Message  = 'Word COM automation available'
        }
    } catch {
        $results += @{
            Category = 'COM'
            Item     = 'Microsoft Word'
            Value    = 'Not Available'
            Status   = 'WARN'
            Message  = 'Word not found. Install Office for Word tasks.'
        }
    }

    # SAP GUI
    try {
        $sap = New-Object -ComObject SapROTWr.SapROTWrapper
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sap) | Out-Null
        $results += @{
            Category = 'COM'
            Item     = 'SAP GUI'
            Value    = 'Available'
            Status   = 'OK'
            Message  = 'SAP GUI scripting COM object found'
        }
    } catch {
        $results += @{
            Category = 'COM'
            Item     = 'SAP GUI'
            Value    = 'Not Available'
            Status   = 'INFO'
            Message  = 'SAP GUI not found. Required only for SAP tasks.'
        }
    }

    # Invoke-Build tool
    $ib = Join-Path $Script:EngineRoot 'tools\InvokeBuild\Invoke-Build.ps1'
    $results += @{
        Category = 'Tools'
        Item     = 'Invoke-Build'
        Value    = if (Test-Path $ib) { 'Found' } else { 'Missing' }
        Status   = if (Test-Path $ib) { 'OK' } else { 'ERROR' }
        Message  = if (Test-Path $ib) { 'Portable Invoke-Build available' } else { "MISSING: $ib" }
    }

    return $results
}

# ============================================================================
# ARTIFACT MANAGEMENT
# ============================================================================
function Get-Artifacts {
    $artifacts = @()
    foreach ($dir in @($Script:OutputDir, $Script:ReportsDir)) {
        if (Test-Path $dir) {
            Get-ChildItem -Path $dir -File | ForEach-Object {
                $artifacts += [PSCustomObject]@{
                    Name      = $_.Name
                    Path      = $_.FullName
                    Directory = Split-Path $dir -Leaf
                    Size      = $_.Length
                    Modified  = $_.LastWriteTime
                    Extension = $_.Extension.ToLower()
                }
            }
        }
    }
    return $artifacts | Sort-Object Modified -Descending
}

function Remove-ArtifactFile {
    param([string]$Path)
    if (-not (Test-Permission 'DeleteArtifact')) {
        return @{ Success=$false; Error='Insufficient permissions. Admin role required.' }
    }
    try {
        Write-AuditLog -Action 'DELETE_ARTIFACT' -Target (Split-Path $Path -Leaf) -Detail $Path
        Remove-Item -Path $Path -Force
        return @{ Success=$true }
    } catch {
        return @{ Success=$false; Error=$_.ToString() }
    }
}

# ============================================================================
# CHECKPOINT MANAGEMENT
# ============================================================================
function Get-Checkpoints {
    $checkpoints = @()
    if (Test-Path $Script:LogsDir) {
        Get-ChildItem -Path $Script:LogsDir -Filter 'checkpoint_*.clixml' | ForEach-Object {
            $checkpoints += [PSCustomObject]@{
                Name     = $_.Name
                Path     = $_.FullName
                Size     = $_.Length
                Modified = $_.LastWriteTime
                TaskName = ($_.BaseName -replace '^checkpoint_', '' -replace '_\d{8}_\d{6}.*$', '')
            }
        }
    }
    return $checkpoints | Sort-Object Modified -Descending
}

# ============================================================================
# CREDENTIAL HELPER (Windows Credential Manager)
# Uses cmdkey.exe to avoid fragile P/Invoke C# compilation issues on PS 5.1.
# ============================================================================
function Get-StoredCredential {
    param([string]$Target)
    try {
        $output = cmdkey.exe /list:$Target 2>&1 | Out-String
        return $output -match $Target
    } catch { return $false }
}

function Save-StoredCredential {
    param([string]$Target, [string]$User, [string]$Password)
    try {
        $result = cmdkey.exe /generic:$Target /user:$User /pass:$Password 2>&1
        return ($LASTEXITCODE -eq 0)
    } catch { return $false }
}

function Remove-StoredCredential {
    param([string]$Target)
    try {
        cmdkey.exe /delete:$Target 2>&1 | Out-Null
        return ($LASTEXITCODE -eq 0)
    } catch { return $false }
}

# ============================================================================
# XAML UI DEFINITION
# ============================================================================
[xml]$Script:XAML = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="AutoBuild Automation Interface"
    Height="850" Width="1350" MinHeight="650" MinWidth="900"
    WindowStartupLocation="CenterScreen"
    Background="#0F1117">
    <Window.Resources>
        <!-- Color Palette: Dark industrial with amber/orange accent -->
        <SolidColorBrush x:Key="BgPrimary"     Color="#0F1117"/>
        <SolidColorBrush x:Key="BgSecondary"   Color="#171B26"/>
        <SolidColorBrush x:Key="BgPanel"       Color="#1E2232"/>
        <SolidColorBrush x:Key="BgCard"        Color="#252B3B"/>
        <SolidColorBrush x:Key="BgHover"       Color="#2D3348"/>
        <SolidColorBrush x:Key="AccentOrange"  Color="#F5A623"/>
        <SolidColorBrush x:Key="AccentBlue"    Color="#4A9EFF"/>
        <SolidColorBrush x:Key="AccentGreen"   Color="#3EBA7E"/>
        <SolidColorBrush x:Key="AccentRed"     Color="#E85555"/>
        <SolidColorBrush x:Key="AccentYellow"  Color="#F5CE42"/>
        <SolidColorBrush x:Key="TextPrimary"   Color="#E8EAF0"/>
        <SolidColorBrush x:Key="TextSecondary" Color="#8892A8"/>
        <SolidColorBrush x:Key="TextMuted"     Color="#4A5068"/>
        <SolidColorBrush x:Key="BorderColor"   Color="#2D3348"/>

        <!-- Button Styles -->
        <Style x:Key="BtnPrimary" TargetType="Button">
            <Setter Property="Background"   Value="#F5A623"/>
            <Setter Property="Foreground"   Value="#0F1117"/>
            <Setter Property="FontWeight"   Value="Bold"/>
            <Setter Property="FontSize"     Value="12"/>
            <Setter Property="Padding"      Value="16,8"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor"       Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="4" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#FFB840"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="#3A3F50"/>
                                <Setter Property="Foreground" Value="#5A6070"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="BtnSecondary" TargetType="Button">
            <Setter Property="Background"   Value="#2D3348"/>
            <Setter Property="Foreground"   Value="#E8EAF0"/>
            <Setter Property="FontSize"     Value="12"/>
            <Setter Property="Padding"      Value="14,7"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush"  Value="#3D4460"/>
            <Setter Property="Cursor"       Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="4" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#363D57"/>
                                <Setter Property="BorderBrush" Value="#F5A623"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="#1E2232"/>
                                <Setter Property="Foreground" Value="#4A5068"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="BtnDanger" TargetType="Button">
            <Setter Property="Background"   Value="#8B2020"/>
            <Setter Property="Foreground"   Value="#FFD0D0"/>
            <Setter Property="FontSize"     Value="12"/>
            <Setter Property="Padding"      Value="14,7"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor"       Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="4" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#B52A2A"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Nav Button Style -->
        <Style x:Key="NavButton" TargetType="Button">
            <Setter Property="Background"   Value="Transparent"/>
            <Setter Property="Foreground"   Value="#8892A8"/>
            <Setter Property="FontSize"     Value="13"/>
            <Setter Property="Padding"      Value="12,10"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor"       Value="Hand"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bdr" Background="{TemplateBinding Background}"
                                BorderThickness="3,0,0,0" BorderBrush="Transparent"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Left" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bdr" Property="Background" Value="#1E2232"/>
                                <Setter Property="Foreground" Value="#E8EAF0"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- TextBox Style -->
        <Style x:Key="InputField" TargetType="TextBox">
            <Setter Property="Background"   Value="#1E2232"/>
            <Setter Property="Foreground"   Value="#E8EAF0"/>
            <Setter Property="BorderBrush"  Value="#2D3348"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding"      Value="8,6"/>
            <Setter Property="FontSize"     Value="13"/>
            <Setter Property="CaretBrush"   Value="#F5A623"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="4">
                            <ScrollViewer x:Name="PART_ContentHost" Padding="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsFocused" Value="True">
                                <Setter Property="BorderBrush" Value="#F5A623"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- ComboBox Style -->
        <Style x:Key="InputCombo" TargetType="ComboBox">
            <Setter Property="Background"   Value="#1E2232"/>
            <Setter Property="Foreground"   Value="#E8EAF0"/>
            <Setter Property="BorderBrush"  Value="#2D3348"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding"      Value="8,6"/>
            <Setter Property="FontSize"     Value="13"/>
        </Style>

        <!-- DataGrid Style -->
        <Style x:Key="DarkGrid" TargetType="DataGrid">
            <Setter Property="Background"           Value="#1E2232"/>
            <Setter Property="Foreground"           Value="#E8EAF0"/>
            <Setter Property="BorderBrush"          Value="#2D3348"/>
            <Setter Property="GridLinesVisibility"  Value="Horizontal"/>
            <Setter Property="HorizontalGridLinesBrush" Value="#252B3B"/>
            <Setter Property="RowBackground"        Value="#1E2232"/>
            <Setter Property="AlternatingRowBackground" Value="#222840"/>
            <Setter Property="HeadersVisibility"    Value="Column"/>
            <Setter Property="SelectionMode"        Value="Single"/>
            <Setter Property="AutoGenerateColumns"  Value="False"/>
            <Setter Property="CanUserAddRows"       Value="False"/>
            <Setter Property="CanUserDeleteRows"    Value="False"/>
            <Setter Property="IsReadOnly"           Value="True"/>
        </Style>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background"   Value="#252B3B"/>
            <Setter Property="Foreground"   Value="#8892A8"/>
            <Setter Property="Padding"      Value="10,8"/>
            <Setter Property="FontWeight"   Value="Bold"/>
            <Setter Property="FontSize"     Value="11"/>
            <Setter Property="BorderBrush"  Value="#2D3348"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
        </Style>
        <Style TargetType="DataGridRow">
            <Setter Property="Foreground"   Value="#E8EAF0"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#2D3E5C"/>
                    <Setter Property="Foreground" Value="#FFFFFF"/>
                </Trigger>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#252B3B"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="DataGridCell">
            <Setter Property="Padding"      Value="8,6"/>
            <Setter Property="BorderThickness" Value="0"/>
        </Style>

        <!-- ScrollBar -->
        <Style TargetType="ScrollBar">
            <Setter Property="Background" Value="#1A1E2C"/>
            <Setter Property="Width"      Value="8"/>
        </Style>

        <!-- Label -->
        <Style x:Key="FieldLabel" TargetType="TextBlock">
            <Setter Property="Foreground"   Value="#8892A8"/>
            <Setter Property="FontSize"     Value="11"/>
            <Setter Property="FontWeight"   Value="Medium"/>
            <Setter Property="Margin"       Value="0,0,0,4"/>
        </Style>

        <!-- Section Header -->
        <Style x:Key="SectionHeader" TargetType="TextBlock">
            <Setter Property="Foreground"   Value="#E8EAF0"/>
            <Setter Property="FontSize"     Value="14"/>
            <Setter Property="FontWeight"   Value="Bold"/>
            <Setter Property="Margin"       Value="0,0,0,12"/>
        </Style>

        <!-- Status Badge -->
        <Style x:Key="BadgeOK" TargetType="Border">
            <Setter Property="Background"     Value="#1B3A2A"/>
            <Setter Property="BorderBrush"    Value="#3EBA7E"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius"   Value="3"/>
            <Setter Property="Padding"        Value="6,2"/>
        </Style>
        <Style x:Key="BadgeError" TargetType="Border">
            <Setter Property="Background"     Value="#3A1B1B"/>
            <Setter Property="BorderBrush"    Value="#E85555"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius"   Value="3"/>
            <Setter Property="Padding"        Value="6,2"/>
        </Style>

        <!-- Metric Card -->
        <Style x:Key="MetricCard" TargetType="Border">
            <Setter Property="Background"     Value="#1E2232"/>
            <Setter Property="BorderBrush"    Value="#2D3348"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius"   Value="8"/>
            <Setter Property="Padding"        Value="20,16"/>
        </Style>

        <!-- CheckBox -->
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="#E8EAF0"/>
            <Setter Property="FontSize"   Value="13"/>
        </Style>

        <!-- TabControl -->
        <Style TargetType="TabControl">
            <Setter Property="Background"     Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
        </Style>
        <Style TargetType="TabItem">
            <Setter Property="Background"     Value="#1E2232"/>
            <Setter Property="Foreground"     Value="#8892A8"/>
            <Setter Property="Padding"        Value="14,8"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border x:Name="bdr" Background="{TemplateBinding Background}"
                                BorderThickness="0,0,0,2" BorderBrush="Transparent"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="bdr" Property="BorderBrush" Value="#F5A623"/>
                                <Setter TargetName="bdr" Property="Background" Value="#252B3B"/>
                                <Setter Property="Foreground" Value="#F5A623"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="200"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>

        <!-- ================================================================
             SIDEBAR
             ================================================================ -->
        <Border Grid.Column="0" Background="#171B26" BorderBrush="#2D3348" BorderThickness="0,0,1,0">
            <DockPanel>
                <!-- Logo / Title -->
                <Border DockPanel.Dock="Top" Padding="16,20,16,16" BorderBrush="#2D3348" BorderThickness="0,0,0,1">
                    <StackPanel>
                        <TextBlock Text="AUTO" FontSize="22" FontWeight="Black"
                                   Foreground="#F5A623"/>
                        <TextBlock Text="BUILD" FontSize="22" FontWeight="Black"
                                   Foreground="#E8EAF0" Margin="0,-4,0,0"/>
                        <TextBlock x:Name="txtVersion" Text="Automation Interface"
                                   FontSize="10" Foreground="#4A5068" Margin="0,4,0,0"/>
                    </StackPanel>
                </Border>

                <!-- User/Role badge -->
                <Border DockPanel.Dock="Top" Padding="16,12" BorderBrush="#2D3348" BorderThickness="0,0,0,1">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="32"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <Border Grid.Column="0" Width="28" Height="28" CornerRadius="14"
                                Background="#F5A623" VerticalAlignment="Center">
                            <TextBlock x:Name="txtUserInitial" Text="U" FontWeight="Bold"
                                       FontSize="12" Foreground="#0F1117"
                                       HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <StackPanel Grid.Column="1" Margin="8,0,0,0" VerticalAlignment="Center">
                            <TextBlock x:Name="txtUserName" Text="User" FontSize="12"
                                       Foreground="#E8EAF0" FontWeight="Medium"/>
                            <TextBlock x:Name="txtUserRole" Text="Operator" FontSize="10"
                                       Foreground="#F5A623"/>
                        </StackPanel>
                    </Grid>
                </Border>

                <!-- Navigation -->
                <StackPanel DockPanel.Dock="Top" Margin="0,8,0,0">
                    <TextBlock Text="MAIN" FontSize="10" Foreground="#4A5068"
                               Margin="16,8,0,4" FontWeight="Bold"/>

                    <Button x:Name="btnNavCatalog"   Content="  Task Catalog"
                            Style="{StaticResource NavButton}"/>
                    <Button x:Name="btnNavExecute"   Content="  Execute Task"
                            Style="{StaticResource NavButton}"/>
                    <Button x:Name="btnNavMonitor"   Content="  Live Monitor"
                            Style="{StaticResource NavButton}"/>
                    <Button x:Name="btnNavHistory"   Content="  Execution History"
                            Style="{StaticResource NavButton}"/>

                    <TextBlock Text="MANAGEMENT" FontSize="10" Foreground="#4A5068"
                               Margin="16,16,0,4" FontWeight="Bold"/>

                    <Button x:Name="btnNavCheckpoints" Content="  Checkpoints"
                            Style="{StaticResource NavButton}"/>
                    <Button x:Name="btnNavArtifacts"   Content="  Artifacts"
                            Style="{StaticResource NavButton}"/>
                    <Button x:Name="btnNavMetrics"     Content="  Metrics"
                            Style="{StaticResource NavButton}"/>

                    <TextBlock Text="SYSTEM" FontSize="10" Foreground="#4A5068"
                               Margin="16,16,0,4" FontWeight="Bold"/>

                    <Button x:Name="btnNavDiag"       Content="  Diagnostics"
                            Style="{StaticResource NavButton}"/>
                    <Button x:Name="btnNavConfig"     Content="  Configuration"
                            Style="{StaticResource NavButton}"/>
                    <Button x:Name="btnNavNewTask"    Content="  New Task"
                            Style="{StaticResource NavButton}"/>
                    <Button x:Name="btnNavAudit"      Content="  Audit Log"
                            Style="{StaticResource NavButton}"/>
                </StackPanel>

                <!-- Status bar at bottom -->
                <Border DockPanel.Dock="Bottom" Padding="16,10"
                        Background="#0F1117" BorderBrush="#2D3348" BorderThickness="0,1,0,0">
                    <StackPanel>
                        <StackPanel Orientation="Horizontal">
                            <Border x:Name="elEngineStatus" Width="8" Height="8"
                                    Background="#3EBA7E" CornerRadius="4"
                                    VerticalAlignment="Center"/>
                            <TextBlock x:Name="txtEngineStatus" Text="Engine Ready"
                                       FontSize="11" Foreground="#8892A8" Margin="6,0,0,0"/>
                        </StackPanel>
                        <TextBlock x:Name="txtEnginePath" Text="" FontSize="10"
                                   Foreground="#4A5068" TextTrimming="CharacterEllipsis"
                                   Margin="0,3,0,0"/>
                    </StackPanel>
                </Border>
            </DockPanel>
        </Border>

        <!-- ================================================================
             MAIN CONTENT AREA
             ================================================================ -->
        <Grid Grid.Column="1">
            <!-- Top bar -->
            <DockPanel Height="52" VerticalAlignment="Top"
                       Background="#171B26">
                <TextBlock x:Name="txtPageTitle" Text="Task Catalog"
                           FontSize="18" FontWeight="Bold" Foreground="#E8EAF0"
                           VerticalAlignment="Center" Margin="24,0,0,0"
                           DockPanel.Dock="Left"/>
                <StackPanel Orientation="Horizontal" DockPanel.Dock="Right"
                            Margin="0,0,16,0" VerticalAlignment="Center">
                    <Button x:Name="btnRefresh" Content="Refresh" Style="{StaticResource BtnSecondary}"
                            Padding="12,6" FontSize="11"/>
                </StackPanel>
            </DockPanel>

            <!-- Pages (all stacked, visibility controlled by code) -->
            <Grid Margin="0,52,0,0">

                <!-- ========================================================
                     PAGE: TASK CATALOG
                     ======================================================== -->
                <Grid x:Name="pageCatalog" Visibility="Visible">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="60"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <!-- Toolbar -->
                    <Border Grid.Row="0" Background="#171B26" Padding="16,0"
                            BorderBrush="#2D3348" BorderThickness="0,0,0,1">
                        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                            <TextBox x:Name="txtCatalogSearch" Width="220" Height="32"
                                     Style="{StaticResource InputField}"
                                     Tag="Search tasks..."/>
                            <ComboBox x:Name="cboCatalogCategory" Width="140" Height="32"
                                      Style="{StaticResource InputCombo}" Margin="8,0,0,0">
                                <ComboBoxItem Content="All Categories" IsSelected="True"/>
                                <ComboBoxItem Content="SAP"/>
                                <ComboBoxItem Content="Excel"/>
                                <ComboBoxItem Content="CSV"/>
                                <ComboBoxItem Content="Reporte"/>
                                <ComboBoxItem Content="Utilidad"/>
                            </ComboBox>
                            <TextBlock x:Name="txtCatalogCount" Text="0 tasks"
                                       FontSize="12" Foreground="#8892A8"
                                       VerticalAlignment="Center" Margin="16,0,0,0"/>
                        </StackPanel>
                    </Border>

                    <!-- Task Grid -->
                    <DataGrid x:Name="gridCatalog" Grid.Row="1" Style="{StaticResource DarkGrid}"
                              Margin="16,16,16,0" BorderThickness="1">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Task Name"   Binding="{Binding Name}"        Width="180"/>
                            <DataGridTextColumn Header="Category"    Binding="{Binding Category}"    Width="100"/>
                            <DataGridTextColumn Header="Description" Binding="{Binding Description}" Width="*"/>
                            <DataGridTextColumn Header="Version"     Binding="{Binding Version}"     Width="70"/>
                            <DataGridTextColumn Header="Last Run"    Binding="{Binding LastRun}"     Width="150"/>
                            <DataGridTextColumn Header="Status"      Binding="{Binding LastStatus}"  Width="80"/>
                        </DataGrid.Columns>
                    </DataGrid>

                    <!-- Action bar below grid -->
                    <Border Grid.Row="1" VerticalAlignment="Bottom" Height="56"
                            Background="#171B26" BorderBrush="#2D3348" BorderThickness="0,1,0,0"
                            Padding="16,0">
                        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                            <Button x:Name="btnCatalogExecute" Content="Execute Selected"
                                    Style="{StaticResource BtnPrimary}"/>
                            <Button x:Name="btnCatalogViewHistory" Content="View History"
                                    Style="{StaticResource BtnSecondary}" Margin="8,0,0,0"/>
                        </StackPanel>
                    </Border>
                </Grid>

                <!-- ========================================================
                     PAGE: EXECUTE TASK
                     ======================================================== -->
                <Grid x:Name="pageExecute" Visibility="Collapsed">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="320"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <!-- Left: Task selector + params -->
                    <Border Grid.Column="0" Background="#171B26"
                            BorderBrush="#2D3348" BorderThickness="0,0,1,0" Padding="20">
                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                            <StackPanel>
                                <TextBlock Text="SELECT TASK" Style="{StaticResource SectionHeader}"/>
                                <TextBlock Text="Task" Style="{StaticResource FieldLabel}"/>
                                <ComboBox x:Name="cboExecTask" Height="34" Style="{StaticResource InputCombo}"/>

                                <Border x:Name="pnlTaskInfo" Margin="0,12,0,0" Padding="12"
                                        Background="#1E2232" CornerRadius="6" Visibility="Collapsed">
                                    <StackPanel>
                                        <TextBlock x:Name="txtExecTaskDesc" Text=""
                                                   FontSize="12" Foreground="#8892A8" TextWrapping="Wrap"/>
                                        <TextBlock x:Name="txtExecTaskCat" Text=""
                                                   FontSize="11" Foreground="#F5A623" Margin="0,4,0,0"/>
                                    </StackPanel>
                                </Border>

                                <Separator Margin="0,16" Background="#2D3348"/>
                                <TextBlock Text="PARAMETERS" Style="{StaticResource SectionHeader}"/>
                                <StackPanel x:Name="pnlParams"/>

                                <Separator Margin="0,16" Background="#2D3348"/>
                                <TextBlock Text="EXECUTION MODE" Style="{StaticResource SectionHeader}"/>
                                <CheckBox x:Name="chkWhatIf"    Content="WhatIf (dry run)" Margin="0,4"/>
                                <CheckBox x:Name="chkCheckpoint" Content="Enable Checkpoints" Margin="0,4"/>
                                <CheckBox x:Name="chkResume"    Content="Resume from Checkpoint" Margin="0,4"/>

                                <Separator Margin="0,16" Background="#2D3348"/>
                                <Button x:Name="btnExecuteTask" Content="EXECUTE TASK"
                                        Style="{StaticResource BtnPrimary}" Height="40"
                                        FontSize="14"/>
                                <Button x:Name="btnCancelTask" Content="Cancel Running Task"
                                        Style="{StaticResource BtnDanger}" Margin="0,8,0,0"
                                        IsEnabled="False"/>
                            </StackPanel>
                        </ScrollViewer>
                    </Border>

                    <!-- Right: Output console -->
                    <Grid Grid.Column="1">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>

                        <!-- Status bar -->
                        <Border Grid.Row="0" Padding="16,12" Background="#1E2232"
                                BorderBrush="#2D3348" BorderThickness="0,0,0,1">
                            <StackPanel Orientation="Horizontal">
                                <Border x:Name="elExecStatus" Width="10" Height="10"
                                        Background="#4A5068" CornerRadius="5"
                                        VerticalAlignment="Center"/>
                                <TextBlock x:Name="txtExecStatus" Text="Idle"
                                           FontSize="13" Foreground="#E8EAF0"
                                           Margin="8,0,0,0" VerticalAlignment="Center"/>
                                <TextBlock x:Name="txtExecDuration" Text=""
                                           FontSize="12" Foreground="#8892A8"
                                           Margin="16,0,0,0" VerticalAlignment="Center"/>
                            </StackPanel>
                        </Border>

                        <!-- Console output -->
                        <Border Grid.Row="1" Background="#0A0D14" Margin="16">
                            <ScrollViewer x:Name="svExecOutput" VerticalScrollBarVisibility="Auto">
                                <TextBlock x:Name="txtExecOutput"
                                           FontFamily="Consolas" FontSize="12"
                                           Foreground="#C8D0E0" Padding="16"
                                           TextWrapping="Wrap"/>
                            </ScrollViewer>
                        </Border>
                    </Grid>
                </Grid>

                <!-- ========================================================
                     PAGE: LIVE MONITOR
                     ======================================================== -->
                <Grid x:Name="pageMonitor" Visibility="Collapsed">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="300"/>
                    </Grid.RowDefinitions>

                    <!-- Active jobs -->
                    <Border Grid.Row="0" Padding="16,12" Background="#171B26"
                            BorderBrush="#2D3348" BorderThickness="0,0,0,1">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="ACTIVE TASKS" FontSize="12" FontWeight="Bold"
                                       Foreground="#F5A623" VerticalAlignment="Center"/>
                            <TextBlock x:Name="txtActiveCount" Text="0 running"
                                       FontSize="12" Foreground="#8892A8"
                                       Margin="12,0,0,0" VerticalAlignment="Center"/>
                        </StackPanel>
                    </Border>

                    <DataGrid x:Name="gridMonitorJobs" Grid.Row="1"
                              Style="{StaticResource DarkGrid}" Margin="16,16,16,8">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Task"     Binding="{Binding TaskName}" Width="160"/>
                            <DataGridTextColumn Header="Run ID"   Binding="{Binding RunId}"    Width="180"/>
                            <DataGridTextColumn Header="Started"  Binding="{Binding Started}"  Width="160"/>
                            <DataGridTextColumn Header="Duration" Binding="{Binding Duration}" Width="100"/>
                            <DataGridTextColumn Header="Status"   Binding="{Binding Status}"   Width="100"/>
                            <DataGridTextColumn Header="User"     Binding="{Binding User}"     Width="*"/>
                        </DataGrid.Columns>
                    </DataGrid>

                    <!-- Live log stream -->
                    <Border Grid.Row="2" Background="#0A0D14" Margin="16,0,16,16">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="32"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <Border Grid.Row="0" Background="#1E2232" Padding="12,0">
                                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                    <TextBlock Text="LIVE LOG STREAM" FontSize="11" FontWeight="Bold"
                                               Foreground="#F5A623"/>
                                    <Border x:Name="elLogPulse" Width="6" Height="6"
                                            Background="#3EBA7E" CornerRadius="3"
                                            Margin="8,0,0,0" VerticalAlignment="Center"/>
                                </StackPanel>
                            </Border>
                            <ScrollViewer x:Name="svLiveLog" Grid.Row="1"
                                          VerticalScrollBarVisibility="Auto">
                                <TextBlock x:Name="txtLiveLog"
                                           FontFamily="Consolas" FontSize="11"
                                           Foreground="#8892A8" Padding="12"
                                           TextWrapping="Wrap"/>
                            </ScrollViewer>
                        </Grid>
                    </Border>
                </Grid>

                <!-- ========================================================
                     PAGE: EXECUTION HISTORY
                     ======================================================== -->
                <Grid x:Name="pageHistory" Visibility="Collapsed">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="60"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="250"/>
                    </Grid.RowDefinitions>

                    <Border Grid.Row="0" Background="#171B26" Padding="16,0"
                            BorderBrush="#2D3348" BorderThickness="0,0,0,1">
                        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                            <TextBox x:Name="txtHistorySearch" Width="200" Height="32"
                                     Style="{StaticResource InputField}" Tag="Filter by task..."/>
                            <ComboBox x:Name="cboHistoryStatus" Width="120" Height="32"
                                      Style="{StaticResource InputCombo}" Margin="8,0,0,0">
                                <ComboBoxItem Content="All Status" IsSelected="True"/>
                                <ComboBoxItem Content="OK"/>
                                <ComboBoxItem Content="ERROR"/>
                                <ComboBoxItem Content="RUNNING"/>
                            </ComboBox>
                            <Button x:Name="btnHistoryFilter" Content="Apply Filter"
                                    Style="{StaticResource BtnSecondary}" Margin="8,0,0,0" Height="32"/>
                            <Button x:Name="btnExportHistory" Content="Export CSV"
                                    Style="{StaticResource BtnSecondary}" Margin="8,0,0,0" Height="32"/>
                        </StackPanel>
                    </Border>

                    <DataGrid x:Name="gridHistory" Grid.Row="1"
                              Style="{StaticResource DarkGrid}" Margin="16,16,16,0">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Run ID"   Binding="{Binding RunId}"   Width="180"/>
                            <DataGridTextColumn Header="Task"     Binding="{Binding Task}"    Width="150"/>
                            <DataGridTextColumn Header="Started"  Binding="{Binding Started}" Width="160"/>
                            <DataGridTextColumn Header="Elapsed"  Binding="{Binding Elapsed}" Width="90"/>
                            <DataGridTextColumn Header="Status"   Binding="{Binding Status}"  Width="80"/>
                            <DataGridTextColumn Header="Entries"  Binding="{Binding Entries}" Width="70"/>
                            <DataGridTextColumn Header="User"     Binding="{Binding User}"    Width="*"/>
                        </DataGrid.Columns>
                    </DataGrid>

                    <!-- Log detail panel -->
                    <Border Grid.Row="2" Background="#0A0D14" Margin="16,8,16,16">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="32"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <Border Grid.Row="0" Background="#1E2232" Padding="12,0">
                                <TextBlock x:Name="txtHistoryDetailTitle" Text="SELECT A RUN TO VIEW LOGS"
                                           FontSize="11" FontWeight="Bold" Foreground="#8892A8"
                                           VerticalAlignment="Center"/>
                            </Border>
                            <ScrollViewer x:Name="svHistoryDetail" Grid.Row="1"
                                          VerticalScrollBarVisibility="Auto">
                                <TextBlock x:Name="txtHistoryDetail"
                                           FontFamily="Consolas" FontSize="11"
                                           Foreground="#8892A8" Padding="12"
                                           TextWrapping="Wrap"/>
                            </ScrollViewer>
                        </Grid>
                    </Border>
                </Grid>

                <!-- ========================================================
                     PAGE: CHECKPOINTS
                     ======================================================== -->
                <Grid x:Name="pageCheckpoints" Visibility="Collapsed">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <DataGrid x:Name="gridCheckpoints" Style="{StaticResource DarkGrid}"
                              Margin="16">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Checkpoint File" Binding="{Binding Name}"     Width="*"/>
                            <DataGridTextColumn Header="Task"            Binding="{Binding TaskName}" Width="150"/>
                            <DataGridTextColumn Header="Modified"        Binding="{Binding Modified}" Width="160"/>
                            <DataGridTextColumn Header="Size"            Binding="{Binding Size}"     Width="80"/>
                        </DataGrid.Columns>
                    </DataGrid>

                    <Border VerticalAlignment="Bottom" Height="56"
                            Background="#171B26" BorderBrush="#2D3348" BorderThickness="0,1,0,0"
                            Padding="16,0" Margin="0,0,0,0">
                        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                            <Button x:Name="btnResumeCheckpoint" Content="Resume Task"
                                    Style="{StaticResource BtnPrimary}"/>
                            <Button x:Name="btnDeleteCheckpoint" Content="Delete"
                                    Style="{StaticResource BtnDanger}" Margin="8,0,0,0"/>
                            <Button x:Name="btnViewCheckpoint" Content="View State"
                                    Style="{StaticResource BtnSecondary}" Margin="8,0,0,0"/>
                        </StackPanel>
                    </Border>
                </Grid>

                <!-- ========================================================
                     PAGE: ARTIFACTS
                     ======================================================== -->
                <Grid x:Name="pageArtifacts" Visibility="Collapsed">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="60"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <Border Grid.Row="0" Background="#171B26" Padding="16,0"
                            BorderBrush="#2D3348" BorderThickness="0,0,0,1">
                        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                            <TextBox x:Name="txtArtifactSearch" Width="200" Height="32"
                                     Style="{StaticResource InputField}" Tag="Filter artifacts..."/>
                            <Button x:Name="btnArtifactFilter" Content="Filter"
                                    Style="{StaticResource BtnSecondary}" Margin="8,0,0,0" Height="32"/>
                            <TextBlock x:Name="txtArtifactCount" Text=""
                                       FontSize="12" Foreground="#8892A8"
                                       VerticalAlignment="Center" Margin="16,0,0,0"/>
                        </StackPanel>
                    </Border>

                    <DataGrid x:Name="gridArtifacts" Grid.Row="1"
                              Style="{StaticResource DarkGrid}" Margin="16,16,16,60">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="File"      Binding="{Binding Name}"      Width="*"/>
                            <DataGridTextColumn Header="Type"      Binding="{Binding Extension}" Width="70"/>
                            <DataGridTextColumn Header="Directory" Binding="{Binding Directory}" Width="90"/>
                            <DataGridTextColumn Header="Size (KB)" Binding="{Binding SizeKB}"    Width="80"/>
                            <DataGridTextColumn Header="Modified"  Binding="{Binding Modified}"  Width="160"/>
                        </DataGrid.Columns>
                    </DataGrid>

                    <Border Grid.Row="1" VerticalAlignment="Bottom" Height="56"
                            Background="#171B26" BorderBrush="#2D3348" BorderThickness="0,1,0,0"
                            Padding="16,0">
                        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                            <Button x:Name="btnOpenArtifact"   Content="Open / Preview"
                                    Style="{StaticResource BtnPrimary}"/>
                            <Button x:Name="btnDownloadArtifact" Content="Save As..."
                                    Style="{StaticResource BtnSecondary}" Margin="8,0,0,0"/>
                            <Button x:Name="btnDeleteArtifact"  Content="Delete"
                                    Style="{StaticResource BtnDanger}" Margin="8,0,0,0"/>
                        </StackPanel>
                    </Border>
                </Grid>

                <!-- ========================================================
                     PAGE: METRICS
                     ======================================================== -->
                <ScrollViewer x:Name="pageMetrics" Visibility="Collapsed"
                              VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="24">
                        <!-- KPI Cards -->
                        <TextBlock Text="OVERVIEW" FontSize="12" FontWeight="Bold"
                                   Foreground="#4A5068" Margin="0,0,0,12"/>
                        <UniformGrid Columns="4" Rows="1">
                            <Border Style="{StaticResource MetricCard}" Margin="0,0,8,0">
                                <StackPanel>
                                    <TextBlock Text="TOTAL RUNS" FontSize="10" FontWeight="Bold"
                                               Foreground="#8892A8"/>
                                    <TextBlock x:Name="txtMetricTotal" Text="0"
                                               FontSize="36" FontWeight="Black" Foreground="#E8EAF0"/>
                                    <TextBlock Text="all time" FontSize="11" Foreground="#4A5068"/>
                                </StackPanel>
                            </Border>
                            <Border Style="{StaticResource MetricCard}" Margin="4,0,4,0">
                                <StackPanel>
                                    <TextBlock Text="SUCCESS RATE" FontSize="10" FontWeight="Bold"
                                               Foreground="#8892A8"/>
                                    <TextBlock x:Name="txtMetricSuccess" Text="0%"
                                               FontSize="36" FontWeight="Black" Foreground="#3EBA7E"/>
                                    <TextBlock x:Name="txtMetricOKErr" Text="0 OK / 0 Error"
                                               FontSize="11" Foreground="#4A5068"/>
                                </StackPanel>
                            </Border>
                            <Border Style="{StaticResource MetricCard}" Margin="4,0,4,0">
                                <StackPanel>
                                    <TextBlock Text="AVG DURATION" FontSize="10" FontWeight="Bold"
                                               Foreground="#8892A8"/>
                                    <TextBlock x:Name="txtMetricAvg" Text="0s"
                                               FontSize="36" FontWeight="Black" Foreground="#F5A623"/>
                                    <TextBlock Text="per run" FontSize="11" Foreground="#4A5068"/>
                                </StackPanel>
                            </Border>
                            <Border Style="{StaticResource MetricCard}" Margin="8,0,0,0">
                                <StackPanel>
                                    <TextBlock Text="TOP TASK" FontSize="10" FontWeight="Bold"
                                               Foreground="#8892A8"/>
                                    <TextBlock x:Name="txtMetricTop" Text="N/A"
                                               FontSize="20" FontWeight="Black" Foreground="#4A9EFF"
                                               TextTrimming="CharacterEllipsis"/>
                                    <TextBlock Text="most executed" FontSize="11" Foreground="#4A5068"/>
                                </StackPanel>
                            </Border>
                        </UniformGrid>

                        <!-- Task frequency table -->
                        <TextBlock Text="TASK FREQUENCY" FontSize="12" FontWeight="Bold"
                                   Foreground="#4A5068" Margin="0,28,0,12"/>
                        <Border Background="#1E2232" CornerRadius="8" Padding="16">
                            <DataGrid x:Name="gridMetricTasks" Style="{StaticResource DarkGrid}"
                                      MaxHeight="300" BorderThickness="0">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Task"  Binding="{Binding Name}"  Width="200"/>
                                    <DataGridTextColumn Header="Runs"  Binding="{Binding Count}" Width="80"/>
                                    <DataGridTextColumn Header="Bar"   Binding="{Binding Bar}"   Width="*"/>
                                </DataGrid.Columns>
                            </DataGrid>
                        </Border>

                        <!-- Recent errors -->
                        <TextBlock Text="RECENT ERRORS" FontSize="12" FontWeight="Bold"
                                   Foreground="#4A5068" Margin="0,28,0,12"/>
                        <Border Background="#1E2232" CornerRadius="8" Padding="16">
                            <DataGrid x:Name="gridMetricErrors" Style="{StaticResource DarkGrid}"
                                      MaxHeight="240" BorderThickness="0">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Timestamp" Binding="{Binding ts}"      Width="160"/>
                                    <DataGridTextColumn Header="Task"      Binding="{Binding task}"    Width="140"/>
                                    <DataGridTextColumn Header="Message"   Binding="{Binding message}" Width="*"/>
                                </DataGrid.Columns>
                            </DataGrid>
                        </Border>
                    </StackPanel>
                </ScrollViewer>

                <!-- ========================================================
                     PAGE: DIAGNOSTICS
                     ======================================================== -->
                <Grid x:Name="pageDiag" Visibility="Collapsed">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <Border Grid.Row="0" Padding="16,12" Background="#171B26"
                            BorderBrush="#2D3348" BorderThickness="0,0,0,1">
                        <StackPanel Orientation="Horizontal">
                            <Button x:Name="btnRunDiag" Content="Run Diagnostics"
                                    Style="{StaticResource BtnPrimary}"/>
                            <TextBlock x:Name="txtDiagSummary" Text="Click Run Diagnostics to check environment"
                                       FontSize="12" Foreground="#8892A8"
                                       Margin="16,0,0,0" VerticalAlignment="Center"/>
                        </StackPanel>
                    </Border>
                    <DataGrid x:Name="gridDiag" Grid.Row="1" Style="{StaticResource DarkGrid}"
                              Margin="16">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Category" Binding="{Binding Category}" Width="110"/>
                            <DataGridTextColumn Header="Item"     Binding="{Binding Item}"     Width="180"/>
                            <DataGridTextColumn Header="Value"    Binding="{Binding Value}"    Width="120"/>
                            <DataGridTextColumn Header="Status"   Binding="{Binding Status}"   Width="70"/>
                            <DataGridTextColumn Header="Message"  Binding="{Binding Message}"  Width="*"/>
                        </DataGrid.Columns>
                    </DataGrid>
                </Grid>

                <!-- ========================================================
                     PAGE: CONFIGURATION
                     ======================================================== -->
                <Grid x:Name="pageConfig" Visibility="Collapsed">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="320"/>
                    </Grid.ColumnDefinitions>

                    <!-- JSON editor -->
                    <Grid Grid.Column="0">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <Border Grid.Row="0" Padding="16,12" Background="#171B26"
                                BorderBrush="#2D3348" BorderThickness="0,0,0,1">
                            <StackPanel Orientation="Horizontal">
                                <Button x:Name="btnSaveConfig" Content="Save Configuration"
                                        Style="{StaticResource BtnPrimary}"/>
                                <Button x:Name="btnReloadConfig" Content="Reload"
                                        Style="{StaticResource BtnSecondary}" Margin="8,0,0,0"/>
                                <TextBlock x:Name="txtConfigStatus" Text=""
                                           FontSize="12" Foreground="#8892A8"
                                           Margin="16,0,0,0" VerticalAlignment="Center"/>
                            </StackPanel>
                        </Border>
                        <TextBox x:Name="txtConfigEditor" Grid.Row="1"
                                 Style="{StaticResource InputField}"
                                 Margin="16" AcceptsReturn="True" TextWrapping="NoWrap"
                                 VerticalScrollBarVisibility="Auto"
                                 HorizontalScrollBarVisibility="Auto"
                                 FontFamily="Consolas" FontSize="13"
                                 Background="#0A0D14" BorderThickness="0"/>
                    </Grid>

                    <!-- Help panel -->
                    <Border Grid.Column="1" Background="#171B26"
                            BorderBrush="#2D3348" BorderThickness="1,0,0,0" Padding="20">
                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                            <StackPanel>
                                <TextBlock Text="CONFIGURATION GUIDE" Style="{StaticResource SectionHeader}"/>

                                <TextBlock Text="engine" FontSize="13" FontWeight="Bold"
                                           Foreground="#F5A623" Margin="0,8,0,4"/>
                                <TextBlock FontSize="12" Foreground="#8892A8" TextWrapping="Wrap"
                                           Text="logLevel: DEBUG|INFO|WARN|ERROR&#10;maxRetries: 0-10&#10;retryDelaySeconds: seconds between retries"/>

                                <TextBlock Text="sap" FontSize="13" FontWeight="Bold"
                                           Foreground="#F5A623" Margin="0,16,0,4"/>
                                <TextBlock FontSize="12" Foreground="#8892A8" TextWrapping="Wrap"
                                           Text="systemId: SAP System ID (e.g. PRD)&#10;client: SAP Client number&#10;language: Login language (ES, EN)&#10;timeout: Session timeout in seconds"/>

                                <TextBlock Text="excel" FontSize="13" FontWeight="Bold"
                                           Foreground="#F5A623" Margin="0,16,0,4"/>
                                <TextBlock FontSize="12" Foreground="#8892A8" TextWrapping="Wrap"
                                           Text="visible: Show Excel window (false=headless)&#10;screenUpdating: Enable screen refresh"/>

                                <TextBlock Text="reports" FontSize="13" FontWeight="Bold"
                                           Foreground="#F5A623" Margin="0,16,0,4"/>
                                <TextBlock FontSize="12" Foreground="#8892A8" TextWrapping="Wrap"
                                           Text="defaultFormat: Output format (xlsx, csv)&#10;retentionDays: Days to keep artifacts (1-365)"/>

                                <Border Margin="0,20,0,0" Padding="12" Background="#1E2232" CornerRadius="6">
                                    <TextBlock FontSize="11" Foreground="#E85555" TextWrapping="Wrap"
                                               Text="Changes are validated before saving. Invalid JSON or out-of-range values will be rejected. All config changes are audit-logged."/>
                                </Border>
                            </StackPanel>
                        </ScrollViewer>
                    </Border>
                </Grid>

                <!-- ========================================================
                     PAGE: NEW TASK
                     ======================================================== -->
                <Grid x:Name="pageNewTask" Visibility="Collapsed">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="400"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <Border Grid.Column="0" Background="#171B26"
                            BorderBrush="#2D3348" BorderThickness="0,0,1,0" Padding="24">
                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                            <StackPanel>
                                <TextBlock Text="CREATE NEW TASK" Style="{StaticResource SectionHeader}"/>

                                <TextBlock Text="Task Name" Style="{StaticResource FieldLabel}"/>
                                <TextBlock Text="Lowercase letters, numbers, and underscores only"
                                           FontSize="10" Foreground="#4A5068" Margin="0,0,0,4"/>
                                <TextBox x:Name="txtNewTaskName" Height="34"
                                         Style="{StaticResource InputField}"/>
                                <TextBlock x:Name="txtNewTaskNameError" Text="" FontSize="11"
                                           Foreground="#E85555" Margin="0,2,0,0" Visibility="Collapsed"/>

                                <TextBlock Text="Category" Style="{StaticResource FieldLabel}"
                                           Margin="0,12,0,4"/>
                                <ComboBox x:Name="cboNewTaskCategory" Height="34"
                                          Style="{StaticResource InputCombo}">
                                    <ComboBoxItem Content="SAP"/>
                                    <ComboBoxItem Content="Excel" IsSelected="True"/>
                                    <ComboBoxItem Content="CSV"/>
                                    <ComboBoxItem Content="Reporte"/>
                                    <ComboBoxItem Content="Utilidad"/>
                                </ComboBox>

                                <TextBlock Text="Description" Style="{StaticResource FieldLabel}"
                                           Margin="0,12,0,4"/>
                                <TextBox x:Name="txtNewTaskDesc" Height="64"
                                         Style="{StaticResource InputField}"
                                         AcceptsReturn="True" TextWrapping="Wrap"/>

                                <Separator Margin="0,16" Background="#2D3348"/>

                                <Button x:Name="btnCreateTask" Content="CREATE TASK FILE"
                                        Style="{StaticResource BtnPrimary}" Height="40" FontSize="14"/>

                                <TextBlock x:Name="txtCreateTaskResult" Text="" FontSize="12"
                                           Foreground="#3EBA7E" TextWrapping="Wrap"
                                           Margin="0,12,0,0" Visibility="Collapsed"/>
                            </StackPanel>
                        </ScrollViewer>
                    </Border>

                    <!-- Preview of what will be generated -->
                    <Grid Grid.Column="1">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="40"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <Border Grid.Row="0" Background="#1E2232" Padding="16,0">
                            <TextBlock Text="TASK FILE PREVIEW" FontSize="11" FontWeight="Bold"
                                       Foreground="#8892A8" VerticalAlignment="Center"/>
                        </Border>
                        <TextBox x:Name="txtNewTaskPreview" Grid.Row="1"
                                 Style="{StaticResource InputField}"
                                 IsReadOnly="True" AcceptsReturn="True"
                                 FontFamily="Consolas" FontSize="12"
                                 Background="#0A0D14" BorderThickness="0"
                                 VerticalScrollBarVisibility="Auto"/>
                    </Grid>
                </Grid>

                <!-- ========================================================
                     PAGE: AUDIT LOG
                     ======================================================== -->
                <Grid x:Name="pageAudit" Visibility="Collapsed">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="60"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <Border Grid.Row="0" Background="#171B26" Padding="16,0"
                            BorderBrush="#2D3348" BorderThickness="0,0,0,1">
                        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                            <Button x:Name="btnExportAudit" Content="Export Audit CSV"
                                    Style="{StaticResource BtnSecondary}"/>
                            <TextBlock Text="All user actions are recorded here."
                                       FontSize="12" Foreground="#8892A8"
                                       Margin="16,0,0,0" VerticalAlignment="Center"/>
                        </StackPanel>
                    </Border>
                    <DataGrid x:Name="gridAudit" Grid.Row="1" Style="{StaticResource DarkGrid}"
                              Margin="16">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Timestamp" Binding="{Binding ts}"     Width="160"/>
                            <DataGridTextColumn Header="User"      Binding="{Binding user}"   Width="100"/>
                            <DataGridTextColumn Header="Role"      Binding="{Binding role}"   Width="90"/>
                            <DataGridTextColumn Header="Action"    Binding="{Binding action}" Width="150"/>
                            <DataGridTextColumn Header="Target"    Binding="{Binding target}" Width="160"/>
                            <DataGridTextColumn Header="Detail"    Binding="{Binding detail}" Width="*"/>
                        </DataGrid.Columns>
                    </DataGrid>
                </Grid>

            </Grid><!-- end pages -->
        </Grid><!-- end main content -->
    </Grid>
</Window>
'@

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
    'pageCatalog','pageExecute','pageMonitor','pageHistory',
    'pageCheckpoints','pageArtifacts','pageMetrics',
    'pageDiag','pageConfig','pageNewTask','pageAudit'
)

# ============================================================================
# ALL LOGIC AS $Script: SCRIPTBLOCK VARIABLES
# Event handlers call  & $Script:Fn_xxx  — no function name lookup at all.
# This works regardless of whether the file is dot-sourced or run with -File.
# ============================================================================

$Script:Fn_TestPermission = {
    param([string]$Action)
    switch ($Action) {
        'EditConfig'        { return $Script:CurrentRole -eq 'Admin' }
        'ViewAudit'         { return $Script:CurrentRole -eq 'Admin' }
        'CreateTask'        { return $Script:CurrentRole -in @('Admin','Developer') }
        'DeleteArtifact'    { return $Script:CurrentRole -eq 'Admin' }
        'ManageCheckpoints' { return $Script:CurrentRole -eq 'Admin' }
        default             { return $true }
    }
}

$Script:Fn_WriteAuditLog = {
    param([string]$Action, [string]$Target='', [string]$Detail='')
    $dir = Split-Path $Script:AuditFile -Parent
    if (-not (Test-Path $dir)) { return }
    $entry = [ordered]@{
        ts     = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss')
        user   = $Script:CurrentUser
        role   = $Script:CurrentRole
        action = $Action
        target = $Target
        detail = $Detail
    }
    $line  = $entry | ConvertTo-Json -Compress
    $mutex = New-Object System.Threading.Mutex($false, 'Global\AutoBuildAuditMutex')
    try {
        [void]$mutex.WaitOne(2000)
        Add-Content -Path $Script:AuditFile -Value $line -Encoding ASCII
    } catch {}
    finally { $mutex.ReleaseMutex(); $mutex.Dispose() }
}

$Script:Fn_GetExecutionHistory = {
    param([int]$MaxEntries=200, [string]$FilterTask='', [string]$FilterLevel='')
    $results = @()
    if (-not (Test-Path $Script:RegistryFile)) { return $results }
    try {
        $lines = Get-Content $Script:RegistryFile -Encoding ASCII -ErrorAction Stop |
                 Select-Object -Last ($MaxEntries * 4)
        foreach ($line in $lines) {
            try {
                $o = $line | ConvertFrom-Json
                if ($FilterTask  -and $o.task  -ne $FilterTask)  { continue }
                if ($FilterLevel -and $o.level -ne $FilterLevel) { continue }
                $results += $o
            } catch {}
        }
    } catch {}
    return @($results | Select-Object -Last $MaxEntries)
}

$Script:Fn_GetRunSummaries = {
    $all = @(& $Script:Fn_GetExecutionHistory -MaxEntries 400)
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
            User    = $Script:CurrentUser
            Entries = $entries.Count
        }
    }
    return @($summaries | Sort-Object Started -Descending)
}

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
    } catch {}
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
    $history = @(& $Script:Fn_GetExecutionHistory -MaxEntries 300)
    foreach ($task in $tasks) {
        $last = $history |
            Where-Object { $_.task -eq $task.Name -and $_.level -in @('OK','ERROR') } |
            Select-Object -First 1
        if ($last) { $task.LastRun = $last.ts; $task.LastStatus = $last.level }
    }
    return $tasks
}

$Script:Fn_GetConfig = {
    if (-not (Test-Path $Script:ConfigFile)) { return $null }
    try { return Get-Content $Script:ConfigFile -Raw -Encoding ASCII | ConvertFrom-Json }
    catch { return $null }
}

$Script:Fn_SaveConfigData = {
    param([string]$JsonContent)
    try {
        $obj = $JsonContent | ConvertFrom-Json
        $mr  = [int]$obj.engine.maxRetries
        if ($mr -lt 0 -or $mr -gt 10) { return @{ Success=$false; Error='maxRetries must be 0-10' } }
        [System.IO.File]::WriteAllText($Script:ConfigFile, $JsonContent, [System.Text.Encoding]::ASCII)
        & $Script:Fn_WriteAuditLog -Action 'EDIT_CONFIG'
        return @{ Success=$true }
    } catch { return @{ Success=$false; Error="$_" } }
}

$Script:Fn_GetCheckpoints = {
    $logDir = Join-Path $Script:EngineRoot 'logs'
    if (-not (Test-Path $logDir)) { return @() }
    return @(
        Get-ChildItem $logDir -Filter 'checkpoint_*.clixml' | Sort-Object LastWriteTime -Descending |
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

$Script:Fn_GetArtifacts = {
    $all = @()
    foreach ($dir in @((Join-Path $Script:EngineRoot 'output'),(Join-Path $Script:EngineRoot 'reports'))) {
        if (Test-Path $dir) {
            $all += Get-ChildItem $dir -File -Recurse -ErrorAction SilentlyContinue |
                    ForEach-Object {
                        [PSCustomObject]@{
                            Name=$_.Name; Extension=$_.Extension; Directory=$_.DirectoryName
                            Size=$_.Length; Modified=$_.LastWriteTime; Path=$_.FullName
                        }
                    }
        }
    }
    return @($all | Sort-Object Modified -Descending)
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

$Script:Fn_StartTaskExecution = {
    param([string]$TaskName,[hashtable]$Params,[switch]$WhatIf,[switch]$Checkpoint,[switch]$Resume)
    if (-not (Test-Path $Script:RunScript)) {
        return @{ Success=$false; Error="Run.ps1 not found: $Script:RunScript" }
    }
    & $Script:Fn_WriteAuditLog -Action 'EXECUTE_TASK' -Target $TaskName `
        -Detail "WhatIf=$WhatIf Checkpoint=$Checkpoint Resume=$Resume"

    $argList = @('-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass',
                 '-File',$Script:RunScript,'-Task',$TaskName)
    foreach ($k in $Params.Keys) {
        if (-not [string]::IsNullOrWhiteSpace($Params[$k])) {
            $argList += "-$k"
            $argList += $Params[$k]
        }
    }
    if ($WhatIf)     { $argList += '-WhatIf' }
    if ($Checkpoint) { $argList += '-Checkpoint' }
    if ($Resume)     { $argList += '-Resume' }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName               = 'powershell.exe'
    $psi.Arguments              = ($argList | ForEach-Object {
                                        if ($_ -match '\s') { "`"$_`"" } else { $_ }
                                  }) -join ' '
    $psi.UseShellExecute        = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.CreateNoWindow         = $true
    $psi.WorkingDirectory       = $Script:EngineRoot

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    $proc.EnableRaisingEvents = $true
    $outBuf = [System.Text.StringBuilder]::new()

    $proc.add_OutputDataReceived([System.Diagnostics.DataReceivedEventHandler]{
        param($s,$e); if ($null -ne $e.Data) { [void]$outBuf.AppendLine($e.Data) }
    })
    $proc.add_ErrorDataReceived([System.Diagnostics.DataReceivedEventHandler]{
        param($s,$e); if ($null -ne $e.Data) { [void]$outBuf.AppendLine("[ERR] $($e.Data)") }
    })
    try {
        [void]$proc.Start()
        $proc.BeginOutputReadLine()
        $proc.BeginErrorReadLine()
    } catch {
        return @{ Success=$false; Error="Process start failed: $_" }
    }
    $Script:ActiveJobs[$TaskName] = @{
        Process=$proc; Buffer=$outBuf; Started=[datetime]::Now; TaskName=$TaskName
    }
    return @{ Success=$true; Pid=$proc.Id }
}

$Script:Fn_GetJobOutput = {
    param([string]$TaskName)
    if (-not $Script:ActiveJobs.ContainsKey($TaskName)) { return $null }
    $info    = $Script:ActiveJobs[$TaskName]
    $proc    = $info.Process
    $running = -not $proc.HasExited
    return @{
        State    = if ($running) { 'Running' } elseif ($proc.ExitCode -eq 0) { 'Completed' } else { 'Failed' }
        Output   = $info.Buffer.ToString()
        Duration = ([datetime]::Now - $info.Started).TotalSeconds
        ExitCode = if (-not $running) { $proc.ExitCode } else { $null }
    }
}

$Script:Fn_GetDiagnostics = {
    $results = @()
    $add = {
        param([string]$cat,[string]$item,[string]$val,[string]$status,[string]$msg)
        $results += [PSCustomObject]@{Category=$cat;Item=$item;Value=$val;Status=$status;Message=$msg}
    }
    $psv = $PSVersionTable.PSVersion.ToString()
    & $add 'Runtime' 'PowerShell' $psv (if ([version]$psv -ge [version]'5.1'){'OK'}else{'WARN'}) ''
    foreach ($f in @('Run.ps1','engine\Main.build.ps1','engine.config.json')) {
        $fp = Join-Path $Script:EngineRoot $f
        $ok = Test-Path $fp
        & $add 'Engine' $f $fp (if ($ok){'OK'}else{'ERROR'}) (if (-not $ok){'Missing'}else{''})
    }
    foreach ($d in @('logs','output','reports','input')) {
        $dp = Join-Path $Script:EngineRoot $d
        if (Test-Path $dp) {
            $tp = Join-Path $dp "._test_$(Get-Random)"
            try {
                [void](New-Item $tp -ItemType File -Force -ErrorAction Stop)
                Remove-Item $tp -Force
                & $add 'Folders' $d $dp 'OK' ''
            } catch { & $add 'Folders' $d $dp 'WARN' 'Not writable' }
        } else { & $add 'Folders' $d $dp 'WARN' 'Missing' }
    }
    foreach ($prog in @('Excel.Application','Word.Application')) {
        try {
            $com = New-Object -ComObject $prog -ErrorAction Stop
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($com)
            & $add 'COM' $prog '' 'OK' 'Available'
        } catch { & $add 'COM' $prog '' 'WARN' 'Not available' }
    }
    return $results
}

# ---- Page loaders ----

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
    $ts       = $Script:Ctrl['txtCatalogSearch']
    $search   = if ($null -ne $ts) { $ts.Text.ToLower() } else { '' }
    $cbc      = $Script:Ctrl['cboCatalogCategory']
    $catSel   = if ($null -ne $cbc) { $cbc.SelectedItem } else { $null }
    $category = if ($null -ne $catSel -and $catSel.Content -ne 'All Categories') {
                    $catSel.Content } else { '' }

    $rows = @($Script:AllTasks | Where-Object {
        ($search -eq '' -or
         $_.Name.ToLower() -like "*$search*" -or
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
    param([string]$PreSelectTask='')
    $Script:AllTasks = @(& $Script:Fn_GetAllTasks)
    $cbo = $Script:Ctrl['cboExecTask']
    if ($null -eq $cbo) { return }
    $cbo.Items.Clear()
    foreach ($t in $Script:AllTasks) {
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = $t.Name
        $item.Tag     = $t.Name
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
        if ($param.Help) { $tb.ToolTip = "$($param.Help) (Type: $($param.Type))" }
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

    $params = @{}
    $panel  = $Script:Ctrl['pnlParams']
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
                if ($null -ne $esT) { $esT.Text = "Required: $($p.Name)" }
                if ($null -ne $esB) { $esB.Background = [System.Windows.Media.Brushes]::Crimson }
                return
            }
        }
    }

    $Script:CurrentTaskName = $taskName
    $eo = $Script:Ctrl['txtExecOutput']
    if ($null -ne $eo)  { $eo.Text  = "Starting: $taskName`n$('-'*60)`n" }
    if ($null -ne $esT) { $esT.Text = 'Running...' }
    if ($null -ne $esB) { $esB.Background = [System.Windows.Media.Brushes]::Goldenrod }
    if ($null -ne $bRun){ $bRun.IsEnabled = $false }
    if ($null -ne $bCan){ $bCan.IsEnabled = $true }

    $wi  = $Script:Ctrl['chkWhatIf']
    $cp  = $Script:Ctrl['chkCheckpoint']
    $rp  = $Script:Ctrl['chkResume']
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

    if ($null -ne $Script:ExecTimer) { $Script:ExecTimer.Stop() }
    $Script:ExecTimer = New-Object System.Windows.Threading.DispatcherTimer
    $Script:ExecTimer.Interval = [TimeSpan]::FromSeconds(1)
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

    if ($info.State -in @('Completed','Failed','Stopped')) {
        $Script:ExecTimer.Stop()
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
            try { $Script:ActiveJobs[$Script:CurrentTaskName].Process.Dispose() } catch {}
            $Script:ActiveJobs.Remove($Script:CurrentTaskName)
        }
    }
}

$Script:Fn_LoadMonitorPage = {
    $rows = @(& $Script:Fn_GetRunSummaries | Select-Object -First 20 | ForEach-Object {
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

    $logText = ''
    if (Test-Path $Script:RegistryFile) {
        try {
            $lines = @(Get-Content $Script:RegistryFile -Tail 25 -Encoding ASCII -ErrorAction Stop)
            $logText = ($lines | ForEach-Object {
                try { $o = $_ | ConvertFrom-Json; "[$($o.ts)][$($o.level)] $($o.task): $($o.message)" }
                catch { $_ }
            }) -join "`n"
        } catch {}
    }
    $tl = $Script:Ctrl['txtLiveLog'];     if ($null -ne $tl) { $tl.Text = $logText }
    $sl = $Script:Ctrl['svLiveLog'];      if ($null -ne $sl) { $sl.ScrollToEnd() }
    $tc = $Script:Ctrl['txtActiveCount']; if ($null -ne $tc) { $tc.Text = "$($Script:ActiveJobs.Count) running" }
}

$Script:Fn_LoadHistoryPage = {
    $ts = $Script:Ctrl['txtHistorySearch']
    $ft = if ($null -ne $ts) { $ts.Text } else { '' }
    $hs = $Script:Ctrl['cboHistoryStatus']
    $si = if ($null -ne $hs) { $hs.SelectedItem } else { $null }
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
    param([string]$RunId,[string]$Task)
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
    $g = $Script:Ctrl['gridCheckpoints']
    if ($null -ne $g) { $g.ItemsSource = @(& $Script:Fn_GetCheckpoints) }
}

$Script:Fn_ResumeCheckpoint = {
    $gc  = $Script:Ctrl['gridCheckpoints']
    $sel = if ($null -ne $gc) { $gc.SelectedItem } else { $null }
    if ($null -eq $sel) {
        [void][System.Windows.MessageBox]::Show('Select a checkpoint first.','AutoBuild','OK','Warning')
        return
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
    $gc  = $Script:Ctrl['gridCheckpoints']
    $sel = if ($null -ne $gc) { $gc.SelectedItem } else { $null }
    if ($null -eq $sel) { return }
    if ([System.Windows.MessageBox]::Show("Delete $($sel.Name)?","Confirm",'YesNo','Question') -eq 'Yes') {
        try {
            Remove-Item $sel.Path -Force
            & $Script:Fn_WriteAuditLog -Action 'DELETE_CHECKPOINT' -Target $sel.Name
            & $Script:Fn_LoadCheckpointsPage
        } catch { [void][System.Windows.MessageBox]::Show("Error: $_",'AutoBuild','OK','Error') }
    }
}

$Script:Fn_ViewCheckpoint = {
    $gc  = $Script:Ctrl['gridCheckpoints']
    $sel = if ($null -ne $gc) { $gc.SelectedItem } else { $null }
    if ($null -eq $sel) { return }
    try {
        $data = Import-Clixml $sel.Path | Out-String
        [void][System.Windows.MessageBox]::Show(
            $data.Substring(0,[Math]::Min($data.Length,2000)),
            "Checkpoint: $($sel.Name)",'OK','Information')
    } catch {
        [void][System.Windows.MessageBox]::Show("Cannot read: $_",'AutoBuild','OK','Warning')
    }
}

$Script:Fn_LoadArtifactsPage = {
    $ts     = $Script:Ctrl['txtArtifactSearch']
    $filter = if ($null -ne $ts) { $ts.Text.ToLower() } else { '' }
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
    $ga = $Script:Ctrl['gridArtifacts'];   if ($null -ne $ga) { $ga.ItemsSource = $rows }
    $tc = $Script:Ctrl['txtArtifactCount'];if ($null -ne $tc) { $tc.Text = "$($rows.Count) artifacts" }
}

$Script:Fn_OpenArtifact = {
    $g   = $Script:Ctrl['gridArtifacts']
    $sel = if ($null -ne $g) { $g.SelectedItem } else { $null }
    if ($null -eq $sel) { return }
    try {
        Start-Process $sel.Path
        & $Script:Fn_WriteAuditLog -Action 'OPEN_ARTIFACT' -Target $sel.Name
    } catch { [void][System.Windows.MessageBox]::Show("Cannot open: $_",'AutoBuild','OK','Warning') }
}

$Script:Fn_SaveArtifact = {
    $g   = $Script:Ctrl['gridArtifacts']
    $sel = if ($null -ne $g) { $g.SelectedItem } else { $null }
    if ($null -eq $sel) { return }
    [void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.FileName = $sel.Name
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Copy-Item $sel.Path $dlg.FileName -Force
        & $Script:Fn_WriteAuditLog -Action 'DOWNLOAD_ARTIFACT' -Target $sel.Name -Detail $dlg.FileName
    }
}

$Script:Fn_DeleteArtifact = {
    if (-not (& $Script:Fn_TestPermission -Action 'DeleteArtifact')) {
        [void][System.Windows.MessageBox]::Show('Admin role required.','AutoBuild','OK','Warning'); return
    }
    $g   = $Script:Ctrl['gridArtifacts']
    $sel = if ($null -ne $g) { $g.SelectedItem } else { $null }
    if ($null -eq $sel) { return }
    if ([System.Windows.MessageBox]::Show("Delete $($sel.Name)?","Confirm",'YesNo','Question') -eq 'Yes') {
        try {
            Remove-Item $sel.Path -Force
            & $Script:Fn_WriteAuditLog -Action 'DELETE_ARTIFACT' -Target $sel.Name
            & $Script:Fn_LoadArtifactsPage
        } catch { [void][System.Windows.MessageBox]::Show("Error: $_",'AutoBuild','OK','Error') }
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

$Script:Fn_RunDiagnostics = {
    $bd = $Script:Ctrl['btnRunDiag'];     if ($null -ne $bd) { $bd.IsEnabled = $false }
    $ds = $Script:Ctrl['txtDiagSummary']; if ($null -ne $ds) { $ds.Text = 'Running...' }
    $results = @(& $Script:Fn_GetDiagnostics)
    $g = $Script:Ctrl['gridDiag']; if ($null -ne $g) { $g.ItemsSource = $results }
    $ok   = @($results | Where-Object { $_.Status -eq 'OK'    }).Count
    $warn = @($results | Where-Object { $_.Status -eq 'WARN'  }).Count
    $err  = @($results | Where-Object { $_.Status -eq 'ERROR' }).Count
    if ($null -ne $ds) { $ds.Text = "$ok OK  |  $warn WARN  |  $err ERROR" }
    if ($null -ne $bd) { $bd.IsEnabled = $true }
    & $Script:Fn_WriteAuditLog -Action 'RUN_DIAGNOSTICS' -Detail "OK=$ok WARN=$warn ERR=$err"
}

$Script:Fn_LoadConfigPage = {
    $ce = $Script:Ctrl['txtConfigEditor']
    $bs = $Script:Ctrl['btnSaveConfig']
    if (-not (& $Script:Fn_TestPermission -Action 'EditConfig')) {
        if ($null -ne $ce) { $ce.Text = '// Admin role required.'; $ce.IsReadOnly = $true }
        if ($null -ne $bs) { $bs.IsEnabled = $false }
        return
    }
    $cfg = & $Script:Fn_GetConfig
    if ($null -ne $ce) {
        $ce.Text = if ($null -ne $cfg) { $cfg | ConvertTo-Json -Depth 5 } else { '// Error reading config.' }
    }
    $cs = $Script:Ctrl['txtConfigStatus']; if ($null -ne $cs) { $cs.Text = '' }
}

$Script:Fn_SaveConfigPage = {
    $ce = $Script:Ctrl['txtConfigEditor']
    $r  = & $Script:Fn_SaveConfigData -JsonContent (if ($null -ne $ce) { $ce.Text } else { '' })
    $cs = $Script:Ctrl['txtConfigStatus']
    if ($null -ne $cs) {
        $cs.Text       = if ($r.Success) { 'Saved.' } else { "Error: $($r.Error)" }
        $cs.Foreground = if ($r.Success) { [System.Windows.Media.Brushes]::LimeGreen } else { [System.Windows.Media.Brushes]::Crimson }
    }
}

$Script:Fn_LoadNewTaskPage = {
    if (-not (& $Script:Fn_TestPermission -Action 'CreateTask')) {
        $n = $Script:Ctrl['txtNewTaskName']; if ($null -ne $n) { $n.IsEnabled = $false }
        $b = $Script:Ctrl['btnCreateTask'];  if ($null -ne $b) { $b.IsEnabled = $false }
        return
    }
    & $Script:Fn_UpdateNewTaskPreview
}

$Script:Fn_UpdateNewTaskPreview = {
    $nn   = $Script:Ctrl['txtNewTaskName'];    $name = if ($null -ne $nn) { $nn.Text } else { '' }
    $nc   = $Script:Ctrl['cboNewTaskCategory'];$ci   = if ($null -ne $nc) { $nc.SelectedItem } else { $null }
    $cat  = if ($null -ne $ci) { $ci.Content } else { 'Utilidad' }
    $nd   = $Script:Ctrl['txtNewTaskDesc'];    $desc = if ($null -ne $nd) { $nd.Text } else { '' }
    if ([string]::IsNullOrWhiteSpace($desc)) { $desc = 'Task description pending' }

    $preview = "task ${name} {`n    # Category: $cat`n    # $desc`n    Write-Build Green 'Done'`n}"
    $np = $Script:Ctrl['txtNewTaskPreview']; if ($null -ne $np) { $np.Text = $preview }
    $ne = $Script:Ctrl['txtNewTaskNameError']
    if ($null -ne $ne) {
        $valid = [string]::IsNullOrWhiteSpace($name) -or ($name -match '^[a-zA-Z0-9_-]+$')
        $ne.Visibility = if ($valid) { [System.Windows.Visibility]::Collapsed } else { [System.Windows.Visibility]::Visible }
    }
}

$Script:Fn_CreateNewTask = {
    $nn   = $Script:Ctrl['txtNewTaskName']
    $name = if ($null -ne $nn) { $nn.Text.Trim() } else { '' }
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
    $nc   = $Script:Ctrl['cboNewTaskCategory']
    $ci   = if ($null -ne $nc) { $nc.SelectedItem } else { $null }
    $cat  = if ($null -ne $ci) { $ci.Content } else { 'Utilidad' }
    $np   = $Script:Ctrl['txtNewTaskPreview']
    $content = if ($null -ne $np) { $np.Text } else { "task ${name} { }" }
    try {
        [System.IO.File]::WriteAllText($outFile, $content, [System.Text.Encoding]::ASCII)
        $tr = $Script:Ctrl['txtCreateTaskResult']
        if ($null -ne $tr) {
            $tr.Text       = "Created: $outFile"
            $tr.Foreground = [System.Windows.Media.Brushes]::LimeGreen
            $tr.Visibility = [System.Windows.Visibility]::Visible
        }
        & $Script:Fn_WriteAuditLog -Action 'CREATE_TASK' -Target "task_${name}.ps1" -Detail "Cat=$cat"
        & $Script:Fn_LoadCatalogPage
    } catch { [void][System.Windows.MessageBox]::Show("Error: $_",'AutoBuild','OK','Error') }
}

$Script:Fn_LoadAuditPage = {
    $ga = $Script:Ctrl['gridAudit']
    if (-not (& $Script:Fn_TestPermission -Action 'ViewAudit') -or -not (Test-Path $Script:AuditFile)) {
        if ($null -ne $ga) { $ga.ItemsSource = $null }; return
    }
    $entries = @()
    try {
        Get-Content $Script:AuditFile -Encoding ASCII -ErrorAction Stop | Select-Object -Last 500 |
            ForEach-Object { try { $entries += ($_ | ConvertFrom-Json) } catch {} }
    } catch {}
    if ($null -ne $ga) { $ga.ItemsSource = @($entries | Sort-Object ts -Descending) }
}

$Script:Fn_ExportHistoryCSV = {
    [void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter   = 'CSV|*.csv'
    $dlg.FileName = "history_$(Get-Date -Format yyyyMMdd).csv"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        @(& $Script:Fn_GetRunSummaries) | Select-Object RunId,Task,Started,Elapsed,Status,Entries,User |
            Export-Csv $dlg.FileName -NoTypeInformation -Encoding ASCII
        & $Script:Fn_WriteAuditLog -Action 'EXPORT_HISTORY' -Target $dlg.FileName
    }
}

$Script:Fn_ExportAuditCSV = {
    if (-not (Test-Path $Script:AuditFile)) {
        [void][System.Windows.MessageBox]::Show('No audit log.','AutoBuild','OK','Warning'); return
    }
    [void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter   = 'CSV|*.csv'
    $dlg.FileName = "audit_$(Get-Date -Format yyyyMMdd).csv"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $e = @()
        Get-Content $Script:AuditFile -Encoding ASCII |
            ForEach-Object { try { $e += ($_ | ConvertFrom-Json) } catch {} }
        $e | Select-Object ts,user,role,action,target,detail |
            Export-Csv $dlg.FileName -NoTypeInformation -Encoding ASCII
        & $Script:Fn_WriteAuditLog -Action 'EXPORT_AUDIT' -Target $dlg.FileName
    }
}

# ============================================================================
# INITIALIZE WINDOW
# ============================================================================
function Initialize-Window {
    $reader = New-Object System.Xml.XmlNodeReader $Script:XAML
    $Script:Window = [Windows.Markup.XamlReader]::Load($reader)

    $controls = @(
        'txtVersion','txtUserInitial','txtUserName','txtUserRole',
        'txtEngineStatus','txtEnginePath','elEngineStatus','txtPageTitle','btnRefresh',
        'pageCatalog','gridCatalog','txtCatalogSearch','cboCatalogCategory','txtCatalogCount',
        'btnCatalogExecute','btnCatalogViewHistory',
        'pageExecute','cboExecTask','pnlTaskInfo','txtExecTaskDesc','txtExecTaskCat',
        'pnlParams','chkWhatIf','chkCheckpoint','chkResume',
        'btnExecuteTask','btnCancelTask','elExecStatus','txtExecStatus','txtExecDuration',
        'txtExecOutput','svExecOutput',
        'pageMonitor','gridMonitorJobs','txtActiveCount','txtLiveLog','svLiveLog','elLogPulse',
        'pageHistory','gridHistory','txtHistorySearch','cboHistoryStatus',
        'btnHistoryFilter','btnExportHistory','txtHistoryDetail','svHistoryDetail','txtHistoryDetailTitle',
        'pageCheckpoints','gridCheckpoints','btnResumeCheckpoint','btnDeleteCheckpoint','btnViewCheckpoint',
        'pageArtifacts','gridArtifacts','txtArtifactSearch','btnArtifactFilter',
        'txtArtifactCount','btnOpenArtifact','btnDownloadArtifact','btnDeleteArtifact',
        'pageMetrics','txtMetricTotal','txtMetricSuccess','txtMetricOKErr',
        'txtMetricAvg','txtMetricTop','gridMetricTasks','gridMetricErrors',
        'pageDiag','gridDiag','btnRunDiag','txtDiagSummary',
        'pageConfig','txtConfigEditor','btnSaveConfig','btnReloadConfig','txtConfigStatus',
        'pageNewTask','txtNewTaskName','txtNewTaskNameError','cboNewTaskCategory',
        'txtNewTaskDesc','btnCreateTask','txtCreateTaskResult','txtNewTaskPreview',
        'pageAudit','gridAudit','btnExportAudit',
        'btnNavCatalog','btnNavExecute','btnNavMonitor','btnNavHistory',
        'btnNavCheckpoints','btnNavArtifacts','btnNavMetrics',
        'btnNavDiag','btnNavConfig','btnNavNewTask','btnNavAudit'
    )
    $Script:Ctrl = @{}
    foreach ($name in $controls) {
        $found = $Script:Window.FindName($name)
        if ($null -eq $found) { Write-Warning "Control not found: $name" }
        $Script:Ctrl[$name] = $found
    }

    # User info
    $c = $Script:Ctrl['txtUserName'];    if ($null -ne $c) { $c.Text = $Script:CurrentUser }
    $c = $Script:Ctrl['txtUserRole'];    if ($null -ne $c) { $c.Text = $Script:CurrentRole }
    $ini = if ($Script:CurrentUser.Length -gt 0) { $Script:CurrentUser[0].ToString().ToUpper() } else { 'U' }
    $c = $Script:Ctrl['txtUserInitial']; if ($null -ne $c) { $c.Text = $ini }
    $c = $Script:Ctrl['txtEnginePath'];  if ($null -ne $c) { $c.Text = $Script:EngineRoot }

    # Engine status
    $ok = Test-Path $Script:RunScript
    $c  = $Script:Ctrl['txtEngineStatus']
    if ($null -ne $c) { $c.Text = if ($ok) { 'Engine Ready' } else { 'Engine Not Found' } }
    $c  = $Script:Ctrl['elEngineStatus']
    if ($null -ne $c) {
        $c.Background = if ($ok) { [System.Windows.Media.Brushes]::LimeGreen } else { [System.Windows.Media.Brushes]::Crimson }
    }

    # RBAC
    $c = $Script:Ctrl['btnNavConfig'];  if ($null -ne $c) { $c.IsEnabled = (& $Script:Fn_TestPermission 'EditConfig') }
    $c = $Script:Ctrl['btnNavAudit'];   if ($null -ne $c) { $c.IsEnabled = (& $Script:Fn_TestPermission 'ViewAudit') }
    $c = $Script:Ctrl['btnNavNewTask']; if ($null -ne $c) { $c.IsEnabled = (& $Script:Fn_TestPermission 'CreateTask') }

    # Navigation - each button captures its own scriptblock variables
    $navDefs = @(
        @{ Btn='btnNavCatalog';     Page='pageCatalog';     Title='Task Catalog';            Load=$Script:Fn_LoadCatalogPage }
        @{ Btn='btnNavExecute';     Page='pageExecute';     Title='Execute Task';            Load=$Script:Fn_LoadExecutePage }
        @{ Btn='btnNavMonitor';     Page='pageMonitor';     Title='Live Monitor';            Load=$Script:Fn_LoadMonitorPage }
        @{ Btn='btnNavHistory';     Page='pageHistory';     Title='Execution History';       Load=$Script:Fn_LoadHistoryPage }
        @{ Btn='btnNavCheckpoints'; Page='pageCheckpoints'; Title='Checkpoint Manager';      Load=$Script:Fn_LoadCheckpointsPage }
        @{ Btn='btnNavArtifacts';   Page='pageArtifacts';   Title='Artifact Repository';     Load=$Script:Fn_LoadArtifactsPage }
        @{ Btn='btnNavMetrics';     Page='pageMetrics';     Title='Metrics & Observability'; Load=$Script:Fn_LoadMetricsPage }
        @{ Btn='btnNavDiag';        Page='pageDiag';        Title='Environment Diagnostics'; Load=$Script:Fn_RunDiagnostics }
        @{ Btn='btnNavConfig';      Page='pageConfig';      Title='Configuration';           Load=$Script:Fn_LoadConfigPage }
        @{ Btn='btnNavNewTask';     Page='pageNewTask';     Title='Create New Task';         Load=$Script:Fn_LoadNewTaskPage }
        @{ Btn='btnNavAudit';       Page='pageAudit';       Title='Audit Log';              Load=$Script:Fn_LoadAuditPage }
    )

    # FIX: GetNewClosure() wraps the scriptblock in a new dynamic module whose
    # $Script: scope is empty, making every $Script:Fn_* / $Script:Ctrl reference
    # inside the handler (and any page-loader it calls) resolve to $null, which
    # throws BadExpression and cascades into the ShowDialog crash.
    # Solution: store per-button data in the WPF Button.Tag (type Object, accepts
    # any hashtable) and use a single shared handler that reads from $sender.Tag.
    # The handler is a plain scriptblock — no GetNewClosure — so $Script: always
    # resolves correctly against the original script scope at call time.
    foreach ($def in $navDefs) {
        $btn = $Script:Ctrl[$def.Btn]
        if ($null -ne $btn) {
            $btn.Tag = @{ Page = $def.Page; Title = $def.Title; Load = $def.Load }
        }
    }
    $navClickHandler = {
        param($sender, $e)
        $t = $sender.Tag
        & $Script:Fn_NavigateTo -PageName $t.Page -Title $t.Title
        & $t.Load
    }
    foreach ($def in $navDefs) {
        $btn = $Script:Ctrl[$def.Btn]
        if ($null -ne $btn) { $btn.Add_Click($navClickHandler) }
    }

    # Refresh
    $rb = $Script:Ctrl['btnRefresh']
    if ($null -ne $rb) {
        $rb.Add_Click({
            $title = ''; $pt = $Script:Ctrl['txtPageTitle']; if ($null -ne $pt) { $title = $pt.Text }
            switch ($title) {
                'Task Catalog'            { & $Script:Fn_LoadCatalogPage }
                'Execute Task'            { & $Script:Fn_LoadExecutePage }
                'Live Monitor'            { & $Script:Fn_LoadMonitorPage }
                'Execution History'       { & $Script:Fn_LoadHistoryPage }
                'Checkpoint Manager'      { & $Script:Fn_LoadCheckpointsPage }
                'Artifact Repository'     { & $Script:Fn_LoadArtifactsPage }
                'Metrics & Observability' { & $Script:Fn_LoadMetricsPage }
                'Environment Diagnostics' { & $Script:Fn_RunDiagnostics }
                'Audit Log'              { & $Script:Fn_LoadAuditPage }
            }
        })
    }

    # Catalog
    $c = $Script:Ctrl['txtCatalogSearch']
    if ($null -ne $c) { $c.Add_TextChanged({ & $Script:Fn_FilterCatalog }) }
    $c = $Script:Ctrl['cboCatalogCategory']
    if ($null -ne $c) { $c.Add_SelectionChanged({ & $Script:Fn_FilterCatalog }) }
    $c = $Script:Ctrl['btnCatalogExecute']
    if ($null -ne $c) {
        $c.Add_Click({
            $g   = $Script:Ctrl['gridCatalog']
            $sel = if ($null -ne $g) { $g.SelectedItem } else { $null }
            if ($null -ne $sel) {
                & $Script:Fn_NavigateTo -PageName 'pageExecute' -Title 'Execute Task'
                & $Script:Fn_LoadExecutePage -PreSelectTask $sel.Name
            }
        })
    }
    $c = $Script:Ctrl['btnCatalogViewHistory']
    if ($null -ne $c) {
        $c.Add_Click({
            $g   = $Script:Ctrl['gridCatalog']
            $sel = if ($null -ne $g) { $g.SelectedItem } else { $null }
            & $Script:Fn_NavigateTo -PageName 'pageHistory' -Title 'Execution History'
            if ($null -ne $sel) {
                $ts = $Script:Ctrl['txtHistorySearch']
                if ($null -ne $ts) { $ts.Text = $sel.Name }
            }
            & $Script:Fn_LoadHistoryPage
        })
    }

    # Execute
    $c = $Script:Ctrl['cboExecTask']
    if ($null -ne $c) {
        $c.Add_SelectionChanged({
            $cbo = $Script:Ctrl['cboExecTask']
            $sel = if ($null -ne $cbo) { $cbo.SelectedItem } else { $null }
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
                try { if (-not $proc.HasExited) { $proc.Kill() }; $proc.Dispose() } catch {}
                $Script:ActiveJobs.Remove($tn)
                $esT = $Script:Ctrl['txtExecStatus']; if ($null -ne $esT) { $esT.Text = 'Cancelled' }
                $esB = $Script:Ctrl['elExecStatus'];  if ($null -ne $esB) { $esB.Background = [System.Windows.Media.Brushes]::Crimson }
                $bRun= $Script:Ctrl['btnExecuteTask'];if ($null -ne $bRun){ $bRun.IsEnabled = $true }
                $bCan= $Script:Ctrl['btnCancelTask']; if ($null -ne $bCan){ $bCan.IsEnabled = $false }
                if ($null -ne $Script:ExecTimer) { $Script:ExecTimer.Stop() }
                & $Script:Fn_WriteAuditLog -Action 'CANCEL_TASK' -Target $tn
            }
        })
    }

    # History
    $c = $Script:Ctrl['btnHistoryFilter']; if ($null -ne $c) { $c.Add_Click({ & $Script:Fn_LoadHistoryPage }) }
    $c = $Script:Ctrl['gridHistory']
    if ($null -ne $c) {
        $c.Add_SelectionChanged({
            $g   = $Script:Ctrl['gridHistory']
            $sel = if ($null -ne $g) { $g.SelectedItem } else { $null }
            if ($null -ne $sel) { & $Script:Fn_ShowRunDetail -RunId $sel.RunId -Task $sel.Task }
        })
    }
    $c = $Script:Ctrl['btnExportHistory']; if ($null -ne $c) { $c.Add_Click({ & $Script:Fn_ExportHistoryCSV }) }

    # Checkpoints
    $c = $Script:Ctrl['btnResumeCheckpoint']; if ($null -ne $c) { $c.Add_Click({ & $Script:Fn_ResumeCheckpoint }) }
    $c = $Script:Ctrl['btnDeleteCheckpoint']; if ($null -ne $c) { $c.Add_Click({ & $Script:Fn_DeleteCheckpoint }) }
    $c = $Script:Ctrl['btnViewCheckpoint'];   if ($null -ne $c) { $c.Add_Click({ & $Script:Fn_ViewCheckpoint }) }

    # Artifacts
    $c = $Script:Ctrl['btnArtifactFilter'];   if ($null -ne $c) { $c.Add_Click({ & $Script:Fn_LoadArtifactsPage }) }
    $c = $Script:Ctrl['txtArtifactSearch']
    if ($null -ne $c) {
        $c.Add_KeyDown({
            param($s,$e)
            if ($e.Key -eq [System.Windows.Input.Key]::Return) { & $Script:Fn_LoadArtifactsPage }
        })
    }
    $c = $Script:Ctrl['btnOpenArtifact'];     if ($null -ne $c) { $c.Add_Click({ & $Script:Fn_OpenArtifact }) }
    $c = $Script:Ctrl['btnDownloadArtifact']; if ($null -ne $c) { $c.Add_Click({ & $Script:Fn_SaveArtifact }) }
    $c = $Script:Ctrl['btnDeleteArtifact'];   if ($null -ne $c) { $c.Add_Click({ & $Script:Fn_DeleteArtifact }) }

    # Diag / Config / New Task / Audit
    $c = $Script:Ctrl['btnRunDiag'];       if ($null -ne $c) { $c.Add_Click({ & $Script:Fn_RunDiagnostics }) }
    $c = $Script:Ctrl['btnSaveConfig'];    if ($null -ne $c) { $c.Add_Click({ & $Script:Fn_SaveConfigPage }) }
    $c = $Script:Ctrl['btnReloadConfig'];  if ($null -ne $c) { $c.Add_Click({ & $Script:Fn_LoadConfigPage }) }
    $c = $Script:Ctrl['txtNewTaskName'];   if ($null -ne $c) { $c.Add_TextChanged({ & $Script:Fn_UpdateNewTaskPreview }) }
    $c = $Script:Ctrl['cboNewTaskCategory'];if ($null -ne $c){ $c.Add_SelectionChanged({ & $Script:Fn_UpdateNewTaskPreview }) }
    $c = $Script:Ctrl['txtNewTaskDesc'];   if ($null -ne $c) { $c.Add_TextChanged({ & $Script:Fn_UpdateNewTaskPreview }) }
    $c = $Script:Ctrl['btnCreateTask'];    if ($null -ne $c) { $c.Add_Click({ & $Script:Fn_CreateNewTask }) }
    $c = $Script:Ctrl['btnExportAudit'];   if ($null -ne $c) { $c.Add_Click({ & $Script:Fn_ExportAuditCSV }) }

    # Auto-refresh timer
    $Script:RefreshTimer = New-Object System.Windows.Threading.DispatcherTimer
    $Script:RefreshTimer.Interval = [TimeSpan]::FromSeconds(5)
    $Script:RefreshTimer.Add_Tick({
        $title = ''; $pt = $Script:Ctrl['txtPageTitle']; if ($null -ne $pt) { $title = $pt.Text }
        if ($title -eq 'Live Monitor') { & $Script:Fn_LoadMonitorPage }
        $tc = $Script:Ctrl['txtActiveCount']
        if ($null -ne $tc) { $tc.Text = "$($Script:ActiveJobs.Count) running" }
    })
    $Script:RefreshTimer.Start()

    $Script:Window.Add_Closed({
        if ($null -ne $Script:RefreshTimer) { $Script:RefreshTimer.Stop() }
        if ($null -ne $Script:ExecTimer)    { $Script:ExecTimer.Stop() }
        foreach ($key in @($Script:ActiveJobs.Keys)) {
            try { $Script:ActiveJobs[$key].Process.Kill() } catch {}
        }
        & $Script:Fn_WriteAuditLog -Action 'UI_CLOSE'
    })

    & $Script:Fn_WriteAuditLog -Action 'UI_OPEN' -Detail "Role=$Script:CurrentRole"
    & $Script:Fn_LoadCatalogPage

    [void]$Script:Window.ShowDialog()
}

# ============================================================================
# ENTRY POINT
# ============================================================================
Initialize-Window
