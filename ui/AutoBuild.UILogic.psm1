#Requires -Version 5.1
# =============================================================================
# ui/AutoBuild.UILogic.psm1
# AutoBuild v3.1 - UI business logic module (testable, decoupled from WPF).
#
# FIX DEBT (AUDIT v3 MEDIUM): AutoBuild.UI.ps1 had 1,500+ lines of mixed
# WPF event handlers and business logic in a single file. The business logic
# functions (config save, task creation, log purge, history export) could not
# be unit-tested without loading the entire WPF stack.
#
# This module extracts the pure business logic into testable, dependency-injected
# functions. AutoBuild.UI.ps1 delegates to these functions via scriptblock
# adapters that supply the WPF context (file paths, config objects).
#
# TESTING: These functions can be called from Pester tests without a display
# server by injecting test paths and mock config objects.
#
# IMPORT in AutoBuild.UI.ps1:
#   Import-Module (Join-Path $Script:UIRoot 'AutoBuild.UILogic.psm1') -Force
# =============================================================================
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# CONFIG VALIDATION (extracted from $Script:Fn_SaveConfigPage)
# ---------------------------------------------------------------------------

function Test-EngineConfigObject {
    <#
    .SYNOPSIS
        Validates an engine.config.json object. Returns a list of validation errors.
        Empty list means config is valid.
    .PARAMETER ConfigObject
        Deserialized PSCustomObject from ConvertFrom-Json.
    .OUTPUTS
        [string[]] — list of error messages. Empty array = valid.
    #>
    param([Parameter(Mandatory)]$ConfigObject)

    $errors = [System.Collections.Generic.List[string]]::new()

    # Helper: safely get a property from a PSCustomObject without throwing under StrictMode
    function Get-SafeProp {
        param($Obj, [string]$Prop)
        if ($null -eq $Obj) { return $null }
        if ($Obj -is [hashtable]) {
            if ($Obj.ContainsKey($Prop)) { return $Obj[$Prop] }
            return $null
        }
        $p = $Obj.PSObject.Properties[$Prop]
        if ($null -eq $p) { return $null }
        return $p.Value
    }

    $cfgEngine        = Get-SafeProp $ConfigObject 'engine'
    $cfgNotifications = Get-SafeProp $ConfigObject 'notifications'
    $cfgSecurity      = Get-SafeProp $ConfigObject 'security'

    # maxRetries
    try {
        $mr = [int](Get-SafeProp $cfgEngine 'maxRetries')
        if ($mr -lt 0 -or $mr -gt 10) { $errors.Add("engine.maxRetries must be 0-10 (got $mr)") }
    } catch { $errors.Add("engine.maxRetries is not a valid integer") }

    # retryDelaySeconds
    try {
        $rd_val = Get-SafeProp $cfgEngine 'retryDelaySeconds'
        if ($null -ne $rd_val) {
            $rd = [double]$rd_val
            if ($rd -lt 0 -or $rd -gt 300) { $errors.Add("engine.retryDelaySeconds must be 0-300 (got $rd)") }
        }
    } catch { $errors.Add("engine.retryDelaySeconds is not a valid number") }

    # logLevel
    $validLevels = @('DEBUG','INFO','WARN','ERROR','FATAL')
    $logLevelVal = Get-SafeProp $cfgEngine 'logLevel'
    if (-not [string]::IsNullOrWhiteSpace($logLevelVal)) {
        if ($validLevels -notcontains $logLevelVal.Trim().ToUpper()) {
            $errors.Add("engine.logLevel must be one of: $($validLevels -join ', ')")
        }
    }

    # smtpServer
    $smtpVal = Get-SafeProp $cfgNotifications 'smtpServer'
    if (-not [string]::IsNullOrWhiteSpace($smtpVal)) {
        $smtp = $smtpVal.Trim()
        if ($smtp.Length -gt 253) { $errors.Add("notifications.smtpServer too long (max 253 chars)") }
        if ($smtp -match '[;&|`$<>!{}()\[\]\\]') { $errors.Add("notifications.smtpServer contains invalid characters") }
    }

    # smtpPort
    try {
        $portVal = Get-SafeProp $cfgNotifications 'smtpPort'
        if ($null -ne $portVal) {
            $port = [int]$portVal
            if ($port -lt 1 -or $port -gt 65535) { $errors.Add("notifications.smtpPort must be 1-65535 (got $port)") }
        }
    } catch { $errors.Add("notifications.smtpPort is not a valid integer") }

    # AD groups (must start with CN= if non-empty)
    foreach ($field in @('adminAdGroup','developerAdGroup')) {
        $val = Get-SafeProp $cfgSecurity $field
        if (-not [string]::IsNullOrWhiteSpace($val)) {
            $val = $val.Trim()
            if (-not ($val -match '^CN=')) { $errors.Add("security.$field must be a valid DN starting with 'CN='") }
            if ($val.Length -gt 512) { $errors.Add("security.$field is too long (max 512 chars)") }
        }
    }

    # User whitelists
    foreach ($field in @('adminUsers','developerUsers')) {
        $val = Get-SafeProp $cfgSecurity $field
        if (-not [string]::IsNullOrWhiteSpace($val)) {
            $users = $val -split '[,;\s]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
            foreach ($u in $users) {
                if ($u -match '[^a-zA-Z0-9._\-]') {
                    $errors.Add("security.${field}: invalid username '$u' (only letters, digits, ., -, _ allowed)")
                }
            }
        }
    }

    return [string[]]$errors.ToArray()
}

function Save-EngineConfig {
    <#
    .SYNOPSIS
        Validates and writes engine.config.json atomically.
    .PARAMETER ConfigJson
        Raw JSON string to validate and save.
    .PARAMETER ConfigFile
        Full path to engine.config.json.
    .OUTPUTS
        @{ Success=$bool; Error=$string }
    #>
    param(
        [Parameter(Mandatory)][string]$ConfigJson,
        [Parameter(Mandatory)][string]$ConfigFile
    )

    try {
        $obj    = $ConfigJson | ConvertFrom-Json
        $errs   = Test-EngineConfigObject -ConfigObject $obj
        if ($errs.Count -gt 0) {
            return @{ Success = $false; Error = ($errs -join '; ') }
        }
        # Atomic write: temp + rename
        $tmp = "$ConfigFile.tmp"
        [System.IO.File]::WriteAllText($tmp, $ConfigJson, [System.Text.Encoding]::ASCII)
        Move-Item -Path $tmp -Destination $ConfigFile -Force -ErrorAction Stop
        return @{ Success = $true; Error = '' }
    } catch {
        return @{ Success = $false; Error = "$_" }
    }
}

# ---------------------------------------------------------------------------
# TASK CREATION (extracted from $Script:Fn_CreateNewTask)
# ---------------------------------------------------------------------------

function New-TaskFromTemplate {
    <#
    .SYNOPSIS
        Creates a new task file from the template with literal string substitution.
        FIX V-06 already applied here: uses [string]::Replace() not regex -replace.
    .PARAMETER TemplateFile   Full path to task_TEMPLATE.ps1.
    .PARAMETER OutputFile     Full path for the new task_<Name>.ps1.
    .PARAMETER Name           Task identifier (alphanumeric, _, -).
    .PARAMETER Category       SAP | Excel | CSV | Report | Utility.
    .PARAMETER Description    Brief description text.
    .PARAMETER Author         Author name.
    .OUTPUTS
        @{ Success=$bool; OutputFile=$string; Error=$string }
    #>
    param(
        [Parameter(Mandatory)][string]$TemplateFile,
        [Parameter(Mandatory)][string]$OutputFile,
        [Parameter(Mandatory)][string]$Name,
        [string]$Category    = 'Utility',
        [string]$Description = 'Description pending',
        [string]$Author      = ''
    )

    if (-not (Test-Path $TemplateFile)) {
        return @{ Success = $false; OutputFile = ''; Error = "Template not found: $TemplateFile" }
    }
    if (Test-Path $OutputFile) {
        return @{ Success = $false; OutputFile = ''; Error = "Task already exists: $OutputFile" }
    }
    if ([string]::IsNullOrWhiteSpace($Author)) {
        $Author = try { ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name -split '\\')[-1] } catch { $env:USERNAME }
    }

    try {
        $content = Get-Content $TemplateFile -Raw -Encoding ASCII
        # FIX V-06: literal replacement only
        $content = $content.Replace('@Description : Brief task description',              "@Description : $Description")
        $content = $content.Replace('@Category    : SAP | Excel | CSV | Report | Utility', "@Category    : $Category")
        $content = $content.Replace('@Author      : Your Name',                            "@Author      : $Author")
        $content = $content.Replace('# Synopsis: Brief description shown in .\Run.ps1 -List', "# Synopsis: $Description")
        $content = $content.Replace('task NOMBRE {',                                       "task $Name {")
        $content = $content.Replace("-TaskName 'NOMBRE'",                                  "-TaskName '$Name'")
        [System.IO.File]::WriteAllText($OutputFile, $content, [System.Text.Encoding]::ASCII)
        return @{ Success = $true; OutputFile = $OutputFile; Error = '' }
    } catch {
        return @{ Success = $false; OutputFile = ''; Error = "$_" }
    }
}

# ---------------------------------------------------------------------------
# LOG PURGE (extracted from $Script:Fn_PurgeOldLogs)
# ---------------------------------------------------------------------------

function Invoke-UILogPurge {
    <#
    .SYNOPSIS
        Purges rotated log archives older than RetentionDays.
        Delegates to Logger.ps1's Invoke-LogPurge when available.
        Falls back to in-place atomic rewrite.
    .PARAMETER LogsDir         Path to logs/ directory.
    .PARAMETER RegistryFile    Path to registry.jsonl.
    .PARAMETER RetentionDays   Archives older than this are deleted.
    .PARAMETER LogPurgeFn      Optional: scriptblock reference to Invoke-LogPurge from Logger.ps1.
    .OUTPUTS
        @{ Success=$bool; PurgedCount=$int; Error=$string }
    #>
    param(
        [Parameter(Mandatory)][string]$LogsDir,
        [Parameter(Mandatory)][string]$RegistryFile,
        [int]$RetentionDays = 30,
        [scriptblock]$LogPurgeFn = $null
    )

    try {
        if ($null -ne $LogPurgeFn) {
            & $LogPurgeFn -LogsDir $LogsDir -RetentionDays $RetentionDays
            return @{ Success = $true; PurgedCount = -1; Error = '' }
        }

        # Fallback: in-place filter with atomic write
        $purged = 0
        $cutoff = [datetime]::Now.AddDays(-$RetentionDays)
        if (Test-Path $RegistryFile) {
            $lines = @(Get-Content $RegistryFile -Encoding ASCII -ErrorAction SilentlyContinue)
            $kept  = @($lines | Where-Object {
                try { [datetime]($_ | ConvertFrom-Json).ts -gt $cutoff } catch { $true }
            })
            $purged = $lines.Count - $kept.Count
            if ($purged -gt 0) {
                $tmp = "$RegistryFile.tmp"
                [System.IO.File]::WriteAllLines($tmp, $kept, [System.Text.Encoding]::ASCII)
                Move-Item -Path $tmp -Destination $RegistryFile -Force
            }
        }
        return @{ Success = $true; PurgedCount = $purged; Error = '' }
    } catch {
        return @{ Success = $false; PurgedCount = 0; Error = "$_" }
    }
}

# ---------------------------------------------------------------------------
# HISTORY EXPORT (extracted from $Script:Fn_ExportHistoryCSV)
# ---------------------------------------------------------------------------

function Export-RunHistoryCsv {
    <#
    .SYNOPSIS
        Exports run summary data to a CSV file.
    .PARAMETER Summaries   Array of run summary objects.
    .PARAMETER OutputFile  Full path for the output CSV.
    .OUTPUTS
        @{ Success=$bool; Error=$string }
    #>
    param(
        [Parameter(Mandatory)][object[]]$Summaries,
        [Parameter(Mandatory)][string]$OutputFile
    )

    try {
        $rows = @($Summaries | Select-Object RunId, Task, Started, Elapsed, Status, Entries, User)
        $rows | Export-Csv $OutputFile -NoTypeInformation -Encoding ASCII
        return @{ Success = $true; Error = '' }
    } catch {
        return @{ Success = $false; Error = "$_" }
    }
}

Export-ModuleMember -Function @(
    'Test-EngineConfigObject',
    'Save-EngineConfig',
    'New-TaskFromTemplate',
    'Invoke-UILogPurge',
    'Export-RunHistoryCsv'
)
