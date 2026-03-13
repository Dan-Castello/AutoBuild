#Requires -Version 5.1
# =============================================================================
# tasks/task_diag_engine.ps1
# @Description : Engine self-diagnostics — config validation, library health, path checks
# @Category    : Utility
# @Version     : 1.1.0
# @Author      : AutoBuild QA
# =============================================================================
# Synopsis: Diagnostics - engine config, library load, path checks, Test-EngineConfiguration
# Params: {}
#
# REMEDIATION v3.1:
#   Uses Test-EngineConfiguration to validate config. OPTIONAL issues (SMTP,
#   SAP, security) are reported informatively, not as failures.

task diag_engine {

    $ctx = New-TaskContext `
        -TaskName 'diag_engine' `
        -Config   $Script:EngineConfig `
        -Root     $Script:EngineRoot `
        -Params   $Script:TaskParams

    Write-BuildLog $ctx 'INFO' 'DIAG ENGINE: Starting engine self-diagnostics'
    Write-Build Cyan "`n  AutoBuild Engine Diagnostics"

    $results  = [System.Collections.Generic.List[hashtable]]::new()
    $failures = 0
    $script:failures = 0
    $stamp    = Get-Date -Format 'yyyyMMdd_HHmmss'

    function Add-Result {
        param([string]$Check, [bool]$Pass, [string]$Detail = '', [bool]$Optional = $false)
        $results.Add(@{ Check = $Check; Pass = $Pass; Detail = $Detail; Optional = $Optional })
        $sym   = if ($Pass) { '[OK]' } else { if ($Optional) { '[INFO]' } else { '[FAIL]' } }
        $lvl   = if ($Pass) { 'INFO' } else { if ($Optional) { 'INFO' } else { 'ERROR' } }
        $color = if ($Pass) { 'Green' } else { if ($Optional) { 'Yellow' } else { 'Red' } }
        Write-BuildLog $ctx $lvl "$sym $Check" -Detail $Detail
        Write-Build $color "  $sym  $Check$(if($Detail){' -- '+$Detail})"
        if (-not $Pass -and -not $Optional) { $script:failures++ }
    }

    # ─── 1. CONFIG VALIDATION ─────────────────────────────────────────────────
    Write-Build Cyan "`n  [1/4] Configuration Validation (Test-EngineConfiguration)"

    $cfgResult = Test-EngineConfiguration -Config $ctx.Config -Root $Script:EngineRoot

    Add-Result 'Configuration passes CRITICAL checks' $cfgResult.Valid $cfgResult.Summary

    foreach ($issue in $cfgResult.CriticalIssues) {
        Add-Result "CRITICAL: $issue" $false $issue
    }
    foreach ($issue in $cfgResult.OptionalIssues) {
        Add-Result "Optional: $issue" $true $issue -Optional $true
    }

    # SMTP-specific informational display
    if (-not $cfgResult.SmtpConfigured) {
        Write-Build Yellow '  [INFO] SMTP not configured — notifications are disabled (this is normal in dev/staging)'
    }

    # ─── 2. LIBRARY AVAILABILITY ─────────────────────────────────────────────
    Write-Build Cyan "`n  [2/4] Library Function Availability"

    $requiredFunctions = @(
        'Write-BuildLog', 'Write-RunResult', 'New-RunId',
        'New-TaskContext',
        'Test-ComAvailable', 'Invoke-ReleaseComObject', 'Remove-ZombieCom',
        'New-ExcelApp', 'New-ExcelWorkbook', 'Get-ExcelSheet',
        'Write-ExcelRange', 'Save-ExcelWorkbook', 'Close-ExcelApp',
        'Invoke-WithRetry',
        'Send-Notification', 'Test-NotificationConfig',
        'Test-EngineConfiguration'
    )

    foreach ($fn in $requiredFunctions) {
        $exists = $null -ne (Get-Command $fn -ErrorAction SilentlyContinue)
        Add-Result "Function: $fn" $exists $(if(-not $exists){'Not found -- check lib load order'})
    }

    # ─── 3. PATH CHECKS ──────────────────────────────────────────────────────
    Write-Build Cyan "`n  [3/4] Path Checks"

    $pathChecks = @{
        'Root'    = $ctx.Paths.Root
        'Input'   = $ctx.Paths.Input
        'Output'  = $ctx.Paths.Output
        'Reports' = $ctx.Paths.Reports
        'Logs'    = $ctx.Paths.Logs
    }

    foreach ($name in $pathChecks.Keys) {
        $p = $pathChecks[$name]
        $exists = Test-Path $p
        if (-not $exists) {
            try { New-Item -ItemType Directory -Path $p -Force | Out-Null; $exists = $true } catch {}
        }
        Add-Result "Path exists: $name" $exists $p
    }

    # ─── 4. ENVIRONMENT SNAPSHOT ─────────────────────────────────────────────
    Write-Build Cyan "`n  [4/4] Environment"

    $psOk  = $PSVersionTable.PSVersion.Major -eq 5
    $edOk  = $PSVersionTable.PSEdition -eq 'Desktop'
    $ibVer = $ctx.Config.engine.ibVersion

    Add-Result "PowerShell 5.x (required)" $psOk "Got: $($PSVersionTable.PSVersion)"
    Add-Result "PowerShell Desktop edition" $edOk "Got: $($PSVersionTable.PSEdition)"
    Add-Result "Invoke-Build version detected" ($ibVer -ne 'unknown') "IB: $ibVer"

    Write-Build White "  PS $($PSVersionTable.PSVersion) $($PSVersionTable.PSEdition)"
    Write-Build White "  CLR: $([System.Environment]::Version)"
    Write-Build White "  OS: $([System.Environment]::OSVersion.VersionString)"
    Write-Build White "  Heap: $([math]::Round([GC]::GetTotalMemory($false)/1MB,2))MB"

    # ─── SUMMARY ──────────────────────────────────────────────────────────────
    Write-Build Cyan "`n  ── DIAG ENGINE SUMMARY ──────────────────────────────────"
    $total  = $results.Count
    $passed = @($results | Where-Object { $_.Pass }).Count
    Write-Build White "  Total: $total  |  Passed: $passed  |  Failed: $failures"

    $outFile = Join-Path $ctx.Paths.Reports "diag_engine_results_${stamp}.json"
    try {
        @{ runId=$ctx.RunId; total=$total; passed=$passed; failed=$failures; checks=$results.ToArray() } |
            ConvertTo-Json -Depth 4 |
            ForEach-Object { [System.IO.File]::WriteAllText($outFile, $_, [System.Text.Encoding]::ASCII) }
        Write-Build Cyan "  Report: $outFile"
    } catch { Write-BuildLog $ctx 'WARN' "Report write failed: $_" }

    if ($failures -gt 0) {
        Write-RunResult -Context $ctx -Success $false -ErrorMsg "$failures check(s) failed"
        throw "DIAG ENGINE: $failures check(s) failed"
    }
    Write-RunResult -Context $ctx -Success $true
}
