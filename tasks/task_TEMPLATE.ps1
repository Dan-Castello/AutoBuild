#Requires -Version 5.1
# =============================================================================
# tasks/task_NOMBRE.ps1
# @Description : Brief task description
# @Category    : SAP | Excel | CSV | Report | Utility
# @Version     : 1.0.0
# @Author      : Your Name
# =============================================================================
#
# MANDATORY RULES:
#   - ASCII only in this file
#   - Filename: task_[a-zA-Z0-9_-]+.ps1
#   - One task per file: task NOMBRE { ... }
#   - No global function definitions (use lib/ for shared logic)
#   - No hardcoded paths (use $ctx.Paths.*)
#   - No New-Object -ComObject directly (use New-ExcelApp, Invoke-ComWithTimeout)
#   - COM: always use try/finally with Invoke-ComCleanup or explicit Release calls
#   - Always call Write-RunResult at task completion (success or failure)
#
# TASK-01 FIX:
#   Do NOT call (Split-Path $BuildRoot -Parent) for the project root.
#   Use $Script:EngineRoot — injected by Main.build.ps1 before this file loads.
#
# RUN-01 FIX:
#   Do NOT declare task-specific parameters in Main.build.ps1 or Run.ps1.
#   Parameters arrive via $Script:TaskParams (from -Params JSON in Run.ps1).
#   Access: $ctx.Params['MyParam']
#   Validate: Test-TaskAsset -Context $ctx -Params @{ MyParam = $ctx.Params['MyParam'] }
#   Example call: .\Run.ps1 NOMBRE -Params '{"Centro":"1000","Almacen":"WH01"}'
# =============================================================================

# Synopsis: Brief description shown in .\Run.ps1 -List
task NOMBRE {

    # ---- Context creation --------------------------------------------------
    # $Script:EngineRoot  : project root (TASK-01 fix, injected by engine)
    # $Script:EngineConfig: master config, read-only reference
    # $Script:TaskParams  : all -Params key/value pairs from Run.ps1
    $ctx = New-TaskContext `
        -TaskName 'NOMBRE' `
        -Config   $Script:EngineConfig `
        -Root     $Script:EngineRoot `
        -Params   $Script:TaskParams

    Write-BuildLog $ctx 'INFO' 'Starting NOMBRE'

    # ---- Prerequisite validation -------------------------------------------
    Test-TaskAsset -Context $ctx -Params @{
        'Centro' = $ctx.Params['Centro']   # Example required parameter
    }

    # ---- Option A: Task without COM ----------------------------------------
    try {
        # ... logic here ...

        Write-RunResult -Context $ctx -Success $true
    } catch {
        Write-BuildLog $ctx 'ERROR' "Task NOMBRE failed: $_" -Detail $_.ScriptStackTrace
        Write-RunResult -Context $ctx -Success $false -ErrorMsg "$_"
        throw
    }

    # ---- Option B: Excel task (uncomment to use) ---------------------------
    # $xl = $null
    # $wb = $null
    # $ws = $null
    # try {
    #     $xl = New-ExcelApp -Context $ctx
    #     if ($null -eq $xl) { throw 'Excel not available' }
    #
    #     $wb = New-ExcelWorkbook -Context $ctx -ExcelApp $xl
    #     $ws = Get-ExcelSheet -Workbook $wb -Index 1
    #
    #     # Build data (use Generic List for large datasets)
    #     $data = [System.Collections.Generic.List[hashtable]]::new()
    #     $data.Add(@{ Column1 = 'value1'; Column2 = 'value2' })
    #
    #     # PERFORMANCE: Write-ExcelRange writes ALL rows in ONE COM call
    #     Write-ExcelRange -Context $ctx -Sheet $ws -Data $data.ToArray()
    #     Invoke-ExcelAutoFit -Sheet $ws
    #
    #     $outFile = Join-Path $ctx.Paths.Reports "NOMBRE_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
    #     Save-ExcelWorkbook -Context $ctx -Workbook $wb -Path $outFile -Format 'xlsx'
    #
    #     Write-RunResult -Context $ctx -Success $true
    # } catch {
    #     Write-BuildLog $ctx 'ERROR' "Task NOMBRE failed: $_" -Detail $_.ScriptStackTrace
    #     Write-RunResult -Context $ctx -Success $false -ErrorMsg "$_"
    #     throw
    # } finally {
    #     # Release in reverse order. finally always runs.
    #     if ($null -ne $ws) { Invoke-ReleaseComObject $ws; $ws = $null }
    #     if ($null -ne $wb) { Close-ExcelWorkbook -Workbook $wb -Save $false }
    #     Close-ExcelApp -ExcelApp $xl
    #     $xl = $null
    # }

    # ---- Option C: SAP task (uncomment to use) -----------------------------
    # $sess = $null
    # try {
    #     $sess = Get-SapSession -Context $ctx
    #     Assert-SapSession -Session $sess -Context $ctx
    #
    #     Invoke-SapTransaction -Context $ctx -Session $sess -TCode 'MM03'
    #     Set-SapField -Context $ctx -Session $sess `
    #         -FieldId 'wnd[0]/usr/ctxtRMMG1-MATNR' -Value $ctx.Params['Material']
    #     Invoke-SapButton -Context $ctx -Session $sess -ButtonId 'wnd[0]/tbar[0]/btn[0]'
    #     # Wait-SapReady is called automatically by Invoke-SapButton
    #
    #     # Use cancellation token for large table exports
    #     $cancel = $false
    #     $tableData = Export-SapTableToArray -Context $ctx -Session $sess `
    #         -TableId 'wnd[0]/usr/tblSAPLMGMMTC_6100' `
    #         -Columns @('MATNR','MAKTX','MEINS') `
    #         -CancellationToken ([ref]$cancel)
    #
    #     Write-RunResult -Context $ctx -Success $true
    # } catch {
    #     Write-BuildLog $ctx 'ERROR' "SAP task NOMBRE failed: $_" -Detail $_.ScriptStackTrace
    #     Write-RunResult -Context $ctx -Success $false -ErrorMsg "$_"
    #     throw
    # } finally {
    #     Release-SapSession -Session $sess
    # }
}
