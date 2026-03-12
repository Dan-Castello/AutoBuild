#Requires -Version 5.1
# =============================================================================
# tasks/task_sap_stock.ps1
# @Description : Consulta stock por almacen via SAP GUI (transaccion MMBE)
# @Category    : SAP
# @Version     : 1.0.0
# @Param       : Centro  string required "Centro SAP (ej: 1000)"
# @Param       : Almacen string optional "Almacen (default: todos)"
# =============================================================================
# Solo ASCII. PS 5.1.

# Synopsis: Extrae stock de almacen desde SAP (MMBE) y genera reporte Excel
task sap_stock {
    $ctx = New-TaskContext `
        -TaskName 'sap_stock' `
        -Config   $Script:Config `
        -Root     (Split-Path $BuildRoot -Parent) `
        -Params   @{ Centro = $Centro; Almacen = $Almacen }

    Write-BuildLog $ctx 'INFO' "Iniciando sap_stock | Centro=$Centro Almacen=$Almacen"

    # ---- Prerequisitos ------------------------------------------------
    Test-TaskAsset -Context $ctx -Params @{ Centro = $Centro }

    # ---- Obtener sesion SAP --------------------------------------------
    $sap = Get-SapSession -Context $ctx
    Assert-SapSession -Session $sap -Context $ctx

    # ---- Abrir Excel para el reporte -----------------------------------
    $xl = $null
    $wb = $null

    try {
        $xl = New-ExcelApp -Context $ctx
        if ($null -eq $xl) {
            throw 'Excel no disponible. Abrir y cerrar Excel manualmente y reintentar.'
        }
        $wb = New-ExcelWorkbook -Context $ctx -ExcelApp $xl
        $ws = Get-ExcelSheet -Workbook $wb -Index 1
        $ws.Name = 'Stock'

        # ---- Navegar en SAP -------------------------------------------
        Invoke-SapTransaction -Context $ctx -Session $sap -TCode 'MMBE'

        # Introducir centro
        Set-SapField -Context $ctx -Session $sap `
            -FieldId 'wnd[0]/usr/ctxtWERKS-LOW' -Value $Centro

        # Introducir almacen si se especifico
        if (-not [string]::IsNullOrWhiteSpace($Almacen)) {
            Set-SapField -Context $ctx -Session $sap `
                -FieldId 'wnd[0]/usr/ctxtLGORT-LOW' -Value $Almacen
        }

        # Ejecutar
        Invoke-SapButton -Context $ctx -Session $sap `
            -ButtonId 'wnd[0]/tbar[1]/btn[8]'

        Write-BuildLog $ctx 'INFO' 'Consulta SAP ejecutada. Exportando resultado...'

        # ---- Guardar reporte ------------------------------------------
        $ts      = Get-Date -Format 'yyyyMMdd_HHmmss'
        $outFile = Join-Path $ctx.Paths.Reports "stock_${Centro}_${ts}.xlsx"

        Save-ExcelWorkbook -Context $ctx -Workbook $wb -Path $outFile

        Write-Build Green "  Reporte guardado: $outFile"
        Write-BuildLog $ctx 'INFO' "Reporte guardado: $outFile"

        Write-RunResult -Context $ctx -Success $true

    } finally {
        Close-ExcelWorkbook -Workbook $wb -Save $false
        Close-ExcelApp -ExcelApp $xl
        $xl = $null
        $wb = $null
    }
}
