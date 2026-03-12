#Requires -Version 5.1
# =============================================================================
# tasks/task_excel_reporte.ps1
# @Description : Genera un reporte Excel a partir de un CSV en input/
# @Category    : Excel
# @Version     : 1.0.0
# @Param       : Extra string optional "Nombre del archivo CSV en input/ (sin ruta)"
# =============================================================================
# Solo ASCII. PS 5.1.

# Synopsis: Convierte CSV de input/ a reporte Excel en reports/
task excel_reporte {
    $ctx = New-TaskContext `
        -TaskName 'excel_reporte' `
        -Config   $Script:Config `
        -Root     (Split-Path $BuildRoot -Parent) `
        -Params   @{ Archivo = $Extra }

    Write-BuildLog $ctx 'INFO' "Iniciando excel_reporte | Archivo=$Extra"

    # ---- Prerequisitos ------------------------------------------------
    $csvNombre = if ([string]::IsNullOrWhiteSpace($Extra)) { 'datos.csv' } else { $Extra }
    $csvPath   = Join-Path $ctx.Paths.Input $csvNombre
    Test-TaskAsset -Context $ctx -Files @{ CSV = $csvPath }

    # ---- Leer CSV -----------------------------------------------------
    Write-BuildLog $ctx 'INFO' "Leyendo CSV: $csvPath"
    $datos = Import-Csv -Path $csvPath -Encoding Default
    if ($datos.Count -eq 0) {
        throw "El CSV esta vacio: $csvPath"
    }
    Write-BuildLog $ctx 'INFO' "CSV leido: $($datos.Count) filas"

    # Convertir a array de hashtables
    $filas = $datos | ForEach-Object {
        $h = @{}
        $_.PSObject.Properties | ForEach-Object { $h[$_.Name] = $_.Value }
        $h
    }

    # ---- Excel --------------------------------------------------------
    $xl = $null
    $wb = $null

    try {
        $xl = New-ExcelApp -Context $ctx
        if ($null -eq $xl) {
            throw 'Excel no disponible. Abrir y cerrar Excel manualmente y reintentar.'
        }

        $wb = New-ExcelWorkbook -Context $ctx -ExcelApp $xl
        $ws = Get-ExcelSheet -Workbook $wb -Index 1
        $ws.Name = 'Reporte'

        Write-ExcelData -Context $ctx -Sheet $ws -Data $filas

        # Autofit columnas
        Invoke-ExcelAutoFit -Sheet $ws

        $ts      = Get-Date -Format 'yyyyMMdd_HHmmss'
        $outFile = Join-Path $ctx.Paths.Reports "reporte_${ts}.xlsx"
        Save-ExcelWorkbook -Context $ctx -Workbook $wb -Path $outFile

        Write-Build Green "  Reporte generado: $outFile"
        Write-BuildLog $ctx 'INFO' "Reporte generado: $outFile"

        Write-RunResult -Context $ctx -Success $true

    } finally {
        Close-ExcelWorkbook -Workbook $wb -Save $false
        Close-ExcelApp -ExcelApp $xl
        $xl = $null
        $wb = $null
    }
}
