#Requires -Version 5.1
# =============================================================================
# lib/ExcelHelper.ps1
# Helpers Excel COM para PS 5.1.
# REGLA FUNDAMENTAL: cada .Property o .Method que retorna un objeto COM debe
# capturarse en variable y liberarse con Release-ComObject antes de continuar.
# Las cadenas inline tipo $xl.Workbooks.Add() o $ws.Cells(1,1).Value2 crean
# COM objects intermedios que el GC de PS 5.1 no libera a tiempo, y el
# ref-count del proceso EXCEL nunca llega a cero aunque se llame Quit().
# =============================================================================
# Solo ASCII. PS 5.1.

Set-StrictMode -Version Latest

function New-ExcelApp {
    <#
    .SYNOPSIS
        Crea una instancia silenciosa de Excel con timeout de seguridad.
        Retorna $null si Excel no esta disponible.
    #>
    param(
        [hashtable]$Context,
        [int]$TimeoutSec = 30
    )

    $xl = Invoke-ComWithTimeout -Context $Context -ProgId 'Excel.Application' `
        -TimeoutSec $TimeoutSec -Label 'Excel'

    if ($null -eq $xl) { return $null }

    try {
        $xl.Visible         = $Context.Config.excel.visible
        $xl.DisplayAlerts   = $false
        $xl.ScreenUpdating  = $Context.Config.excel.screenUpdating
        $xl.EnableEvents    = $false
        $xl.Interactive     = $false
    } catch {
        Write-BuildLog $Context 'WARN' "Advertencia configurando Excel: $_"
    }

    return $xl
}

function New-ExcelWorkbook {
    <#
    .SYNOPSIS
        Crea un nuevo libro vacio.
        NOTA: captura la coleccion Workbooks y la libera inmediatamente.
    #>
    param(
        [hashtable]$Context,
        $ExcelApp
    )

    if ($null -eq $ExcelApp) { throw 'ExcelApp es nulo' }
    $wbs = $null
    try {
        $wbs = $ExcelApp.Workbooks
        $wb  = $wbs.Add()
        Write-BuildLog $Context 'DEBUG' 'Workbook creado'
        return $wb
    } catch {
        Write-BuildLog $Context 'ERROR' "Error al crear workbook: $_"
        throw
    } finally {
        if ($null -ne $wbs) { Release-ComObject $wbs; $wbs = $null }
    }
}

function Open-ExcelWorkbook {
    <#
    .SYNOPSIS
        Abre un libro existente.
        Captura y libera la coleccion Workbooks.
    #>
    param(
        [hashtable]$Context,
        $ExcelApp,
        [string]$Path,
        [bool]$ReadOnly = $true
    )

    if ($null -eq $ExcelApp) { throw 'ExcelApp es nulo' }
    if (-not (Test-Path $Path)) { throw "Archivo no encontrado: $Path" }

    $wbs = $null
    try {
        $wbs = $ExcelApp.Workbooks
        $wb  = $wbs.Open($Path, 0, $ReadOnly)
        Write-BuildLog $Context 'DEBUG' "Workbook abierto: $Path"
        return $wb
    } catch {
        Write-BuildLog $Context 'ERROR' "Error al abrir workbook $Path : $_"
        throw
    } finally {
        if ($null -ne $wbs) { Release-ComObject $wbs; $wbs = $null }
    }
}

function Get-ExcelSheet {
    <#
    .SYNOPSIS
        Obtiene una hoja por indice liberando la coleccion Worksheets.
        Siempre usar esta funcion en lugar de $wb.Worksheets.Item(n) directamente.
        Devuelve el objeto Worksheet; el llamante es responsable de liberarlo.
    #>
    param(
        $Workbook,
        [int]$Index = 1
    )

    if ($null -eq $Workbook) { throw 'Workbook es nulo' }
    $sheets = $null
    try {
        $sheets = $Workbook.Worksheets
        $ws     = $sheets.Item($Index)
        return $ws
    } finally {
        if ($null -ne $sheets) { Release-ComObject $sheets; $sheets = $null }
    }
}

function Add-ExcelSheet {
    <#
    .SYNOPSIS
        Agrega una hoja nueva al workbook liberando la coleccion Worksheets.
        Devuelve el objeto Worksheet; el llamante es responsable de liberarlo.
    #>
    param($Workbook)

    if ($null -eq $Workbook) { throw 'Workbook es nulo' }
    $sheets = $null
    try {
        $sheets = $Workbook.Worksheets
        $ws     = $sheets.Add()
        return $ws
    } finally {
        if ($null -ne $sheets) { Release-ComObject $sheets; $sheets = $null }
    }
}

function Save-ExcelWorkbook {
    <#
    .SYNOPSIS
        Guarda un workbook en la ruta indicada.
        Formato 51 = xlOpenXMLWorkbook (.xlsx).
    #>
    param(
        [hashtable]$Context,
        $Workbook,
        [string]$Path,
        [int]$Format = 51
    )

    if ($null -eq $Workbook) { throw 'Workbook es nulo' }
    try {
        $Workbook.SaveAs($Path, $Format)
        Write-BuildLog $Context 'INFO' "Workbook guardado: $Path"
    } catch {
        Write-BuildLog $Context 'ERROR' "Error al guardar workbook en $Path : $_"
        throw
    }
}

function Invoke-ExcelAutoFit {
    <#
    .SYNOPSIS
        Ajusta el ancho de todas las columnas de una hoja.
        Captura y libera el objeto Columns intermedio.
    #>
    param($Sheet)

    if ($null -eq $Sheet) { return }
    $cols = $null
    try {
        $cols = $Sheet.Columns
        $cols.AutoFit() | Out-Null
    } catch {
        # AutoFit no es critico
    } finally {
        if ($null -ne $cols) { Release-ComObject $cols; $cols = $null }
    }
}

function Get-ExcelUsedRange {
    <#
    .SYNOPSIS
        Devuelve (rowCount, colCount) del rango usado sin dejar COM objects vivos.
    .OUTPUTS
        Hashtable @{ Rows = int; Cols = int }
    #>
    param($Sheet)

    $used = $null
    $rows = $null
    $cols = $null
    try {
        $used     = $Sheet.UsedRange
        $rowsObj  = $used.Rows
        $colsObj  = $used.Columns
        $rowCount = $rowsObj.Count
        $colCount = $colsObj.Count
        return @{ Rows = $rowCount; Cols = $colCount }
    } finally {
        if ($null -ne $rowsObj) { Release-ComObject $rowsObj }
        if ($null -ne $colsObj) { Release-ComObject $colsObj }
        if ($null -ne $used)    { Release-ComObject $used }
    }
}

function Get-ExcelCellValue {
    <#
    .SYNOPSIS
        Lee el valor de texto de una celda y libera el objeto COM inmediatamente.
    #>
    param(
        $Sheet,
        [int]$Row,
        [int]$Col
    )

    $cell = $null
    try {
        $cell = $Sheet.Cells($Row, $Col)
        return "$($cell.Text)"
    } finally {
        if ($null -ne $cell) { Release-ComObject $cell; $cell = $null }
    }
}

function Write-ExcelData {
    <#
    .SYNOPSIS
        Escribe un array de hashtables en una hoja Excel.
        Cada celda se captura como variable y se libera inmediatamente.
    #>
    param(
        [hashtable]$Context,
        $Sheet,
        [array]$Data,
        [int]$StartRow = 1,
        [string[]]$Columns = @()
    )

    if ($null -eq $Sheet) { throw 'Sheet es nulo' }
    if ($Data.Count -eq 0) {
        Write-BuildLog $Context 'WARN' 'Write-ExcelData: array de datos vacio'
        return
    }

    if ($Columns.Count -eq 0) {
        $Columns = @($Data[0].Keys | Sort-Object)
    }

    for ($c = 0; $c -lt $Columns.Count; $c++) {
        $cell = $Sheet.Cells($StartRow, $c + 1)
        $cell.Value2 = $Columns[$c]
        Release-ComObject $cell
    }

    for ($r = 0; $r -lt $Data.Count; $r++) {
        for ($c = 0; $c -lt $Columns.Count; $c++) {
            $val  = $Data[$r][$Columns[$c]]
            $cell = $Sheet.Cells($StartRow + $r + 1, $c + 1)
            $cell.Value2 = if ($null -eq $val) { '' } else { "$val" }
            Release-ComObject $cell
        }
    }

    Write-BuildLog $Context 'DEBUG' "Escritas $($Data.Count) filas en Excel"
}

function Close-ExcelWorkbook {
    <#
    .SYNOPSIS
        Cierra un workbook, drena su referencia COM y dispara un ciclo GC.
        El GC es necesario para que el ref-count interno de .NET llegue a
        cero antes de llamar a Close-ExcelApp.
    #>
    param(
        $Workbook,
        [bool]$Save = $false
    )

    if ($null -eq $Workbook) { return }
    try { $Workbook.Close($Save) } catch {}
    Release-ComObject $Workbook

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function Close-ExcelApp {
    <#
    .SYNOPSIS
        Cierra Excel, drena referencias COM, espera confirmacion del OS.
        Si el proceso sigue vivo al agotar el timeout lo elimina con Kill().
        Garantia: no quedan procesos EXCEL headless tras esta llamada.
    #>
    param(
        $ExcelApp,
        [int]$WaitSec = 15
    )

    if ($null -eq $ExcelApp) { return }
    try { $ExcelApp.Quit() } catch {}
    Release-ComObject $ExcelApp

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()

    $deadline = [datetime]::Now.AddSeconds($WaitSec)
    while ([datetime]::Now -lt $deadline) {
        $vivos = @(Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowHandle -eq [IntPtr]::Zero })
        if ($vivos.Count -eq 0) { return }
        Start-Sleep -Milliseconds 400
    }

    # Timeout agotado: forzar eliminacion de cualquier proceso headless restante
    @(Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowHandle -eq [IntPtr]::Zero }) | ForEach-Object {
        try { $_.Kill() } catch {}
    }
    Start-Sleep -Milliseconds 500
}
