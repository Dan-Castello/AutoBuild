#Requires -Version 5.1
# =============================================================================
# tasks/task_diag_multi_com.ps1
# @Description : Diagnostico de interaccion multi-COM: abre Excel y Word
#                de forma secuencial en la misma tarea, transfiere datos
#                de Excel a Word, verifica liberacion limpia de ambos.
#                Prueba el patron real de tareas de produccion que necesitan
#                mas de una aplicacion Office.
# @Category    : Utilidad
# @Version     : 1.0.0
# =============================================================================
# Solo ASCII. PS 5.1.
# REGLA: solo una instancia COM activa a la vez (Regla 6).
# Excel primero -> liberar -> Word despues.

# Synopsis: Prueba apertura secuencial Excel+Word, transferencia de datos y liberacion
task diag_multi_com {
    $ctx = New-TaskContext `
        -TaskName 'diag_multi_com' `
        -Config   $Script:Config `
        -Root     (Split-Path $BuildRoot -Parent)

    Write-BuildLog $ctx 'INFO' 'Iniciando diagnostico multi-COM (Excel + Word secuencial)'

    $errores   = 0
    $datosExcel = @()

    # ================================================================
    # FASE 1 - Excel: crear datos y exportar
    # ================================================================
    Write-Build Cyan  ""
    Write-Build Cyan  "  [FASE 1] Excel - generacion de datos"
    $xl = $null
    $wb = $null
    $ws = $null
    $xlFile = Join-Path $ctx.Paths.Output "diag_multi_com_datos_$($ctx.RunId).xlsx"
    try {
        if (-not (Test-ComAvailable -ProgId 'Excel.Application' -TimeoutSec 15)) {
            throw 'Excel COM no disponible'
        }

        $xl = New-ExcelApp -Context $ctx -TimeoutSec 30
        if ($null -eq $xl) { throw 'New-ExcelApp devolvio null' }
        Write-Build Green "  [XL]   Instancia Excel    : OK"

        $wb = New-ExcelWorkbook -Context $ctx -ExcelApp $xl
        $ws = Get-ExcelSheet -Workbook $wb -Index 1
        $ws.Name = 'Datos'

        $datosExcel = @(
            @{ Producto = 'Item-A'; Cantidad = '10'; Precio = '100.00'; Total = '1000.00' },
            @{ Producto = 'Item-B'; Cantidad = '5';  Precio = '250.00'; Total = '1250.00' },
            @{ Producto = 'Item-C'; Cantidad = '8';  Precio = '75.00';  Total = '600.00'  }
        )

        Write-ExcelData -Context $ctx -Sheet $ws -Data $datosExcel `
            -Columns @('Producto','Cantidad','Precio','Total')

        Invoke-ExcelAutoFit -Sheet $ws
        Release-ComObject $ws ; $ws = $null

        Save-ExcelWorkbook -Context $ctx -Workbook $wb -Path $xlFile
        Write-Build Green "  [XL]   Datos guardados    : OK -> $xlFile"

        Write-Build Green "  [XL]   Datos en memoria   : $($datosExcel.Count) filas listos para Word"

    } catch {
        $errores++
        Write-Build Red "  [XL]   FAIL: $_"
        Write-BuildLog $ctx 'ERROR' "Fase Excel fallo: $_"
    } finally {
        if ($null -ne $ws) { Release-ComObject $ws ; $ws = $null }
        Close-ExcelWorkbook -Workbook $wb -Save $false
        Close-ExcelApp      -ExcelApp $xl
        $wb = $null
        $xl = $null
        Write-Build Green "  [XL]   Liberado           : OK"

        $zombisXl = @(Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowHandle -eq [IntPtr]::Zero }).Count
        if ($zombisXl -gt 0) {
            Write-Build Yellow "  [XL]   Zombis detectados  : $zombisXl"
        }
    }

    # ================================================================
    # FASE 2 - Word: redactar reporte con los datos de Excel
    # ================================================================
    Write-Build Cyan  ""
    Write-Build Cyan  "  [FASE 2] Word - redaccion de reporte con datos de Excel"
    $wd  = $null
    $doc = $null
    $docxFile = Join-Path $ctx.Paths.Output "diag_multi_com_reporte_$($ctx.RunId).docx"
    $sel      = $null
    try {
        if (-not (Test-ComAvailable -ProgId 'Word.Application' -TimeoutSec 15)) {
            throw 'Word COM no disponible'
        }

        $wd = New-WordApp -Context $ctx -TimeoutSec 30
        if ($null -eq $wd) { throw 'New-WordApp devolvio null' }
        Write-Build Green "  [WD]   Instancia Word     : OK"

        $doc = New-WordDocument -Context $ctx -WordApp $wd
        $sel = Get-WordSelection -WordApp $wd

        Add-WordParagraph -Context $ctx -Selection $sel -Text "REPORTE DE DIAGNOSTICO MULTI-COM"
        Add-WordParagraph -Context $ctx -Selection $sel -Text "RunId : $($ctx.RunId)"
        Add-WordParagraph -Context $ctx -Selection $sel -Text "Fecha : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        Add-WordParagraph -Context $ctx -Selection $sel -Text ""
        Add-WordParagraph -Context $ctx -Selection $sel -Text "Datos transferidos desde Excel:"
        foreach ($fila in $datosExcel) {
            $linea = "  $($fila.Producto)  Cant:$($fila.Cantidad)  Precio:$($fila.Precio)  Total:$($fila.Total)"
            Add-WordParagraph -Context $ctx -Selection $sel -Text $linea
        }

        Release-ComObject $sel ; $sel = $null

        $charCount = Get-WordCharCount -Document $doc
        Write-Build Green "  [WD]   Contenido          : OK ($charCount caracteres)"

        Save-WordDocument -Context $ctx -Document $doc -Path $docxFile

        if (Test-Path $docxFile) {
            $sz = (Get-Item $docxFile).Length
            Write-Build Green "  [WD]   .docx guardado    : OK ($sz bytes)"
        } else {
            Write-Build Red   "  [WD]   .docx guardado    : FAIL"
            $errores++
        }

        $pdfFile = Join-Path $ctx.Paths.Output "diag_multi_com_reporte_$($ctx.RunId).pdf"
        try {
            Export-WordToPdf -Context $ctx -Document $doc -Path $pdfFile
            if (Test-Path $pdfFile) {
                Write-Build Green "  [WD]   PDF exportado     : OK ($($(Get-Item $pdfFile).Length) bytes)"
            }
        } catch {
            Write-Build Yellow "  [WD]   PDF exportado     : WARN - $_"
        }

    } catch {
        $errores++
        Write-Build Red "  [WD]   FAIL: $_"
        Write-BuildLog $ctx 'ERROR' "Fase Word fallo: $_"
    } finally {
        if ($null -ne $sel) { Release-ComObject $sel; $sel = $null }

        Close-WordDocument -Document $doc -Save $false
        Close-WordApp      -WordApp $wd
        $doc = $null
        $wd  = $null
        Write-Build Green "  [WD]   Liberado           : OK"

        $zombisWd = @(Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowHandle -eq [IntPtr]::Zero }).Count
        if ($zombisWd -gt 0) {
            Write-Build Yellow "  [WD]   Zombis detectados  : $zombisWd"
        }
    }

    # ---- Resumen -----------------------------------------------------------
    Write-Build Cyan  ""
    Write-Build Cyan  "  Archivos generados:"
    foreach ($f in @($xlFile, $docxFile)) {
        if (Test-Path $f) {
            Write-Build Cyan  "    $([System.IO.Path]::GetFileName($f)) ($($(Get-Item $f).Length) bytes)"
        }
    }

    Write-Build Cyan ""
    if ($errores -eq 0) {
        Write-Build Green "  RESULTADO: OK - Excel y Word operaron secuencialmente sin conflictos"
    } else {
        Write-Build Red   "  RESULTADO: $errores error(es) en la prueba multi-COM"
    }

    Write-BuildLog $ctx 'INFO' "diag_multi_com completado. Errores=$errores"
    Write-RunResult -Context $ctx -Success ($errores -eq 0)

    if ($errores -gt 0) {
        throw "diag_multi_com detecto $errores error(es)"
    }
}
