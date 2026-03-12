#Requires -Version 5.1
# =============================================================================
# tasks/task_diag_pdf.ps1
# @Description : Diagnostico de generacion de PDF en el entorno.
#                Prueba tres rutas: Word->PDF (ExportAsFixedFormat),
#                Excel->PDF (ExportAsFixedFormat) y Microsoft Print to PDF
#                via Word como fallback.
# @Category    : Utilidad
# @Version     : 1.0.0
# =============================================================================
# Solo ASCII. PS 5.1.

# Synopsis: Verifica generacion de PDF desde Word y Excel (Office 16)
task diag_pdf {
    $ctx = New-TaskContext `
        -TaskName 'diag_pdf' `
        -Config   $Script:Config `
        -Root     (Split-Path $BuildRoot -Parent)

    Write-BuildLog $ctx 'INFO' 'Iniciando diagnostico de PDF'

    $errores = 0
    $avisos  = 0

    # ---- 1. Verificar impresora PDF disponible -----------------------------
    Write-Build Cyan  "  [PRINT] Verificando impresoras PDF disponibles..."
    $impresorasPdf = @(Get-WmiObject -Class Win32_Printer -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like '*PDF*' -or $_.Name -like '*pdf*' })

    if ($impresorasPdf.Count -gt 0) {
        foreach ($imp in $impresorasPdf) {
            Write-Build Green "  [PRINT] $($imp.Name) : OK"
        }
    } else {
        Write-Build Yellow "  [PRINT] Sin impresoras PDF detectadas (Word/Excel pueden exportar sin impresora)"
        $avisos++
    }

    # ---- 2. PDF via Word (ExportAsFixedFormat) ------------------------------
    Write-Build Cyan  ""
    Write-Build Cyan  "  [WORD->PDF] Probando exportacion PDF desde Word..."
    $wd  = $null
    $doc = $null
    try {
        $comOk = Test-ComAvailable -ProgId 'Word.Application' -TimeoutSec 15
        if (-not $comOk) {
            Write-Build Yellow "  [WORD->PDF] Word COM no disponible - saltando prueba Word"
            $avisos++
        } else {
            $wd = New-WordApp -Context $ctx -TimeoutSec 30
            if ($null -eq $wd) { throw 'New-WordApp devolvio null' }

            $doc = New-WordDocument -Context $ctx -WordApp $wd

            $sel = Get-WordSelection -WordApp $wd
            Add-WordParagraph -Context $ctx -Selection $sel -Text "Diagnostico PDF - AutoBuild"
            Add-WordParagraph -Context $ctx -Selection $sel -Text "RunId: $($ctx.RunId)"
            Add-WordParagraph -Context $ctx -Selection $sel -Text "Fecha: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

            $pdfWord = Join-Path $ctx.Paths.Output "diag_pdf_word_$($ctx.RunId).pdf"
            Export-WordToPdf -Context $ctx -Document $doc -Path $pdfWord

            if (Test-Path $pdfWord) {
                $sz = (Get-Item $pdfWord).Length
                # PDF debe empezar con %PDF
                $header = [System.IO.File]::ReadAllBytes($pdfWord) | Select-Object -First 4
                $isPdf  = ($header[0] -eq 0x25 -and $header[1] -eq 0x50 -and
                           $header[2] -eq 0x44 -and $header[3] -eq 0x46)
                if ($isPdf) {
                    Write-Build Green "  [WORD->PDF] Generado      : OK ($sz bytes, firma %PDF verificada)"
                } else {
                    Write-Build Red   "  [WORD->PDF] Firma PDF invalida (primeros bytes incorrectos)"
                    $errores++
                }
            } else {
                Write-Build Red   "  [WORD->PDF] Archivo no generado"
                $errores++
            }
        }
    } catch {
        Write-Build Red   "  [WORD->PDF] FAIL: $_"
        $errores++
    } finally {
        Close-WordDocument -Document $doc -Save $false
        Close-WordApp      -WordApp $wd
        $doc = $null
        $wd  = $null
    }

    # ---- 3. PDF via Excel (ExportAsFixedFormat) ----------------------------
    Write-Build Cyan  ""
    Write-Build Cyan  "  [XL->PDF] Probando exportacion PDF desde Excel..."
    $xl = $null
    $wb = $null
    try {
        $comOk2 = Test-ComAvailable -ProgId 'Excel.Application' -TimeoutSec 15
        if (-not $comOk2) {
            Write-Build Yellow "  [XL->PDF] Excel COM no disponible - saltando prueba Excel"
            $avisos++
        } else {
            $xl = New-ExcelApp -Context $ctx -TimeoutSec 30
            if ($null -eq $xl) { throw 'New-ExcelApp devolvio null' }

            $wb = New-ExcelWorkbook -Context $ctx -ExcelApp $xl
            $ws = Get-ExcelSheet -Workbook $wb -Index 1
            $ws.Name = 'DiagPDF'

            $datos = @(
                @{ Campo = 'Tarea';  Valor = 'diag_pdf'              },
                @{ Campo = 'RunId';  Valor = $ctx.RunId              },
                @{ Campo = 'Fecha';  Valor = (Get-Date -Format 'yyyy-MM-dd') }
            )
            Write-ExcelData -Context $ctx -Sheet $ws -Data $datos -Columns @('Campo','Valor')

            $pdfXl = Join-Path $ctx.Paths.Output "diag_pdf_excel_$($ctx.RunId).pdf"

            # xlTypePDF=0, xlQualityStandard=0, IncludeDocProperties=true,
            # IgnorePrintAreas=false, OpenAfterPublish=false
            $ws.ExportAsFixedFormat(0, $pdfXl, 0, $true, $false)

            if (Test-Path $pdfXl) {
                $sz2  = (Get-Item $pdfXl).Length
                $hdr2 = [System.IO.File]::ReadAllBytes($pdfXl) | Select-Object -First 4
                $isPdf2 = ($hdr2[0] -eq 0x25 -and $hdr2[1] -eq 0x50 -and
                           $hdr2[2] -eq 0x44 -and $hdr2[3] -eq 0x46)
                if ($isPdf2) {
                    Write-Build Green "  [XL->PDF]  Generado      : OK ($sz2 bytes, firma %PDF verificada)"
                } else {
                    Write-Build Red   "  [XL->PDF]  Firma PDF invalida"
                    $errores++
                }
            } else {
                Write-Build Red   "  [XL->PDF]  Archivo no generado"
                $errores++
            }
        }
    } catch {
        Write-Build Red   "  [XL->PDF] FAIL: $_"
        $errores++
    } finally {
        Close-ExcelWorkbook -Workbook $wb -Save $false
        Close-ExcelApp      -ExcelApp $xl
        $wb = $null
        $xl = $null
    }

    # ---- 4. Inventario de PDFs generados -----------------------------------
    Write-Build Cyan  ""
    $pdfsGenerados = @(Get-ChildItem -Path $ctx.Paths.Output -Filter "diag_pdf_*_$($ctx.RunId).pdf" -ErrorAction SilentlyContinue)
    Write-Build Cyan  "  [INV]  PDFs generados en esta ejecucion: $($pdfsGenerados.Count)"
    foreach ($pdf in $pdfsGenerados) {
        Write-Build Cyan  "         $($pdf.Name) ($($pdf.Length) bytes)"
    }

    # ---- Resumen -----------------------------------------------------------
    Write-Build Cyan ""
    if ($errores -eq 0 -and $avisos -eq 0) {
        Write-Build Green "  RESULTADO: OK - generacion de PDF funciona correctamente"
    } elseif ($errores -eq 0) {
        Write-Build Yellow "  RESULTADO: $avisos aviso(s) - alguna ruta no disponible"
    } else {
        Write-Build Red   "  RESULTADO: $errores error(es), $avisos aviso(s)"
    }

    Write-BuildLog $ctx 'INFO' "diag_pdf completado. Errores=$errores Avisos=$avisos"
    Write-RunResult -Context $ctx -Success ($errores -eq 0)

    if ($errores -gt 0) {
        throw "diag_pdf detecto $errores error(es)"
    }
}
