#Requires -Version 5.1
# =============================================================================
# tasks/task_test_pdf.ps1
# @Description : Prueba funcional PDF - generacion via Word COM y validacion
# @Category    : Test
# @Version     : 1.0.0
# @Author      : AutoBuild QA
# @Environment : Office 16, PS 5.1 Desktop
# =============================================================================
# Synopsis: Prueba funcional PDF - genera documentos Word y los exporta a PDF
# Params: {"OpenResult":"false","PageCount":"3"}
#
# NOTA: En PS 5.1 corporativo sin librerias externas (ClosedXML, iTextSharp),
# la forma mas robusta de generar PDF es via Word COM ExportAsFixedFormat.
# Esta tarea verifica ese pipeline completo con multiples documentos y
# validacion de tamano/existencia de los PDFs resultantes.

task test_pdf {

    $ctx = New-TaskContext `
        -TaskName 'test_pdf' `
        -Config   $Script:EngineConfig `
        -Root     $Script:EngineRoot `
        -Params   $Script:TaskParams

    $openResult = $false
    try { $openResult = [bool]::Parse([string]$ctx.Params['OpenResult']) } catch {}

    $pageCount = 3
    try { $pv = $ctx.Params['PageCount']; if ($pv -gt '') { $pageCount = [int]$pv } } catch {}
    if ($pageCount -lt 1) { $pageCount = 3 }
    if ($pageCount -gt 10) { $pageCount = 10 }

    $stamp         = Get-Date -Format 'yyyyMMdd_HHmmss'
    $script:passed = 0
    $script:failed = 0
    $results       = [System.Collections.Generic.List[hashtable]]::new()

    function Chk {
        param([string]$Name, [bool]$Pass, [string]$Detail = '')
        $results.Add(@{ Name=$Name; Pass=$Pass; Detail=[string]$Detail })
        $sym   = if ($Pass) { '[OK]  ' } else { '[FAIL]' }
        $color = if ($Pass) { 'Green' } else { 'Red' }
        Write-Build $color ('  {0} {1}{2}' -f $sym, $Name,
                             $(if ($Detail) { ' -- '+$Detail } else { '' }))
        Write-BuildLog $ctx $(if($Pass){'INFO'}else{'ERROR'}) "$sym $Name" -Detail $Detail
        if ($Pass) { $script:passed++ } else { $script:failed++ }
    }

    $wd       = $null
    $pdfFiles = [System.Collections.Generic.List[string]]::new()

    Write-Build Cyan ("`n  TEST PDF -- {0} pagina(s) por documento`n" -f $pageCount)
    Write-BuildLog $ctx 'INFO' ('test_pdf iniciado, pageCount={0}' -f $pageCount)

    # ---- 1. DISPONIBILIDAD -------------------------------------------------
    Write-Build Cyan '  -- [1/5] Disponibilidad Word COM (requerido para PDF)'
    $avail = Test-ComAvailable -ProgId 'Word.Application' -TimeoutSec 20
    Chk 'Word.Application COM disponible (pipeline PDF)' $avail
    if (-not $avail) {
        Write-Build Yellow '  Word no disponible. Abortando test.'
        Write-RunResult -Context $ctx -Success $false -ErrorMsg 'Word COM no disponible para PDF'
        return
    }

    try {
        $wd = New-WordApp -Context $ctx -TimeoutSec 30
        Chk 'New-WordApp' ($null -ne $wd)
        if ($null -eq $wd) { throw 'Word app null' }

        # ---- 2. DOCUMENTO SIMPLE -------------------------------------------
        Write-Build Cyan '  -- [2/5] PDF desde documento simple'

        $doc1    = $null
        $sel1    = $null
        $pdf1    = $null
        try {
            $doc1 = New-WordDocument -Context $ctx -WordApp $wd
            $sel1 = Get-WordSelection -Context $ctx -WordApp $wd

            Add-WordParagraph -Context $ctx -Selection $sel1 -Text 'AutoBuild - PDF Test Document 1'
            Add-WordParagraph -Context $ctx -Selection $sel1 -Text ('Generado: ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
            Add-WordParagraph -Context $ctx -Selection $sel1 -Text ('RunId: ' + $ctx.RunId)
            Add-WordParagraph -Context $ctx -Selection $sel1 -Text ''
            Add-WordParagraph -Context $ctx -Selection $sel1 -Text 'Este documento verifica la generacion de PDF via Word COM ExportAsFixedFormat.'

            # Agregar N paginas de contenido
            for ($p = 1; $p -le $pageCount; $p++) {
                Add-WordParagraph -Context $ctx -Selection $sel1 -Text ''
                Add-WordParagraph -Context $ctx -Selection $sel1 -Text ('--- Pagina {0} de {1} ---' -f $p, $pageCount)
                for ($ln = 1; $ln -le 15; $ln++) {
                    Add-WordParagraph -Context $ctx -Selection $sel1 -Text `
                        ('Linea {0:D3} de la pagina {1}: Lorem ipsum dolor sit amet, consectetur adipiscing elit.' -f $ln, $p)
                }
                # Salto de pagina excepto en la ultima
                if ($p -lt $pageCount) {
                    try { $sel1.InsertBreak(7) } catch {}  # 7 = wdPageBreak
                }
            }

            $pdf1 = Join-Path $ctx.Paths.Reports ('test_pdf_simple_{0}.pdf' -f $stamp)
            $sw   = [System.Diagnostics.Stopwatch]::StartNew()
            Export-WordToPdf -Context $ctx -Document $doc1 -Path $pdf1
            $sw.Stop()
            $pdf1Ok = Test-Path $pdf1
            Chk 'PDF simple generado' $pdf1Ok `
                $(if ($pdf1Ok) { '{0}ms, {1}KB' -f $sw.ElapsedMilliseconds,
                  [math]::Round((Get-Item $pdf1).Length/1KB,1) } else { 'ExportAsFixedFormat fallo' })
            if ($pdf1Ok) { $pdfFiles.Add($pdf1) }

        } catch {
            Write-BuildLog $ctx 'ERROR' ('PDF simple: {0}' -f $_)
            Chk 'PDF simple generado' $false ('{0}' -f $_)
        } finally {
            if ($null -ne $sel1) { try { Invoke-ReleaseComObject $sel1 } catch {} }
            if ($null -ne $doc1) { try { $doc1.Close($false) } catch {}; try { Invoke-ReleaseComObject $doc1 } catch {} }
        }

        # ---- 3. DOCUMENTO CON TABLA ----------------------------------------
        Write-Build Cyan '  -- [3/5] PDF desde documento con tabla'

        $doc2  = $null
        $sel2  = $null
        $pdf2  = $null
        try {
            $doc2 = New-WordDocument -Context $ctx -WordApp $wd
            $sel2 = Get-WordSelection -Context $ctx -WordApp $wd

            Add-WordParagraph -Context $ctx -Selection $sel2 -Text 'AutoBuild - PDF Test Document 2 (con tabla)'
            Add-WordParagraph -Context $ctx -Selection $sel2 -Text ('Fecha: ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
            Add-WordParagraph -Context $ctx -Selection $sel2 -Text ''
            Add-WordParagraph -Context $ctx -Selection $sel2 -Text 'Tabla de datos de prueba:'
            Add-WordParagraph -Context $ctx -Selection $sel2 -Text ''

            # Tabla de 6 filas x 4 columnas
            $tblCom2    = $null
            $tablesCom2 = $null
            $tblOk2     = $false
            try {
                $tablesCom2 = $doc2.Tables
                $selRange2  = $sel2.Range
                $tblCom2    = $tablesCom2.Add($selRange2, 6, 4)
                if ($null -ne $tblCom2) {
                    $hdrs = @('Codigo','Descripcion','Cantidad','Precio')
                    for ($c = 1; $c -le 4; $c++) {
                        $cellCom = $null
                        try { $cellCom = $tblCom2.Cell(1,$c); $cellCom.Range.Text = $hdrs[$c-1] }
                        finally { if ($null -ne $cellCom) { Invoke-ReleaseComObject $cellCom } }
                    }
                    for ($r = 2; $r -le 6; $r++) {
                        $rowData = @(('C{0:D4}' -f ($r-1)), ('Producto {0}' -f ($r-1)),
                                     [string](($r-1)*10), ('{0:F2}' -f (($r-1)*25.99)))
                        for ($c = 1; $c -le 4; $c++) {
                            $cellCom = $null
                            try { $cellCom = $tblCom2.Cell($r,$c); $cellCom.Range.Text = $rowData[$c-1] }
                            finally { if ($null -ne $cellCom) { Invoke-ReleaseComObject $cellCom } }
                        }
                    }
                    $tblOk2 = $true
                }
            } finally {
                if ($null -ne $tblCom2)    { Invoke-ReleaseComObject $tblCom2    }
                if ($null -ne $tablesCom2) { Invoke-ReleaseComObject $tablesCom2 }
            }

            Add-WordParagraph -Context $ctx -Selection $sel2 -Text ''
            Add-WordParagraph -Context $ctx -Selection $sel2 -Text 'Fin del documento de prueba con tabla.'

            $pdf2   = Join-Path $ctx.Paths.Reports ('test_pdf_tabla_{0}.pdf' -f $stamp)
            $sw3    = [System.Diagnostics.Stopwatch]::StartNew()
            Export-WordToPdf -Context $ctx -Document $doc2 -Path $pdf2
            $sw3.Stop()
            $pdf2Ok = Test-Path $pdf2
            Chk 'PDF con tabla generado' $pdf2Ok `
                $(if ($pdf2Ok) { '{0}ms, {1}KB' -f $sw3.ElapsedMilliseconds,
                  [math]::Round((Get-Item $pdf2).Length/1KB,1) } else { 'Export fallo' })
            if ($pdf2Ok) { $pdfFiles.Add($pdf2) }

        } catch {
            Write-BuildLog $ctx 'ERROR' ('PDF tabla: {0}' -f $_)
            Chk 'PDF con tabla generado' $false ('{0}' -f $_)
        } finally {
            if ($null -ne $sel2) { try { Invoke-ReleaseComObject $sel2 } catch {} }
            if ($null -ne $doc2) { try { $doc2.Close($false) } catch {}; try { Invoke-ReleaseComObject $doc2 } catch {} }
        }

        # ---- 4. VALIDAR ARCHIVOS PDF ---------------------------------------
        Write-Build Cyan '  -- [4/5] Validar archivos PDF generados'

        Chk 'Al menos 1 PDF generado' ($pdfFiles.Count -ge 1) `
            ('PDFs={0}' -f $pdfFiles.Count)

        $totalPdfKb = 0.0
        $allPdfsArr = $pdfFiles.ToArray()
        foreach ($pdf in $allPdfsArr) {
            if (Test-Path $pdf) {
                $sz = (Get-Item $pdf).Length
                $totalPdfKb = $totalPdfKb + ($sz / 1KB)
                # Verificar magic bytes PDF (%PDF-)
                $magic = $null
                try {
                    $bytes = [System.IO.File]::ReadAllBytes($pdf)
                    if ($bytes.Count -ge 4) {
                        $magic = [System.Text.Encoding]::ASCII.GetString($bytes[0..3])
                    }
                } catch {}
                $pdfName = [System.IO.Path]::GetFileName($pdf)
                Chk ('PDF valido: {0}' -f $pdfName) ($magic -eq '%PDF') `
                    ('MagicBytes={0}, Tamano={1}KB' -f $magic, [math]::Round($sz/1KB,1))
            }
        }
        Chk 'Tamano total PDFs > 10KB' ($totalPdfKb -gt 10) `
            ('Total={0:F1}KB' -f $totalPdfKb)

        # ---- 5. REPORTE JSON DE PDFS ----------------------------------------
        Write-Build Cyan '  -- [5/5] Reporte de archivos PDF'

        $pdfInfoList = [System.Collections.Generic.List[hashtable]]::new()
        foreach ($pdf in $allPdfsArr) {
            $exists = Test-Path $pdf
            $pdfInfoList.Add(@{
                File    = [System.IO.Path]::GetFileName($pdf)
                Exists  = $exists
                SizeKB  = if ($exists) { [math]::Round((Get-Item $pdf).Length/1KB,1) } else { 0 }
                Path    = $pdf
            })
        }
        Chk 'Lista PDF generada' ($pdfInfoList.Count -gt 0) `
            ('{0} archivos registrados' -f $pdfInfoList.Count)

    } catch {
        Write-BuildLog $ctx 'ERROR' ('test_pdf fallo: {0}' -f $_) -Detail $_.ScriptStackTrace
        $script:failed++
    } finally {
        if ($null -ne $wd) { try { $wd.Quit() } catch {}; try { Invoke-ReleaseComObject $wd } catch {} }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
    }

    # ---- RESUMEN -----------------------------------------------------------
    Write-Build Cyan ("`n  TEST PDF: Total={0} Pasaron={1} Fallaron={2}" -f
                      $results.Count, $script:passed, $script:failed)

    if ($pdfFiles.Count -gt 0) {
        Write-Build Cyan '  Archivos PDF generados:'
        foreach ($pdf in $pdfFiles.ToArray()) {
            Write-Build White ('    ' + $pdf)
        }
    }

    $jsonOut = Join-Path $ctx.Paths.Reports ('test_pdf_{0}.json' -f $stamp)
    try {
        @{ runId=$ctx.RunId; passed=$script:passed; failed=$script:failed
           pdfCount=$pdfFiles.Count; pdfFiles=$pdfFiles.ToArray()
           checks=$results.ToArray() } |
            ConvertTo-Json -Depth 3 |
            ForEach-Object { [System.IO.File]::WriteAllText($jsonOut, $_, [System.Text.Encoding]::UTF8) }
        Write-Build Cyan ('  Reporte: {0}' -f $jsonOut)
    } catch {}

    if ($openResult -and $pdfFiles.Count -gt 0 -and (Test-Path $pdfFiles[0])) {
        try { Start-Process $pdfFiles[0] } catch {}
    }

    if ($script:failed -gt 0) {
        Write-RunResult -Context $ctx -Success $false `
            -ErrorMsg ('{0} verificacion(es) fallida(s)' -f $script:failed)
        throw ('test_pdf: {0} check(s) failed' -f $script:failed)
    }
    Write-RunResult -Context $ctx -Success $true
}
