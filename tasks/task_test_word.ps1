#Requires -Version 5.1
# =============================================================================
# tasks/task_test_word.ps1
# @Description : Prueba funcional Word COM - documento, parrafos, tabla, PDF
# @Category    : Test
# @Version     : 1.0.0
# @Author      : AutoBuild QA
# @Environment : Office 16, PS 5.1 Desktop
# =============================================================================
# Synopsis: Prueba funcional Word - crea documento, escribe parrafos y tabla, exporta PDF
# Params: {"OpenResult":"false"}

task test_word {

    $ctx = New-TaskContext `
        -TaskName 'test_word' `
        -Config   $Script:EngineConfig `
        -Root     $Script:EngineRoot `
        -Params   $Script:TaskParams

    $openResult = $false
    try { $openResult = [bool]::Parse([string]$ctx.Params['OpenResult']) } catch {}

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

    $wd      = $null
    $doc     = $null
    $sel     = $null
    $docPath = $null
    $pdfPath = $null

    Write-Build Cyan "`n  TEST WORD COM`n"
    Write-BuildLog $ctx 'INFO' 'test_word iniciado'

    # ---- 1. DISPONIBILIDAD -------------------------------------------------
    Write-Build Cyan '  -- [1/5] Disponibilidad COM'
    $avail = Test-ComAvailable -ProgId 'Word.Application' -TimeoutSec 20
    Chk 'Word.Application COM disponible' $avail
    if (-not $avail) {
        Write-Build Yellow '  Word no disponible. Abortando test.'
        Write-RunResult -Context $ctx -Success $false -ErrorMsg 'Word COM no disponible'
        return
    }

    try {

        # ---- 2. INSTANCIAR Y CREAR DOCUMENTO --------------------------------
        Write-Build Cyan '  -- [2/5] Instanciar Word y crear documento'

        $wd = New-WordApp -Context $ctx -TimeoutSec 30
        Chk 'New-WordApp' ($null -ne $wd)
        if ($null -eq $wd) { throw 'Word app null' }

        $doc = New-WordDocument -Context $ctx -WordApp $wd
        Chk 'New-WordDocument' ($null -ne $doc)
        if ($null -eq $doc) { throw 'Document null' }

        $sel = Get-WordSelection -Context $ctx -WordApp $wd
        Chk 'Get-WordSelection' ($null -ne $sel)
        if ($null -eq $sel) { throw 'Selection null' }

        # ---- 3. ESCRIBIR CONTENIDO -----------------------------------------
        Write-Build Cyan '  -- [3/5] Escribir parrafos y tabla'

        # Titulo
        Add-WordParagraph -Context $ctx -Selection $sel -Text 'AutoBuild - Reporte de Prueba Funcional Word'
        Add-WordParagraph -Context $ctx -Selection $sel -Text ''
        Add-WordParagraph -Context $ctx -Selection $sel -Text ('Generado: ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
        Add-WordParagraph -Context $ctx -Selection $sel -Text ('RunId: ' + $ctx.RunId)
        Add-WordParagraph -Context $ctx -Selection $sel -Text ('Usuario: ' + $ctx.User + ' en ' + $ctx.Hostname)
        Add-WordParagraph -Context $ctx -Selection $sel -Text ''

        # Parrafos de contenido
        $paragraphsToWrite = [System.Collections.Generic.List[string]]::new()
        $paragraphsToWrite.Add('Seccion 1: Resumen Ejecutivo')
        $paragraphsToWrite.Add('Este documento fue generado automaticamente por el framework AutoBuild v3.3 como parte de la prueba funcional de integracion COM de Word. El objetivo es verificar que la escritura de parrafos, tablas y la exportacion a PDF funcionan correctamente bajo Office 16 y PowerShell 5.1.')
        $paragraphsToWrite.Add('')
        $paragraphsToWrite.Add('Seccion 2: Parametros de Entorno')
        $paragraphsToWrite.Add('PowerShell Version: ' + $PSVersionTable.PSVersion.ToString())
        $paragraphsToWrite.Add('OS Build: ' + [System.Environment]::OSVersion.Version.Build.ToString())
        $paragraphsToWrite.Add('CLR Version: ' + [System.Environment]::Version.ToString())
        $paragraphsToWrite.Add('')
        $paragraphsToWrite.Add('Seccion 3: Verificaciones Realizadas')
        for ($li = 1; $li -le 10; $li++) {
            $paragraphsToWrite.Add(('Verificacion {0:D2}: Item de prueba funcional -- {1}' -f $li, (Get-Date -Format 'HH:mm:ss.fff')))
        }
        $paragraphsToWrite.Add('')
        $paragraphsToWrite.Add('Seccion 4: Conclusiones')
        $paragraphsToWrite.Add('Si este documento existe y contiene el contenido esperado, la integracion COM de Word esta operativa.')

        $linesArr = $paragraphsToWrite.ToArray()
        foreach ($line in $linesArr) {
            Add-WordParagraph -Context $ctx -Selection $sel -Text $line
        }
        Chk ('Parrafos escritos ({0} lineas)' -f $linesArr.Count) $true ''

        # Verificar char count
        $charCount = 0
        try {
            $charCount = Get-WordCharCount -Context $ctx -Document $doc
            Chk 'Documento tiene contenido (charCount > 0)' ($charCount -gt 100) `
                ('CharCount={0}' -f $charCount)
        } catch {
            Chk 'Documento tiene contenido' $false ('{0}' -f $_)
        }

        # Tabla simple via COM directo
        $tableOk = $false
        $tblCom  = $null
        $tablesCol = $null
        try {
            # Agregar parrafo antes de la tabla
            Add-WordParagraph -Context $ctx -Selection $sel -Text ''
            Add-WordParagraph -Context $ctx -Selection $sel -Text 'Tabla de Resultados:'
            Add-WordParagraph -Context $ctx -Selection $sel -Text ''

            # Insertar tabla 4 filas x 3 columnas via Selection.Range
            $tablesCol = $doc.Tables
            $selRange  = $sel.Range
            $tblCom    = $tablesCol.Add($selRange, 4, 3)
            if ($null -ne $tblCom) {
                # Cabecera
                $headers = @('ID','Descripcion','Valor')
                for ($c = 1; $c -le 3; $c++) {
                    $cellCom = $null
                    try {
                        $cellCom = $tblCom.Cell(1, $c)
                        $cellCom.Range.Text = $headers[$c - 1]
                    } finally { if ($null -ne $cellCom) { Invoke-ReleaseComObject $cellCom } }
                }
                # Filas de datos
                for ($r = 2; $r -le 4; $r++) {
                    $rowData = @([string]($r-1), ('Item {0}' -f ($r-1)), [string](($r-1) * 100))
                    for ($c = 1; $c -le 3; $c++) {
                        $cellCom = $null
                        try {
                            $cellCom = $tblCom.Cell($r, $c)
                            $cellCom.Range.Text = $rowData[$c - 1]
                        } finally { if ($null -ne $cellCom) { Invoke-ReleaseComObject $cellCom } }
                    }
                }
                $tableOk = $true
            }
        } catch {
            Write-BuildLog $ctx 'WARN' ('Tabla COM: {0}' -f $_)
        } finally {
            if ($null -ne $tblCom)    { Invoke-ReleaseComObject $tblCom    }
            if ($null -ne $tablesCol) { Invoke-ReleaseComObject $tablesCol }
        }
        Chk 'Tabla 4x3 insertada' $tableOk ''

        # Parrafo final
        Add-WordParagraph -Context $ctx -Selection $sel -Text ''
        Add-WordParagraph -Context $ctx -Selection $sel -Text '-- Fin del documento de prueba --'

        # ---- 4. GUARDAR DOCX -----------------------------------------------
        Write-Build Cyan '  -- [4/5] Guardar .docx'
        $docPath = Join-Path $ctx.Paths.Reports ('test_word_{0}.docx' -f $stamp)

        $saveOk = $false
        try {
            [string]$docPathStr = $docPath
            $doc.SaveAs([ref]$docPathStr, [ref]16)    # 16 = wdFormatXMLDocument
            $saveOk = Test-Path $docPath
        } catch {
            Write-BuildLog $ctx 'ERROR' ('SaveAs docx: {0}' -f $_)
        }
        Chk 'Guardado como .docx' $saveOk $(if ($saveOk) { $docPath } else { 'SaveAs fallo' })

        # Verificar tamano minimo (debe tener contenido)
        if ($saveOk) {
            $fileSize = (Get-Item $docPath).Length
            Chk 'Archivo .docx tiene tamano > 5KB' ($fileSize -gt 5120) `
                ('Tamano={0}KB' -f [math]::Round($fileSize/1KB, 1))
        } else {
            Chk 'Archivo .docx tiene tamano > 5KB' $false 'Omitido: archivo no guardado'
        }

        # ---- 5. EXPORTAR PDF -----------------------------------------------
        Write-Build Cyan '  -- [5/5] Exportar a PDF'
        $pdfPath = Join-Path $ctx.Paths.Reports ('test_word_{0}.pdf' -f $stamp)

        $pdfOk = $false
        try {
            Export-WordToPdf -Context $ctx -Document $doc -Path $pdfPath
            $pdfOk = Test-Path $pdfPath
        } catch {
            Write-BuildLog $ctx 'ERROR' ('Export PDF: {0}' -f $_)
        }
        Chk 'Exportado como .pdf' $pdfOk $(if ($pdfOk) { $pdfPath } else { 'ExportAsFixedFormat fallo' })

        if ($pdfOk) {
            $pdfSize = (Get-Item $pdfPath).Length
            Chk 'Archivo .pdf tiene tamano > 5KB' ($pdfSize -gt 5120) `
                ('Tamano={0}KB' -f [math]::Round($pdfSize/1KB, 1))
        } else {
            Chk 'Archivo .pdf tiene tamano > 5KB' $false 'Omitido: PDF no generado'
        }

    } catch {
        Write-BuildLog $ctx 'ERROR' ('test_word fallo: {0}' -f $_) -Detail $_.ScriptStackTrace
        $script:failed++
    } finally {
        if ($null -ne $sel) { try { Invoke-ReleaseComObject $sel } catch {} }
        if ($null -ne $doc) { try { $doc.Close($false) } catch {}; try { Invoke-ReleaseComObject $doc } catch {} }
        if ($null -ne $wd)  { try { $wd.Quit()         } catch {}; try { Invoke-ReleaseComObject $wd  } catch {} }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
    }

    # ---- RESUMEN -----------------------------------------------------------
    Write-Build Cyan ("`n  TEST WORD: Total={0} Pasaron={1} Fallaron={2}" -f
                      $results.Count, $script:passed, $script:failed)

    $jsonOut = Join-Path $ctx.Paths.Reports ('test_word_{0}.json' -f $stamp)
    try {
        @{ runId=$ctx.RunId; passed=$script:passed; failed=$script:failed
           docxFile=$docPath; pdfFile=$pdfPath; checks=$results.ToArray() } |
            ConvertTo-Json -Depth 3 |
            ForEach-Object { [System.IO.File]::WriteAllText($jsonOut, $_, [System.Text.Encoding]::UTF8) }
        Write-Build Cyan ('  Reporte: {0}' -f $jsonOut)
    } catch {}

    if ($openResult -and $null -ne $docPath -and (Test-Path $docPath)) {
        try { Start-Process $docPath } catch {}
    }

    if ($script:failed -gt 0) {
        Write-RunResult -Context $ctx -Success $false `
            -ErrorMsg ('{0} verificacion(es) fallida(s)' -f $script:failed)
        throw ('test_word: {0} check(s) failed' -f $script:failed)
    }
    Write-RunResult -Context $ctx -Success $true
}
