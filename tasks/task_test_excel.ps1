#Requires -Version 5.1
# =============================================================================
# tasks/task_test_excel.ps1
# @Description : Prueba funcional Excel COM - escritura, formato, lectura, CSV
# @Category    : Test
# @Version     : 1.0.0
# @Author      : AutoBuild QA
# @Environment : Office 16, PS 5.1 Desktop
# =============================================================================
# Synopsis: Prueba funcional Excel - escribe datos, lee back, exporta CSV
# Params: {"RowCount":"100","OpenResult":"false"}
#
# PATRONES PS 5.1 APLICADOS:
#   - Export-Csv en lugar de Range.Value2 = [object[,]] (evita op_Addition COM)
#   - Escritura celda-a-celda via Cells(r,c).Value2
#   - List[hashtable] + .Add() - sin += sobre arrays
#   - UsedRange.Rows.Count directo - sin Get-ExcelUsedRange wrapper
#   - Todas las variables inicializadas antes del try
# =============================================================================

task test_excel {

    $ctx = New-TaskContext `
        -TaskName 'test_excel' `
        -Config   $Script:EngineConfig `
        -Root     $Script:EngineRoot `
        -Params   $Script:TaskParams

    $rowCount  = 100
    try { $rv = $ctx.Params['RowCount']; if ($rv -gt '') { $rowCount = [int]$rv } } catch {}
    if ($rowCount -lt 1) { $rowCount = 100 }

    $openResult = $false
    try { $openResult = [bool]::Parse([string]$ctx.Params['OpenResult']) } catch {}

    $stamp         = Get-Date -Format 'yyyyMMdd_HHmmss'
    $script:passed = 0   # scope:script para que Chk() anidada lo vea
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

    # Variables COM -- inicializadas antes del try
    $xl       = $null
    $wb       = $null
    $ws       = $null
    $wb2      = $null
    $ws2      = $null
    $xlsxPath = $null
    $csvPath  = $null

    Write-Build Cyan "`n  TEST EXCEL COM -- $rowCount filas`n"
    Write-BuildLog $ctx 'INFO' ('test_excel iniciado, rowCount={0}' -f $rowCount)

    # ---- 1. DISPONIBILIDAD ------------------------------------------------
    Write-Build Cyan '  -- [1/5] Disponibilidad COM'
    $avail = Test-ComAvailable -ProgId 'Excel.Application' -TimeoutSec 20
    Chk 'Excel.Application COM disponible' $avail
    if (-not $avail) {
        Write-Build Yellow '  Excel no disponible. Abortando test.'
        Write-RunResult -Context $ctx -Success $false -ErrorMsg 'Excel COM no disponible'
        return
    }

    try {

        # ---- 2. INSTANCIAR Y CREAR WORKBOOK --------------------------------
        Write-Build Cyan '  -- [2/5] Instanciar Excel y crear workbook'

        $xl = New-ExcelApp -Context $ctx -TimeoutSec 30
        Chk 'New-ExcelApp' ($null -ne $xl)
        if ($null -eq $xl) { throw 'Excel app null' }

        try { $wb = New-ExcelWorkbook -Context $ctx -ExcelApp $xl } catch { $wb = $null }
        Chk 'New-ExcelWorkbook' ($null -ne $wb) $(if ($null -eq $wb) { 'Workbooks.Add() null' } else { '' })
        if ($null -eq $wb) { throw 'Workbook null' }

        $ws = $null
        try { $ws = Get-ExcelSheet -Workbook $wb -Index 1 } catch {}
        Chk 'Get-ExcelSheet (hoja 1)' ($null -ne $ws)
        if ($null -eq $ws) { throw 'Sheet null' }

        # Renombrar hoja
        try { $ws.Name = 'Datos' } catch {}

        # ---- 3. ESCRITURA DATOS (celda-a-celda, sin array 2D) ---------------
        Write-Build Cyan ('  -- [3/5] Escribir {0} filas de datos' -f $rowCount)

        # Generar datos y exportar a CSV nativo (sin COM)
        $csvPath = Join-Path $ctx.Paths.Reports ('test_excel_data_{0}.csv' -f $stamp)
        $psoList = [System.Collections.Generic.List[psobject]]::new()
        for ($i = 1; $i -le $rowCount; $i++) {
            $psoList.Add([pscustomobject]@{
                ID        = $i
                Producto  = ('Prod_{0:D4}' -f $i)
                Categoria = @('Alpha','Beta','Gamma','Delta')[$i % 4]
                Precio    = [math]::Round(10 + ([math]::Sin($i) + 1) * 500, 2)
                Stock     = ($i * 7) % 200
                Activo    = $(if ($i % 3 -ne 0) { 'SI' } else { 'NO' })
            })
        }

        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $psoList | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        $sw.Stop()
        $csvOk = Test-Path $csvPath
        Chk ('Export-Csv {0} filas' -f $rowCount) $csvOk ('{0}ms' -f $sw.ElapsedMilliseconds)

        if ($csvOk) {
            # Importar CSV y escribir en Excel via Cells()
            $imported = @(Import-Csv -Path $csvPath -Encoding UTF8)
            $headers  = @('ID','Producto','Categoria','Precio','Stock','Activo')

            # Fila de cabecera
            for ($c = 0; $c -lt $headers.Count; $c++) {
                $cell = $null
                try {
                    $cell = $ws.Cells(1, $c + 1)
                    $cell.Value2 = $headers[$c]
                } finally { if ($null -ne $cell) { Invoke-ReleaseComObject $cell } }
            }

            # Filas de datos
            $sw2 = [System.Diagnostics.Stopwatch]::StartNew()
            for ($r = 0; $r -lt $imported.Count; $r++) {
                $row  = $imported[$r]
                $vals = @([string]$row.ID, [string]$row.Producto, [string]$row.Categoria,
                          [string]$row.Precio, [string]$row.Stock, [string]$row.Activo)
                for ($c = 0; $c -lt $vals.Count; $c++) {
                    $cell = $null
                    try {
                        $cell = $ws.Cells($r + 2, $c + 1)
                        $cell.Value2 = $vals[$c]
                    } finally { if ($null -ne $cell) { Invoke-ReleaseComObject $cell } }
                }
            }
            $sw2.Stop()
            Chk ('Escritura Excel {0} filas via Cells()' -f $rowCount) $true `
                ('{0}ms ({1}ms/fila)' -f $sw2.ElapsedMilliseconds,
                 [math]::Round($sw2.ElapsedMilliseconds / [math]::Max($rowCount,1), 1))

            # AutoFit
            Invoke-ExcelAutoFit -Sheet $ws

            # Verificar UsedRange.Rows directo
            $usedRng  = $null
            $usedRows = $null
            $rbRows   = 0
            try {
                $usedRng  = $ws.UsedRange
                $usedRows = $usedRng.Rows
                $rbRows   = [int]$usedRows.Count
            } finally {
                if ($null -ne $usedRows) { Invoke-ReleaseComObject $usedRows }
                if ($null -ne $usedRng)  { Invoke-ReleaseComObject $usedRng  }
            }
            $expRows = $rowCount + 1
            Chk 'UsedRange.Rows correcto' ($rbRows -eq $expRows) `
                ('Esperado={0} Real={1}' -f $expRows, $rbRows)

        } else {
            Chk 'Escritura Excel via Cells()' $false 'Omitido: CSV no generado'
            Chk 'UsedRange.Rows correcto'     $false 'Omitido: CSV no generado'
        }

        # ---- 4. GUARDAR Y LEER DE VUELTA -----------------------------------
        Write-Build Cyan '  -- [4/5] Guardar XLSX y leer de vuelta'
        $xlsxPath = Join-Path $ctx.Paths.Reports ('test_excel_{0}.xlsx' -f $stamp)

        $saveOk = $false
        try {
            # Force Excel to flush pending cell writes before SaveAs.
            # In Visible=false mode, Cells().Value2 assignments are buffered.
            # ScreenUpdating=true + Calculate forces the engine to process
            # all pending operations and update UsedRange before we save.
            try {
                $xl.ScreenUpdating = $true
                $xl.Calculate()
                $xl.ScreenUpdating = $false
            } catch {}
            Save-ExcelWorkbook -Context $ctx -Workbook $wb -Path $xlsxPath -Format 'xlsx'
            $saveOk = Test-Path $xlsxPath
        } catch { Write-BuildLog $ctx 'ERROR' ('Save XLSX: {0}' -f $_) }
        Chk 'Guardado como .xlsx' $saveOk $(if ($saveOk) { $xlsxPath } else { 'SaveAs fallo' })

        if ($saveOk) {
            $wb2 = $null
            try {
                $wb2 = Open-ExcelWorkbook -Context $ctx -ExcelApp $xl -Path $xlsxPath -ReadOnly $true
                Chk 'Open-ExcelWorkbook (read-back)' ($null -ne $wb2)

                if ($null -ne $wb2) {
                    $ws2 = $null
                    try {
                        $ws2 = Get-ExcelSheet -Workbook $wb2 -Index 1
                        if ($null -ne $ws2) {
                            # Read UsedRange.Value2 as 2D array - most reliable row count.
                            # UsedRange.Rows.Count can return stale value if Excel cached
                            # the range before flush. Value2 forces a full range read.
                            $expRows   = $rowCount + 1
                            $rbRows2   = 0
                            $usedRng2  = $null
                            $usedRowsO = $null
                            $val2D     = $null
                            try {
                                $usedRng2  = $ws2.UsedRange
                                $usedRowsO = $usedRng2.Rows
                                # Primary: GetLength(0) on Value2 2D array
                                try {
                                    $val2D   = $usedRng2.Value2
                                    if ($null -ne $val2D -and $val2D.GetType().IsArray) {
                                        $rbRows2 = $val2D.GetLength(0)
                                    }
                                } catch {}
                                # Fallback: Rows.Count COM object
                                if ($rbRows2 -eq 0) {
                                    $rbRows2 = [int]$usedRowsO.Count
                                }
                            } finally {
                                if ($null -ne $usedRowsO) { Invoke-ReleaseComObject $usedRowsO }
                                if ($null -ne $usedRng2)  { Invoke-ReleaseComObject $usedRng2  }
                            }
                            Chk 'Read-back: filas correctas' ($rbRows2 -eq $expRows) `
                                ('Esperado={0} ReadBack={1}' -f $expRows, $rbRows2)

                            # Leer celda A2 via Cells() - verificar primer dato escrito
                            $cellA2 = $null
                            $valA2  = ''
                            try {
                                $cellA2 = $ws2.Cells(2, 1)
                                $valA2  = [string]$cellA2.Value2
                            } finally {
                                if ($null -ne $cellA2) { Invoke-ReleaseComObject $cellA2 }
                            }
                            Chk 'Read-back: celda A2 = 1' ($valA2 -eq '1') "A2='$valA2'"
                        } else {
                            Chk 'Read-back: filas correctas' $false 'Sheet null'
                            Chk 'Read-back: celda A2 = 1'   $false 'Sheet null'
                        }
                    } catch {
                        Chk 'Read-back: filas correctas' $false ('{0}' -f $_)
                        Chk 'Read-back: celda A2 = 1'   $false 'Omitido por excepcion'
                    } finally {
                        if ($null -ne $ws2) { Invoke-ReleaseComObject $ws2; $ws2 = $null }
                    }
                }
            } catch {
                Chk 'Open-ExcelWorkbook (read-back)' $false ('{0}' -f $_)
            } finally {
                if ($null -ne $wb2) {
                    try { $wb2.Close($false) } catch {}
                    try { Invoke-ReleaseComObject $wb2 } catch {}
                    $wb2 = $null
                }
            }
        } else {
            Chk 'Open-ExcelWorkbook (read-back)' $false 'Omitido: XLSX no guardado'
            Chk 'Read-back: filas correctas'     $false 'Omitido: XLSX no guardado'
            Chk 'Read-back: celda A2 = 1'        $false 'Omitido: XLSX no guardado'
        }

        # ---- 5. VALIDAR CSV ------------------------------------------------
        Write-Build Cyan '  -- [5/5] Validar CSV'
        if ($csvOk) {
            $csvRows = @(Import-Csv -Path $csvPath -Encoding UTF8 -ErrorAction SilentlyContinue)
            Chk 'CSV filas correctas' ($csvRows.Count -eq $rowCount) `
                ('Esperado={0} CSV={1}' -f $rowCount, $csvRows.Count)
            Chk 'CSV columna ID presente' ($csvRows.Count -gt 0 -and $null -ne $csvRows[0].ID)
            Chk 'CSV columna Precio presente' ($csvRows.Count -gt 0 -and $null -ne $csvRows[0].Precio)
        } else {
            Chk 'CSV filas correctas'       $false 'Omitido: CSV no generado'
            Chk 'CSV columna ID presente'   $false 'Omitido: CSV no generado'
            Chk 'CSV columna Precio presente' $false 'Omitido: CSV no generado'
        }

    } catch {
        Write-BuildLog $ctx 'ERROR' ('test_excel fallo: {0}' -f $_) -Detail $_.ScriptStackTrace
        $script:failed++
    } finally {
        if ($null -ne $ws2) { try { Invoke-ReleaseComObject $ws2 } catch {} }
        if ($null -ne $wb2) { try { $wb2.Close($false) } catch {}; try { Invoke-ReleaseComObject $wb2 } catch {} }
        if ($null -ne $ws)  { try { Invoke-ReleaseComObject $ws  } catch {} }
        if ($null -ne $wb)  { try { $wb.Close($false)  } catch {}; try { Invoke-ReleaseComObject $wb  } catch {} }
        if ($null -ne $xl)  { try { $xl.Quit()         } catch {}; try { Invoke-ReleaseComObject $xl  } catch {} }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
    }

    # ---- RESUMEN -----------------------------------------------------------
    Write-Build Cyan ("`n  TEST EXCEL: Total={0} Pasaron={1} Fallaron={2}" -f
                      $results.Count, $script:passed, $script:failed)

    $jsonOut = Join-Path $ctx.Paths.Reports ('test_excel_{0}.json' -f $stamp)
    try {
        @{ runId=$ctx.RunId; passed=$script:passed; failed=$script:failed
           xlsxFile=$xlsxPath; csvFile=$csvPath; checks=$results.ToArray() } |
            ConvertTo-Json -Depth 3 |
            ForEach-Object { [System.IO.File]::WriteAllText($jsonOut, $_, [System.Text.Encoding]::UTF8) }
        Write-Build Cyan ('  Reporte: {0}' -f $jsonOut)
    } catch {}

    if ($openResult -and $null -ne $xlsxPath -and (Test-Path $xlsxPath)) {
        try { Start-Process $xlsxPath } catch {}
    }

    if ($script:failed -gt 0) {
        Write-RunResult -Context $ctx -Success $false `
            -ErrorMsg ('{0} verificacion(es) fallida(s)' -f $script:failed)
        throw ('test_excel: {0} check(s) failed' -f $script:failed)
    }
    Write-RunResult -Context $ctx -Success $true
}
