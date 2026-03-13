#Requires -Version 5.1
# =============================================================================
# tasks/task_test_csv.ps1
# @Description : Prueba funcional CSV - escritura, lectura, transformacion, validacion
# @Category    : Test
# @Version     : 1.0.0
# @Author      : AutoBuild QA
# @Environment : PS 5.1 Desktop (sin dependencia COM)
# =============================================================================
# Synopsis: Prueba funcional CSV nativa - sin COM, sin Excel, puro PowerShell
# Params: {"RowCount":"500","Delimiter":","}

task test_csv {

    $ctx = New-TaskContext `
        -TaskName 'test_csv' `
        -Config   $Script:EngineConfig `
        -Root     $Script:EngineRoot `
        -Params   $Script:TaskParams

    $rowCount  = 500
    try { $rv = $ctx.Params['RowCount']; if ($rv -gt '') { $rowCount = [int]$rv } } catch {}
    if ($rowCount -lt 1) { $rowCount = 500 }

    $delimiter = ','
    try { $dv = $ctx.Params['Delimiter']; if ($dv -gt '') { $delimiter = $dv } } catch {}

    $stamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
    $script:passed = 0
    $script:failed = 0
    $results = [System.Collections.Generic.List[hashtable]]::new()

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

    Write-Build Cyan ("`n  TEST CSV -- {0} filas, delimitador='{1}'`n" -f $rowCount, $delimiter)
    Write-BuildLog $ctx 'INFO' ('test_csv iniciado, rowCount={0}' -f $rowCount)

    $csvPath1   = Join-Path $ctx.Paths.Reports ('test_csv_datos_{0}.csv' -f $stamp)
    $csvPath2   = Join-Path $ctx.Paths.Reports ('test_csv_filtrado_{0}.csv' -f $stamp)
    $csvPath3   = Join-Path $ctx.Paths.Reports ('test_csv_transformado_{0}.csv' -f $stamp)

    # ---- 1. GENERAR Y ESCRIBIR CSV -----------------------------------------
    Write-Build Cyan ('  -- [1/5] Generar y escribir {0} filas' -f $rowCount)

    $psoList = [System.Collections.Generic.List[psobject]]::new()
    for ($i = 1; $i -le $rowCount; $i++) {
        $psoList.Add([pscustomobject]@{
            ID         = $i
            Codigo     = ('COD-{0:D5}' -f $i)
            Descripcion = ('Articulo numero {0} de prueba funcional' -f $i)
            Categoria  = @('Electronica','Ropa','Alimentos','Herramientas','Libros')[$i % 5]
            Precio     = [math]::Round(1.5 + ([math]::Abs([math]::Sin($i))) * 999, 2)
            Cantidad   = ($i * 13) % 100 + 1
            Activo     = $(if ($i % 7 -ne 0) { 'true' } else { 'false' })
            FechaAlta  = (Get-Date).AddDays(-$i).ToString('yyyy-MM-dd')
        })
    }

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        if ($delimiter -eq ',') {
            $psoList | Export-Csv -Path $csvPath1 -NoTypeInformation -Encoding UTF8
        } else {
            $psoList | Export-Csv -Path $csvPath1 -NoTypeInformation -Encoding UTF8 -Delimiter $delimiter
        }
        $sw.Stop()
        Chk ('Export-Csv {0} filas' -f $rowCount) (Test-Path $csvPath1) `
            ('{0}ms, {1}KB' -f $sw.ElapsedMilliseconds,
             [math]::Round((Get-Item $csvPath1).Length / 1KB, 1))
    } catch {
        $sw.Stop()
        Chk ('Export-Csv {0} filas' -f $rowCount) $false ('{0}' -f $_)
    }

    # ---- 2. LEER Y VALIDAR ESTRUCTURA --------------------------------------
    Write-Build Cyan '  -- [2/5] Leer y validar estructura'

    $imported = $null
    if (Test-Path $csvPath1) {
        try {
            $sw2 = [System.Diagnostics.Stopwatch]::StartNew()
            $imported = @(Import-Csv -Path $csvPath1 -Encoding UTF8)
            $sw2.Stop()
            Chk 'Import-Csv exitoso' $true ('{0}ms' -f $sw2.ElapsedMilliseconds)
            Chk 'Filas leidas correctas' ($imported.Count -eq $rowCount) `
                ('Esperado={0} Leido={1}' -f $rowCount, $imported.Count)
        } catch {
            Chk 'Import-Csv exitoso' $false ('{0}' -f $_)
        }
    } else {
        Chk 'Import-Csv exitoso' $false 'Archivo CSV no existe'
    }

    if ($null -ne $imported -and $imported.Count -gt 0) {
        $firstRow = $imported[0]
        $lastRow  = $imported[$imported.Count - 1]
        $cols     = @('ID','Codigo','Descripcion','Categoria','Precio','Cantidad','Activo','FechaAlta')
        $allCols  = $true
        foreach ($col in $cols) {
            if ($null -eq $firstRow.$col) { $allCols = $false; break }
        }
        Chk 'Todas las columnas presentes (8)' $allCols ($cols -join ', ')
        Chk 'Primera fila ID=1' ($firstRow.ID -eq '1') "ID='$($firstRow.ID)'"
        Chk ('Ultima fila ID={0}' -f $rowCount) ($lastRow.ID -eq [string]$rowCount) `
            "ID='$($lastRow.ID)'"
        Chk 'Campo Precio no vacio en fila 1' ($firstRow.Precio -ne '') "Precio='$($firstRow.Precio)'"

        # Verificar integridad: todos los IDs son numeros del 1 al rowCount
        $idList = [System.Collections.Generic.List[int]]::new()
        foreach ($row in $imported) {
            $n = 0
            if ([int]::TryParse($row.ID, [ref]$n)) { $idList.Add($n) }
        }
        $idArr  = $idList.ToArray()
        [System.Array]::Sort($idArr)
        $idOk   = ($idArr.Count -eq $rowCount -and $idArr[0] -eq 1 -and $idArr[$idArr.Count-1] -eq $rowCount)
        Chk 'IDs 1..N sin duplicados ni huecos' $idOk `
            ('Count={0} Min={1} Max={2}' -f $idArr.Count,
             $(if($idArr.Count -gt 0){$idArr[0]}else{'N/A'}),
             $(if($idArr.Count -gt 0){$idArr[$idArr.Count-1]}else{'N/A'}))
    } else {
        foreach ($n in @('Todas las columnas presentes (8)','Primera fila ID=1',
                         'Ultima fila ID correcto','Campo Precio no vacio','IDs sin duplicados')) {
            Chk $n $false 'Omitido: datos no disponibles'
        }
    }

    # ---- 3. FILTRAR Y RE-EXPORTAR ------------------------------------------
    Write-Build Cyan '  -- [3/5] Filtrar (Activo=true) y re-exportar'

    if ($null -ne $imported) {
        $filteredList = [System.Collections.Generic.List[psobject]]::new()
        foreach ($row in $imported) {
            if ([string]$row.Activo -eq 'true') { $filteredList.Add($row) }
        }
        $expectedActive = 0
        for ($i = 1; $i -le $rowCount; $i++) { if ($i % 7 -ne 0) { $expectedActive++ } }

        Chk 'Filtrado Activo=true cuenta correcta' ($filteredList.Count -eq $expectedActive) `
            ('Esperado={0} Filtrado={1}' -f $expectedActive, $filteredList.Count)

        try {
            $filteredList | Export-Csv -Path $csvPath2 -NoTypeInformation -Encoding UTF8
            Chk 'CSV filtrado exportado' (Test-Path $csvPath2) $csvPath2
        } catch {
            Chk 'CSV filtrado exportado' $false ('{0}' -f $_)
        }
    } else {
        Chk 'Filtrado Activo=true' $false 'Omitido: datos no disponibles'
        Chk 'CSV filtrado exportado' $false 'Omitido'
    }

    # ---- 4. TRANSFORMAR (calculo de campo nuevo) ----------------------------
    Write-Build Cyan '  -- [4/5] Transformar (calcular Subtotal = Precio * Cantidad)'

    if ($null -ne $imported) {
        $transformList = [System.Collections.Generic.List[psobject]]::new()
        foreach ($row in $imported) {
            $precio   = 0.0
            $cantidad = 0
            [double]::TryParse($row.Precio,   [System.Globalization.NumberStyles]::Any,
                [System.Globalization.CultureInfo]::InvariantCulture, [ref]$precio)   | Out-Null
            [int]::TryParse($row.Cantidad, [ref]$cantidad) | Out-Null
            $transformList.Add([pscustomobject]@{
                ID        = $row.ID
                Codigo    = $row.Codigo
                Categoria = $row.Categoria
                Precio    = $precio
                Cantidad  = $cantidad
                Subtotal  = [math]::Round($precio * $cantidad, 2)
                Activo    = $row.Activo
            })
        }

        try {
            $transformList | Export-Csv -Path $csvPath3 -NoTypeInformation -Encoding UTF8
            Chk 'CSV transformado exportado' (Test-Path $csvPath3) $csvPath3
        } catch {
            Chk 'CSV transformado exportado' $false ('{0}' -f $_)
        }

        # Verificar que Subtotal = Precio * Cantidad en primera fila
        if ($transformList.Count -gt 0) {
            $f    = $transformList[0]
            $calc = [math]::Round($f.Precio * $f.Cantidad, 2)
            Chk 'Subtotal calculado correctamente (fila 1)' ($f.Subtotal -eq $calc) `
                ('Precio={0} Cantidad={1} Subtotal={2} Esperado={3}' -f $f.Precio, $f.Cantidad, $f.Subtotal, $calc)
        }

        # Total acumulado
        $totalSubtotal = 0.0
        foreach ($row in $transformList) { $totalSubtotal += $row.Subtotal }
        Chk 'Suma Subtotal > 0' ($totalSubtotal -gt 0) `
            ('Total={0:N2}' -f $totalSubtotal)
    } else {
        foreach ($n in @('CSV transformado exportado','Subtotal calculado','Suma Subtotal')) {
            Chk $n $false 'Omitido'
        }
    }

    # ---- 5. LECTURA RAPIDA CON Get-Content ---------------------------------
    Write-Build Cyan '  -- [5/5] Lectura rapida (Get-Content, sin Import-Csv)'

    if (Test-Path $csvPath1) {
        try {
            $rawLines = @(Get-Content -Path $csvPath1 -Encoding UTF8)
            # rawLines[0] = header, rawLines[1..N] = datos
            $dataLines = $rawLines.Count - 1
            Chk 'Get-Content: lineas de datos correctas' ($dataLines -eq $rowCount) `
                ('Esperado={0} Leido={1}' -f $rowCount, $dataLines)
            Chk 'Get-Content: cabecera contiene ID' ($rawLines[0] -like '*ID*') `
                "Header='$($rawLines[0].Substring(0, [Math]::Min(60,$rawLines[0].Length)))'"
        } catch {
            Chk 'Get-Content rapido' $false ('{0}' -f $_)
        }
    } else {
        Chk 'Get-Content rapido' $false 'Archivo no existe'
    }

    # ---- RESUMEN -----------------------------------------------------------
    Write-Build Cyan ("`n  TEST CSV: Total={0} Pasaron={1} Fallaron={2}" -f
                      $results.Count, $script:passed, $script:failed)
    Write-Build Cyan ('  Archivos: {0}' -f $ctx.Paths.Reports)

    $jsonOut = Join-Path $ctx.Paths.Reports ('test_csv_{0}.json' -f $stamp)
    try {
        @{ runId=$ctx.RunId; passed=$script:passed; failed=$script:failed
           files=@($csvPath1,$csvPath2,$csvPath3); checks=$results.ToArray() } |
            ConvertTo-Json -Depth 3 |
            ForEach-Object { [System.IO.File]::WriteAllText($jsonOut, $_, [System.Text.Encoding]::UTF8) }
        Write-Build Cyan ('  Reporte: {0}' -f $jsonOut)
    } catch {}

    if ($script:failed -gt 0) {
        Write-RunResult -Context $ctx -Success $false `
            -ErrorMsg ('{0} verificacion(es) fallida(s)' -f $script:failed)
        throw ('test_csv: {0} check(s) failed' -f $script:failed)
    }
    Write-RunResult -Context $ctx -Success $true
}
