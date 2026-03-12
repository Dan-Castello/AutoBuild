#Requires -Version 5.1
# =============================================================================
# tasks/task_diag_csv.ps1
# @Description : Diagnostico de lectura y escritura de CSV en PS 5.1.
# @Category    : Utilidad
# @Version     : 1.1.0
# =============================================================================
# Solo ASCII. PS 5.1.
# REGLA: toda variable que recibe resultado de cmdlet/pipeline usa @()
#        para garantizar que .Count funcione con StrictMode -Version Latest.

# Synopsis: Prueba ciclo completo CSV (escritura, lectura, codificacion, rutas)
task diag_csv {
    $ctx = New-TaskContext `
        -TaskName 'diag_csv' `
        -Config   $Script:Config `
        -Root     (Split-Path $BuildRoot -Parent)

    Write-BuildLog $ctx 'INFO' 'Iniciando diagnostico de CSV'

    $errores = 0

    # ---- 1. Verificar carpetas ---------------------------------------------
    foreach ($dir in @('input','output')) {
        $p = $ctx.Paths.$dir
        if (Test-Path $p) {
            Write-Build Green "  [DIR]  $dir : $p OK"
        } else {
            Write-Build Red   "  [DIR]  $dir : FALTA - $p"
            $errores++
        }
    }
    if ($errores -gt 0) {
        throw "Carpetas de trabajo no encontradas - ejecutar Setup.ps1 primero"
    }

    # ---- 2. Generar CSV de prueba ------------------------------------------
    $csvOut = Join-Path $ctx.Paths.Output "diag_csv_$($ctx.RunId).csv"

    $filas = @(
        [PSCustomObject]@{ ID = '1'; Nombre = 'AutoBuild'; Valor = '100'; Activo = 'True'  },
        [PSCustomObject]@{ ID = '2'; Nombre = 'Motor';     Valor = '200'; Activo = 'True'  },
        [PSCustomObject]@{ ID = '3'; Nombre = 'Prueba';    Valor = '300'; Activo = 'False' }
    )

    Write-Build Cyan  "  [WRITE] Exportando CSV de prueba..."
    try {
        $filas | Export-Csv -Path $csvOut -NoTypeInformation -Encoding ASCII
        Write-Build Green "  [WRITE] Export-Csv : OK -> $csvOut"
    } catch {
        Write-Build Red   "  [WRITE] Export-Csv : FAIL - $_"
        $errores++
    }

    # ---- 3. Verificar lineas escritas --------------------------------------
    if (Test-Path $csvOut) {
        $lines = @(Get-Content $csvOut -Encoding ASCII).Count
        Write-Build Green "  [WRITE] Lineas escritas : $lines (incluye encabezado)"
    } else {
        Write-Build Red   "  [WRITE] Archivo no creado"
        $errores++
    }

    # ---- 4. Leer el CSV ----------------------------------------------------
    Write-Build Cyan  "  [READ]  Leyendo CSV generado..."
    $leido = @()
    try {
        $leido = @(Import-Csv -Path $csvOut -Encoding ASCII)
        Write-Build Green "  [READ]  Import-Csv : OK ($($leido.Count) filas)"
    } catch {
        Write-Build Red   "  [READ]  Import-Csv : FAIL - $_"
        $errores++
    }

    # ---- 5. Integridad -----------------------------------------------------
    if ($leido.Count -eq 3) {
        $fila1 = $leido[0]
        if ($fila1.ID -eq '1' -and $fila1.Nombre -eq 'AutoBuild') {
            Write-Build Green "  [INT]  Integridad datos : OK"
        } else {
            Write-Build Red   "  [INT]  FAIL (ID=$($fila1.ID) Nombre=$($fila1.Nombre))"
            $errores++
        }

        $cols          = @($leido[0].PSObject.Properties.Name)
        $colsEsperadas = @('ID','Nombre','Valor','Activo')
        $colsFaltantes = @($colsEsperadas | Where-Object { $_ -notin $cols })
        if ($colsFaltantes.Count -eq 0) {
            Write-Build Green "  [INT]  Columnas : OK ($($cols -join ', '))"
        } else {
            Write-Build Red   "  [INT]  Columnas faltantes : $($colsFaltantes -join ', ')"
            $errores++
        }
    } elseif ($leido.Count -gt 0) {
        Write-Build Red   "  [INT]  Filas esperadas 3, leidas $($leido.Count)"
        $errores++
    } else {
        Write-Build Red   "  [INT]  Sin filas leidas"
        $errores++
    }

    # ---- 6. Encoding Default (ANSI regional) -------------------------------
    $csvAnsi = Join-Path $ctx.Paths.Output "diag_csv_ansi_$($ctx.RunId).csv"
    Write-Build Cyan  "  [ENC]  Probando encoding Default (ANSI regional)..."
    try {
        $filas | Export-Csv -Path $csvAnsi -NoTypeInformation -Encoding Default
        $leidoAnsi = @(Import-Csv -Path $csvAnsi -Encoding Default)
        if ($leidoAnsi.Count -eq 3) {
            Write-Build Green "  [ENC]  Encoding Default : OK ($($leidoAnsi.Count) filas)"
        } else {
            Write-Build Yellow "  [ENC]  Encoding Default : conteo inesperado ($($leidoAnsi.Count) filas)"
        }
    } catch {
        Write-Build Yellow "  [ENC]  Encoding Default : WARN - $_"
    }

    # ---- 7. CSV en input/ --------------------------------------------------
    $csvInputList = @(Get-ChildItem -Path $ctx.Paths.Input -Filter '*.csv' -ErrorAction SilentlyContinue)
    if ($csvInputList.Count -gt 0) {
        $csvInput = $csvInputList[0]
        Write-Build Cyan  "  [INPUT] CSV encontrado : $($csvInput.Name)"
        try {
            $ext = @(Import-Csv -Path $csvInput.FullName -Encoding Default)
            Write-Build Green "  [INPUT] Lectura : OK ($($ext.Count) filas)"
            if ($ext.Count -gt 0) {
                Write-Build Cyan  "  [INPUT] Columnas : $($ext[0].PSObject.Properties.Name -join ', ')"
            }
        } catch {
            Write-Build Yellow "  [INPUT] WARN al leer: $_"
        }
    } else {
        Write-Build Cyan  "  [INPUT] Sin archivos CSV en input\ (normal en primera ejecucion)"
    }

    # ---- Resumen -----------------------------------------------------------
    Write-Build Cyan ""
    if ($errores -eq 0) {
        Write-Build Green "  RESULTADO: OK - operaciones CSV funcionan correctamente"
    } else {
        Write-Build Red   "  RESULTADO: $errores error(es) en operaciones CSV"
    }

    Write-BuildLog $ctx 'INFO' "diag_csv completado. Errores=$errores"
    Write-RunResult -Context $ctx -Success ($errores -eq 0)

    if ($errores -gt 0) {
        throw "diag_csv detecto $errores error(es)"
    }
}
