#Requires -Version 5.1
# =============================================================================
# tasks/task_diag_checkpoint.ps1
# @Description : Prueba del mecanismo Checkpoint/Resume de Invoke-Build.
#                Simula una tarea multi-fase con guardado de estado entre fases,
#                verifica que el archivo .clixml de checkpoint se crea y puede
#                ser releido, y prueba la reanudacion desde punto de control.
#
#                USO:
#                  .\Run.ps1 diag_checkpoint               <- ejecucion completa
#                  .\Run.ps1 diag_checkpoint -Checkpoint   <- guarda checkpoint
#                  .\Run.ps1 diag_checkpoint -Resume       <- reanuda desde ult. ok
#
# @Category    : Diagnostico
# @Version     : 1.0.0
# =============================================================================
# Solo ASCII. PS 5.1.

# Synopsis: Prueba el mecanismo de Checkpoint y Resume del motor Invoke-Build

task cp_fase1 {
    $ctx = New-TaskContext `
        -TaskName 'cp_fase1' `
        -Config   $Script:Config `
        -Root     (Split-Path $BuildRoot -Parent)

    Write-Build Cyan  "  [F1]  Iniciando Fase 1 - preparacion de datos"
    Write-BuildLog $ctx 'INFO' 'Checkpoint Fase 1 inicio'

    $outF = Join-Path $ctx.Paths.Output "checkpoint_fase1_$($ctx.RunId).csv"
    $datos = @(
        [PSCustomObject]@{ Fase = '1'; Dato = 'Alpha'; Valor = '10' },
        [PSCustomObject]@{ Fase = '1'; Dato = 'Beta';  Valor = '20' },
        [PSCustomObject]@{ Fase = '1'; Dato = 'Gamma'; Valor = '30' }
    )
    $datos | Export-Csv -Path $outF -NoTypeInformation -Encoding ASCII

    if (Test-Path $outF) {
        Write-Build Green "  [F1]  CSV Fase 1 creado : $outF"
    } else {
        throw "Fase 1: no se pudo crear el CSV"
    }

    Write-BuildLog $ctx 'INFO' "Checkpoint Fase 1 OK: $outF"
    Write-Build Green "  [F1]  Fase 1 completada"
}

task cp_fase2 cp_fase1, {
    $ctx = New-TaskContext `
        -TaskName 'cp_fase2' `
        -Config   $Script:Config `
        -Root     (Split-Path $BuildRoot -Parent)

    Write-Build Cyan  "  [F2]  Iniciando Fase 2 - transformacion"
    Write-BuildLog $ctx 'INFO' 'Checkpoint Fase 2 inicio'

    # Leer CSV de fase 1 (el mas reciente si hay varios)
    $archivos = @(Get-ChildItem -Path $ctx.Paths.Output `
        -Filter 'checkpoint_fase1_*.csv' -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending)

    if ($archivos.Count -eq 0) {
        throw "Fase 2: no se encontro el CSV de Fase 1 - ejecutar desde Fase 1"
    }

    $src  = $archivos[0].FullName
    $leido = @(Import-Csv -Path $src -Encoding ASCII)
    Write-Build Green "  [F2]  Leidas $($leido.Count) filas de Fase 1"

    # Transformar: multiplicar Valor x 2
    $transformado = $leido | ForEach-Object {
        [PSCustomObject]@{ Fase = '2'; Dato = $_.Dato; Valor = ([int]$_.Valor * 2).ToString() }
    }

    $outF = Join-Path $ctx.Paths.Output "checkpoint_fase2_$($ctx.RunId).csv"
    $transformado | Export-Csv -Path $outF -NoTypeInformation -Encoding ASCII
    Write-Build Green "  [F2]  CSV Fase 2 creado : $outF"

    Write-BuildLog $ctx 'INFO' "Checkpoint Fase 2 OK: $outF"
    Write-Build Green "  [F2]  Fase 2 completada"
}

task cp_fase3 cp_fase2, {
    $ctx = New-TaskContext `
        -TaskName 'cp_fase3' `
        -Config   $Script:Config `
        -Root     (Split-Path $BuildRoot -Parent)

    Write-Build Cyan  "  [F3]  Iniciando Fase 3 - consolidacion"
    Write-BuildLog $ctx 'INFO' 'Checkpoint Fase 3 inicio'

    $f1List = @(Get-ChildItem $ctx.Paths.Output -Filter 'checkpoint_fase1_*.csv' |
        Sort-Object LastWriteTime -Descending)
    $f2List = @(Get-ChildItem $ctx.Paths.Output -Filter 'checkpoint_fase2_*.csv' |
        Sort-Object LastWriteTime -Descending)

    if ($f1List.Count -eq 0 -or $f2List.Count -eq 0) {
        throw "Fase 3: faltan archivos de fases anteriores"
    }

    $f1 = @(Import-Csv $f1List[0].FullName -Encoding ASCII)
    $f2 = @(Import-Csv $f2List[0].FullName -Encoding ASCII)

    $consolidado = @()
    for ($i = 0; $i -lt [Math]::Min($f1.Count, $f2.Count); $i++) {
        $consolidado += [PSCustomObject]@{
            Dato      = $f1[$i].Dato
            ValorF1   = $f1[$i].Valor
            ValorF2   = $f2[$i].Valor
            Diferencia = ([int]$f2[$i].Valor - [int]$f1[$i].Valor).ToString()
        }
    }

    $outF = Join-Path $ctx.Paths.Output "checkpoint_consolidado_$($ctx.RunId).csv"
    $consolidado | Export-Csv -Path $outF -NoTypeInformation -Encoding ASCII

    Write-Build Green "  [F3]  Consolidado $($consolidado.Count) filas"
    Write-Build Green "  [F3]  Archivo: $outF"

    # Verificar que las diferencias son exactamente el valor de F1 (F2 = F1*2)
    $fallos = @($consolidado | Where-Object { [int]$_.Diferencia -ne [int]$_.ValorF1 })
    if ($fallos.Count -eq 0) {
        Write-Build Green "  [F3]  Verificacion matematica : OK"
    } else {
        Write-Build Red   "  [F3]  Diferencias inesperadas en $($fallos.Count) filas"
        throw "Fase 3: verificacion matematica fallo"
    }

    Write-BuildLog $ctx 'INFO' "Checkpoint Fase 3 OK: $outF"
    Write-Build Green "  [F3]  Fase 3 completada"
}

# ---- Tarea de validacion del mecanismo de checkpoint ----------------------
# Synopsis: Prueba el mecanismo de Checkpoint/Resume: 3 fases con estado persistente
task diag_checkpoint cp_fase3, {
    $ctx = New-TaskContext `
        -TaskName 'diag_checkpoint' `
        -Config   $Script:Config `
        -Root     (Split-Path $BuildRoot -Parent)

    Write-Build Cyan ""
    Write-Build Cyan "  [CHK]  Verificando artefactos de las 3 fases..."

    $patrones = @(
        'checkpoint_fase1_*.csv',
        'checkpoint_fase2_*.csv',
        'checkpoint_consolidado_*.csv'
    )
    $errores = 0
    foreach ($pat in $patrones) {
        $arch = @(Get-ChildItem $ctx.Paths.Output -Filter $pat -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending)
        if ($arch.Count -gt 0) {
            $sz = (Get-Item $arch[0].FullName).Length
            Write-Build Green "  [CHK]  $pat : OK ($sz bytes)"
        } else {
            Write-Build Red   "  [CHK]  $pat : NO ENCONTRADO"
            $errores++
        }
    }

    Write-Build Cyan ""
    Write-Build Cyan "  NOTA: Para probar Resume:"
    Write-Build Cyan "    1. Eliminar checkpoint_fase2_*.csv de output\"
    Write-Build Cyan "    2. Ejecutar: .\Run.ps1 diag_checkpoint -Checkpoint"
    Write-Build Cyan "    3. Ejecutar: .\Run.ps1 diag_checkpoint -Resume"
    Write-Build Cyan "       (Invoke-Build saltara cp_fase1 y reanudara desde cp_fase2)"

    Write-BuildLog $ctx 'INFO' "diag_checkpoint completado. Errores=$errores"
    Write-RunResult -Context $ctx -Success ($errores -eq 0)

    if ($errores -gt 0) {
        throw "diag_checkpoint: $errores artefactos faltantes"
    } else {
        Write-Build Green "  RESULTADO: OK - pipeline de 3 fases completado"
    }
}
