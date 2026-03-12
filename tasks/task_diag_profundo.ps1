#Requires -Version 5.1
# =============================================================================
# tasks/task_diag_profundo.ps1
# @Description : Suite de diagnostico profundo. Ejecuta en orden:
#                diag_motor -> diag_errores -> diag_com_stress ->
#                diag_concurrencia -> diag_checkpoint
#                Genera un reporte CSV consolidado de resultados.
# @Category    : Diagnostico
# @Version     : 1.0.0
# =============================================================================
# Solo ASCII. PS 5.1.

# Synopsis: Suite completa de pruebas profundas: motor, errores, stress COM, concurrencia, checkpoint
task diag_profundo diag_motor, diag_errores, diag_com_stress, diag_concurrencia, diag_checkpoint, {
    $ctx = New-TaskContext `
        -TaskName 'diag_profundo' `
        -Config   $Script:Config `
        -Root     (Split-Path $BuildRoot -Parent)

    Write-Build Cyan  ""
    Write-Build Cyan  "  ================================================"
    Write-Build Cyan  "  SUITE DIAGNOSTICO PROFUNDO - RESUMEN"
    Write-Build Cyan  "  ================================================"
    Write-Build Cyan  "  RunId : $($ctx.RunId)"
    Write-Build Cyan  "  Fecha : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Write-Build Cyan  ""

    # Consultar el registro para ver resultados de este run
    $regFile = Join-Path $ctx.Paths.Logs 'registry.jsonl'
    $resumen = @()

    if (Test-Path $regFile) {
        # Write-RunResult escribe level='OK' o level='ERROR' (ver lib/Logger.ps1)
        # El campo de tarea es 'task', no 'taskName'
        $entradas = @(Get-Content $regFile -Encoding ASCII |
            Where-Object { $_ -ne '' } |
            ForEach-Object { try { $_ | ConvertFrom-Json } catch { $null } } |
            Where-Object {
                if ($null -eq $_) { return $false }
                $lp = $_.PSObject.Properties['level']
                $null -ne $lp -and ($lp.Value -eq 'OK' -or $lp.Value -eq 'ERROR')
            })

        # Las 5 tareas del suite
        $tareasEsperadas = @('diag_motor','diag_errores','diag_com_stress','diag_concurrencia','diag_checkpoint')
        foreach ($tarea in $tareasEsperadas) {
            $entrada = @($entradas | Where-Object {
                $tp = $_.PSObject.Properties['task']
                $null -ne $tp -and $tp.Value -eq $tarea
            }) | Select-Object -Last 1

            $estado = if ($null -ne $entrada) {
                $lp = $entrada.PSObject.Properties['level']
                if ($null -ne $lp -and $lp.Value -eq 'OK') { 'OK' } else { 'FAIL' }
            } else {
                'N/A'
            }
            $resumen += [PSCustomObject]@{
                Tarea  = $tarea
                Estado = $estado
                RunId  = $ctx.RunId
                Fecha  = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            }
            $color = switch ($estado) { 'OK' { 'Green' } 'FAIL' { 'Red' } default { 'Yellow' } }
            Write-Build $color "  $($tarea.PadRight(22)) : $estado"
        }
    } else {
        Write-Build Yellow "  Registro no encontrado en $regFile"
    }

    # Guardar CSV de resumen
    if ($resumen.Count -gt 0) {
        $outF = Join-Path $ctx.Paths.Reports "diag_profundo_$($ctx.RunId).csv"
        $resumen | Export-Csv -Path $outF -NoTypeInformation -Encoding ASCII
        Write-Build Cyan  ""
        Write-Build Cyan  "  Reporte guardado: $outF"
    }

    $totalFail = @($resumen | Where-Object { $_.Estado -eq 'FAIL' }).Count
    Write-Build Cyan  ""
    Write-Build Cyan  "  ================================================"

    Write-BuildLog $ctx 'INFO' "diag_profundo completado. Fallos=$totalFail"
    Write-RunResult -Context $ctx -Success ($totalFail -eq 0)

    if ($totalFail -gt 0) {
        Write-Build Red   "  RESULTADO: $totalFail tarea(s) con fallos"
        throw "diag_profundo: $totalFail tarea(s) fallaron"
    } else {
        Write-Build Green "  RESULTADO: OK - todas las pruebas profundas pasaron"
    }
}
