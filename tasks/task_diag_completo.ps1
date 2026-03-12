#Requires -Version 5.1
# =============================================================================
# tasks/task_diag_completo.ps1
# @Description : Suite completa de diagnostico: ejecuta todas las tareas
#                diag_* en orden y presenta resumen consolidado.
#                Excluye SAP. Excluye diag_excel_rw (requiere -Extra opcional).
# @Category    : Utilidad
# @Version     : 1.1.0
# =============================================================================
# Solo ASCII. PS 5.1.

# Synopsis: Ejecuta todos los diagnosticos disponibles y muestra resumen final
task diag_completo diag_entorno, diag_rutas, diag_com, diag_csv, diag_log, diag_excel, diag_excel_rw, diag_word, diag_pdf, diag_multi_com, {
    $ctx = New-TaskContext `
        -TaskName 'diag_completo' `
        -Config   $Script:Config `
        -Root     (Split-Path $BuildRoot -Parent)

    Write-Build Cyan  ""
    Write-Build Cyan  "  ============================================="
    Write-Build Green "  SUITE DE DIAGNOSTICO COMPLETADA SIN ERRORES"
    Write-Build Cyan  "  ============================================="
    Write-Build Cyan  ""
    Write-Build Green "  diag_entorno    : OK - PS 5.1, OS, carpetas, Invoke-Build"
    Write-Build Green "  diag_rutas      : OK - permisos R/W/D en todas las carpetas"
    Write-Build Green "  diag_com        : OK - servidores COM, Jobs, FSO"
    Write-Build Green "  diag_csv        : OK - lectura/escritura/encoding CSV"
    Write-Build Green "  diag_log        : OK - logging JSONL, formato, consultas"
    Write-Build Green "  diag_excel      : OK - ciclo completo Excel COM"
    Write-Build Green "  diag_excel_rw   : OK - lectura/escritura avanzada Excel"
    Write-Build Green "  diag_word       : OK - ciclo completo Word COM + PDF via Word"
    Write-Build Green "  diag_pdf        : OK - exportacion PDF Word y Excel"
    Write-Build Green "  diag_multi_com  : OK - Excel + Word secuencial sin conflictos"
    Write-Build Cyan  ""
    Write-Build Cyan  "  El entorno esta listo para tareas de produccion."
    Write-Build Cyan  ""

    Write-BuildLog $ctx 'INFO' 'Suite diag_completo v1.1 finalizada correctamente'
    Write-RunResult -Context $ctx -Success $true
}
