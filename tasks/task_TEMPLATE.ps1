#Requires -Version 5.1
# =============================================================================
# tasks/task_NOMBRE.ps1
# @Description : Descripcion breve de la tarea
# @Category    : SAP | Excel | CSV | Reporte | Utilidad
# @Version     : 1.0.0
# @Param       : Centro  string required "Centro SAP (ej: 1000)"
# @Param       : Almacen string optional "Almacen (default: todos)"
# =============================================================================
# REGLAS OBLIGATORIAS:
#   - Solo ASCII en este archivo
#   - Nombre de archivo: task_[a-zA-Z0-9_-]+.ps1
#   - La tarea se define con: task NOMBRE { ... }
#   - No definir funciones globales aqui (usar lib/)
#   - No hardcodear rutas (usar $ctx.Paths.*)
#   - COM: siempre usar New-ExcelApp/Invoke-ComWithTimeout + try/finally
#   - Retornar Write-RunResult al final exitoso

# Synopsis: Descripcion breve que aparece en .\Run.ps1 -List
task NOMBRE {
    # ---- Contexto ------------------------------------------------------
    $ctx = New-TaskContext `
        -TaskName 'NOMBRE' `
        -Config   $Script:Config `
        -Root     (Split-Path $BuildRoot -Parent) `
        -Params   @{ Centro = $Centro }

    Write-BuildLog $ctx 'INFO' 'Iniciando NOMBRE'

    # ---- Prerequisitos (validar antes de hacer nada) ------------------
    Test-TaskAsset -Context $ctx -Params @{ Centro = $Centro }

    # ---- Logica principal ---------------------------------------------
    # ... implementar aqui ...

    # ---- Si usa Excel/Word --------------------------------------------
    $xl = $null
    $wb = $null
    try {
        $xl = New-ExcelApp -Context $ctx
        if ($null -eq $xl) { throw 'Excel no disponible' }

        # ... logica Excel ...

        Write-RunResult -Context $ctx -Success $true
    } finally {
        # SIEMPRE liberar COM aunque haya excepcion
        Close-ExcelWorkbook -Workbook $wb -Save $false
        Close-ExcelApp      -ExcelApp $xl
        $xl = $null
        $wb = $null
    }
}
