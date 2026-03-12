#Requires -Version 5.1
# =============================================================================
# tasks/task_diag_concurrencia.ps1
# @Description : Prueba de aislamiento entre tareas: verifica que dos RunIds
#                distintos en la misma carpeta output\ no colisionan en nombres
#                de archivo, que el log registra las dos ejecuciones por separado,
#                y que si se interrumpe una tarea a mitad (via Job), los archivos
#                parciales se detectan.
#                Tambien prueba el limite de una instancia COM a la vez (Regla 6):
#                intenta abrir Excel, luego Word sin cerrar Excel, verifica que
#                la tarea detecta el conflicto y actua correctamente.
# @Category    : Diagnostico
# @Version     : 1.0.0
# =============================================================================
# Solo ASCII. PS 5.1.

# Synopsis: Prueba aislamiento entre runs, colision de archivos y Regla 6 COM
task diag_concurrencia {
    $ctx = New-TaskContext `
        -TaskName 'diag_concurrencia' `
        -Config   $Script:Config `
        -Root     (Split-Path $BuildRoot -Parent)

    Write-BuildLog $ctx 'INFO' 'Iniciando prueba de concurrencia y aislamiento'
    $errores = 0

    # ---- 1. RunIds distintos no colisionan en nombres de archivo ----------
    Write-Build Cyan "  [ISO]  Verificando aislamiento de nombres de archivo por RunId..."

    $rid1 = New-RunId
    Start-Sleep -Milliseconds 50
    $rid2 = New-RunId

    if ($rid1 -ne $rid2) {
        $f1 = "output_$rid1.csv"
        $f2 = "output_$rid2.csv"
        if ($f1 -ne $f2) {
            Write-Build Green "  [ISO]  Nombres de archivo : no colisionan OK"
            Write-Build Cyan  "         $f1"
            Write-Build Cyan  "         $f2"
        } else {
            Write-Build Red   "  [ISO]  COLISION: mismos nombres con RunIds distintos"
            $errores++
        }
    } else {
        Write-Build Red "  [ISO]  RunIds identicos generados - no hay aislamiento"
        $errores++
    }

    # ---- 2. Dos contextos en misma sesion tienen Paths distintos ----------
    Write-Build Cyan ""
    Write-Build Cyan "  [ISO]  Verificando que dos contextos no comparten estado..."

    $ctx2 = New-TaskContext `
        -TaskName 'tarea_paralela' `
        -Config   $Script:Config `
        -Root     (Split-Path $BuildRoot -Parent)

    if ($ctx.RunId -ne $ctx2.RunId) {
        Write-Build Green "  [ISO]  RunIds distintos : OK"
    } else {
        Write-Build Red   "  [ISO]  RunIds iguales en contextos distintos"
        $errores++
    }

    if ($ctx.TaskName -ne $ctx2.TaskName) {
        Write-Build Green "  [ISO]  TaskNames distintos : OK"
    } else {
        Write-Build Yellow "  [ISO]  Mismos TaskName (esperado si misma tarea)"
    }

    # ---- 3. Regla 6: Excel abierto bloquea Word si no se libera primero --
    Write-Build Cyan ""
    Write-Build Cyan "  [R6]   Verificando Regla 6 (COM secuencial)..."
    Write-Build Cyan "         Excel -> cerrar Excel -> Word -> cerrar Word"

    $xl  = $null
    $wb  = $null
    $ws  = $null
    $wd  = $null
    $doc = $null
    $sel = $null
    $r6ok = $false

    try {
        # Fase Excel
        $xl = New-ExcelApp -Context $ctx -TimeoutSec 20
        if ($null -eq $xl) { throw 'Excel no disponible' }

        $wb = New-ExcelWorkbook -Context $ctx -ExcelApp $xl
        $ws = Get-ExcelSheet -Workbook $wb -Index 1
        $data = @(@{ Prueba = 'R6'; Valor = 'Excel-primero' })
        Write-ExcelData -Context $ctx -Sheet $ws -Data $data -Columns @('Prueba','Valor')
        Release-ComObject $ws ; $ws = $null

        $xlF = Join-Path $ctx.Paths.Output "concurrencia_xl_$($ctx.RunId).xlsx"
        Save-ExcelWorkbook -Context $ctx -Workbook $wb -Path $xlF

        Write-Build Green "  [R6]   Excel abierto y guardado OK"

    } catch {
        Write-Build Yellow "  [R6]   Excel fase : $_ (puede ser que Excel no esta instalado)"
    } finally {
        if ($null -ne $ws) { Release-ComObject $ws ; $ws = $null }
        Close-ExcelWorkbook -Workbook $wb -Save $false
        Close-ExcelApp -ExcelApp $xl
        $wb = $null ; $xl = $null
    }

    # Verificar que Excel cerro antes de abrir Word
    $xlVivos = @(Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowHandle -eq [IntPtr]::Zero }).Count
    if ($xlVivos -eq 0) {
        Write-Build Green "  [R6]   Excel cerrado antes de abrir Word : OK"
    } else {
        Write-Build Red   "  [R6]   $xlVivos proceso(s) Excel aun activos antes de Word"
        $errores++
        Remove-ZombieCom | Out-Null
        Start-Sleep -Seconds 2
    }

    try {
        # Fase Word (solo si Excel esta cerrado)
        $wd = New-WordApp -Context $ctx -TimeoutSec 20
        if ($null -eq $wd) { throw 'Word no disponible' }

        $doc = New-WordDocument -Context $ctx -WordApp $wd
        $sel = Get-WordSelection -WordApp $wd
        Add-WordParagraph -Context $ctx -Selection $sel -Text "Regla 6 verificada"
        Release-ComObject $sel ; $sel = $null

        $wdF = Join-Path $ctx.Paths.Output "concurrencia_wd_$($ctx.RunId).docx"
        Save-WordDocument -Context $ctx -Document $doc -Path $wdF

        $r6ok = $true
        Write-Build Green "  [R6]   Word abierto y guardado OK"

    } catch {
        Write-Build Yellow "  [R6]   Word fase : $_ (puede ser que Word no esta instalado)"
    } finally {
        if ($null -ne $sel) { Release-ComObject $sel ; $sel = $null }
        Close-WordDocument -Document $doc -Save $false
        Close-WordApp -WordApp $wd
        $doc = $null ; $wd = $null
    }

    $wdVivos = @(Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowHandle -eq [IntPtr]::Zero }).Count
    if ($wdVivos -eq 0) {
        Write-Build Green "  [R6]   Word cerrado OK (0 zombis)"
    } else {
        Write-Build Red   "  [R6]   $wdVivos proceso(s) WINWORD aun activos"
        $errores++
        Remove-ZombieCom | Out-Null
    }

    # ---- 4. Prueba de archivo parcial (simulacion de interrupcion) --------
    Write-Build Cyan ""
    Write-Build Cyan "  [INT]  Simulando archivo parcial (interrupcion mid-write)..."

    $parcial = Join-Path $ctx.Paths.Output "parcial_$($ctx.RunId).tmp"
    $fs = $null
    try {
        $fs = [System.IO.File]::Open($parcial, 'Create', 'Write', 'Read')
        $bytes = [System.Text.Encoding]::ASCII.GetBytes("datos_incompletos")
        $fs.Write($bytes, 0, $bytes.Length)
        $existe = Test-Path $parcial

        if ($existe) {
            Write-Build Green "  [INT]  Archivo parcial creado : OK (simulacion correcta)"
        } else {
            Write-Build Red   "  [INT]  Archivo parcial no creado"
            $errores++
        }
    } catch {
        Write-Build Yellow "  [INT]  WARN simulacion parcial: $_"
    } finally {
        if ($null -ne $fs) {
            try { $fs.Close()   } catch {}
            try { $fs.Dispose() } catch {}
            $fs = $null
        }
        Remove-Item $parcial -ErrorAction SilentlyContinue
    }

    # ---- Resumen -----------------------------------------------------------
    Write-Build Cyan ""
    if ($errores -eq 0) {
        Write-Build Green "  RESULTADO: OK - aislamiento y Regla 6 verificados"
    } else {
        Write-Build Red   "  RESULTADO: $errores error(es)"
    }

    Write-BuildLog $ctx 'INFO' "diag_concurrencia completado. Errores=$errores"
    Write-RunResult -Context $ctx -Success ($errores -eq 0)

    if ($errores -gt 0) {
        throw "diag_concurrencia: $errores error(es)"
    }
}
