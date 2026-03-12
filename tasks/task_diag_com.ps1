#Requires -Version 5.1
# =============================================================================
# tasks/task_diag_com.ps1
# @Description : Diagnostico del subsistema COM en PS 5.1.
#                Comprueba servidores COM relevantes, detecta procesos zombi,
#                verifica que el mecanismo de timeout funciona correctamente.
# @Category    : Utilidad
# @Version     : 1.0.0
# =============================================================================
# Solo ASCII. PS 5.1. No instancia Excel con datos; solo prueba el mecanismo.

# Synopsis: Verifica servidores COM disponibles, timeout y estado de procesos
task diag_com {
    $ctx = New-TaskContext `
        -TaskName 'diag_com' `
        -Config   $Script:Config `
        -Root     (Split-Path $BuildRoot -Parent)

    Write-BuildLog $ctx 'INFO' 'Iniciando diagnostico COM'

    $errores = 0
    $avisos  = 0

    # ---- 0. Limpiar zombis de ejecuciones previas --------------------------
    # Un proceso WINWORD/EXCEL sin ventana de una ejecucion anterior bloquea
    # la instanciacion COM y contamina el inventario de procesos.
    $zombisPrevios = Remove-ZombieCom
    if ($zombisPrevios -gt 0) {
        Write-Build Yellow "  [PRE]  Zombis de ejecucion previa eliminados: $zombisPrevios"
        Write-BuildLog $ctx 'WARN' "Zombis previos eliminados: $zombisPrevios"
        Start-Sleep -Seconds 2   # Dar tiempo al OS para liberar handles
    } else {
        Write-Build Green "  [PRE]  Sin procesos zombi previos"
    }

    # ---- 1. Inventario de servidores COM de interes ------------------------
    $servidores = [ordered]@{
        'Excel.Application'          = 'Excel (Office 16)'
        'Word.Application'           = 'Word (Office 16)'
        'Scripting.FileSystemObject' = 'FileSystemObject (siempre disponible)'
        'WScript.Shell'              = 'WScript.Shell (siempre disponible)'
        'Shell.Application'          = 'Shell.Application'
        'MSXML2.DOMDocument.6.0'     = 'MSXML 6.0'
    }

    Write-Build Cyan  "  [SCAN] Comprobando servidores COM (5s por servidor)..."
    foreach ($kv in $servidores.GetEnumerator()) {
        $disponible = Test-ComAvailable -ProgId $kv.Key -TimeoutSec 5
        $label      = $kv.Value
        $progid     = $kv.Key
        if ($disponible) {
            Write-Build Green "  [COM]  $label"
            Write-Build Green "         ProgId: $progid -> disponible"
        } else {
            # WScript y FSO siempre deben responder: si no, es error grave
            if ($progid -in @('Scripting.FileSystemObject','WScript.Shell')) {
                Write-Build Red   "  [COM]  $label"
                Write-Build Red   "         ProgId: $progid -> ERROR: debe estar siempre disponible"
                $errores++
            } else {
                Write-Build Yellow "  [COM]  $label"
                Write-Build Yellow "         ProgId: $progid -> no disponible (normal si la app no esta abierta)"
                $avisos++
            }
        }
    }

    # ---- 2. Verificar mecanismo de timeout (Job + Wait-Job) ----------------
    Write-Build Cyan  ""
    Write-Build Cyan  "  [MECH] Verificando mecanismo de timeout COM..."
    $jobTest = $null
    try {
        $jobTest = Start-Job -ScriptBlock {
            Start-Sleep -Seconds 1
            return 'job-ok'
        }
        $done = Wait-Job -Job $jobTest -Timeout 10
        if ($null -ne $done) {
            $result = Receive-Job -Job $jobTest
            if ($result -eq 'job-ok') {
                Write-Build Green "  [MECH] Start-Job/Wait-Job  : OK (Jobs funcionan en este entorno)"
            } else {
                Write-Build Yellow "  [MECH] Start-Job/Wait-Job  : resultado inesperado '$result'"
                $avisos++
            }
        } else {
            Write-Build Red   "  [MECH] Start-Job/Wait-Job  : TIMEOUT inesperado en job de 1s"
            Write-Build Red   "         Los Jobs de PS pueden estar bloqueados por politica"
            $errores++
        }
    } catch {
        Write-Build Red   "  [MECH] Start-Job           : FAIL - $_"
        Write-Build Red   "         Sin Jobs, el mecanismo anti-freeze COM no funciona"
        $errores++
    } finally {
        if ($null -ne $jobTest) {
            Remove-Job -Job $jobTest -Force -ErrorAction SilentlyContinue
        }
    }

    # ---- 3. Inventario de procesos COM activos actualmente -----------------
    Write-Build Cyan  ""
    Write-Build Cyan  "  [PROC] Inventario de procesos COM activos..."
    $procsInteres = @('EXCEL','WINWORD','OUTLOOK','POWERPNT')
    foreach ($proc in $procsInteres) {
        $instancias = @(Get-Process -Name $proc -ErrorAction SilentlyContinue)
        if ($instancias.Count -eq 0) {
            Write-Build Cyan  "  [PROC] $proc : no en ejecucion"
        } else {
            foreach ($p in $instancias) {
                $ventana = if ($p.MainWindowHandle -ne [IntPtr]::Zero) { 'CON ventana' } else { 'SIN ventana (posible zombi)' }
                $color   = if ($p.MainWindowHandle -ne [IntPtr]::Zero) { 'Cyan' } else { 'Yellow' }
                Write-Build $color "  [PROC] $proc PID=$($p.Id) : $ventana"
                if ($p.MainWindowHandle -eq [IntPtr]::Zero) { $avisos++ }
            }
        }
    }

    # ---- 4. Test de FileSystemObject (COM ligero, siempre debe funcionar) --
    Write-Build Cyan  ""
    Write-Build Cyan  "  [FSO]  Prueba funcional de FileSystemObject..."
    $fso = $null
    $ts  = $null
    try {
        $fso = New-Object -ComObject 'Scripting.FileSystemObject' -ErrorAction Stop
        $tmpPath = Join-Path $ctx.Paths.Output "diag_com_fso_$($ctx.RunId).tmp"
        $ts = $fso.CreateTextFile($tmpPath, $true)
        $ts.WriteLine('AutoBuild COM test OK')
        $ts.Close()
        Release-ComObject $ts
        $ts = $null

        if (Test-Path $tmpPath) {
            $contenido = Get-Content $tmpPath -Encoding Default
            if ($contenido -like '*AutoBuild COM test OK*') {
                Write-Build Green "  [FSO]  Escritura via COM   : OK"
            } else {
                Write-Build Red   "  [FSO]  Escritura via COM   : contenido inesperado"
                $errores++
            }
            Remove-Item $tmpPath -Force -ErrorAction SilentlyContinue
        } else {
            Write-Build Red "  [FSO]  Archivo no creado por FSO"
            $errores++
        }
    } catch {
        Write-Build Red   "  [FSO]  FAIL - $_"
        $errores++
    } finally {
        if ($null -ne $ts) {
            try { $ts.Close() } catch {}
            Release-ComObject $ts
            $ts = $null
        }
        if ($null -ne $fso) {
            Release-ComObject $fso
            $fso = $null
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        Write-Build Green "  [GC]   Liberacion COM FSO  : OK"
    }

    # ---- Resumen -----------------------------------------------------------
    Write-Build Cyan  ""
    if ($errores -eq 0 -and $avisos -eq 0) {
        Write-Build Green "  RESULTADO: OK - subsistema COM en buen estado"
    } elseif ($errores -eq 0) {
        Write-Build Yellow "  RESULTADO: $avisos aviso(s) - revisar procesos COM activos"
    } else {
        Write-Build Red   "  RESULTADO: $errores error(es), $avisos aviso(s)"
    }

    Write-BuildLog $ctx 'INFO' "diag_com completado. Errores=$errores Avisos=$avisos"
    Write-RunResult -Context $ctx -Success ($errores -eq 0)

    if ($errores -gt 0) {
        throw "diag_com detecto $errores error(es)"
    }
}
