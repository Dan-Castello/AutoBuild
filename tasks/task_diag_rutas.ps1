#Requires -Version 5.1
# =============================================================================
# tasks/task_diag_rutas.ps1
# @Description : Diagnostico de rutas y permisos del sistema de archivos.
#                Verifica lectura, escritura y eliminacion en cada carpeta
#                de trabajo del proyecto. Detecta problemas de permisos antes
#                de que fallen las tareas de produccion.
# @Category    : Utilidad
# @Version     : 1.0.0
# =============================================================================
# Solo ASCII. PS 5.1. Sin COM.

# Synopsis: Verifica permisos de lectura/escritura en input, output, reports y logs
task diag_rutas {
    $ctx = New-TaskContext `
        -TaskName 'diag_rutas' `
        -Config   $Script:Config `
        -Root     (Split-Path $BuildRoot -Parent)

    Write-BuildLog $ctx 'INFO' 'Iniciando diagnostico de rutas y permisos'

    $errores = 0
    $Root    = Split-Path $BuildRoot -Parent

    # ---- Funcion local: prueba completa R/W/D en una carpeta ---------------
    # No se define como funcion global (regla del proyecto)
    $testDir = {
        param([string]$DirPath, [string]$DirName, [hashtable]$Ctx)
        $ok = $true

        # Existencia
        if (-not (Test-Path $DirPath)) {
            Write-Build Red   "  [$DirName] NO EXISTE: $DirPath"
            return $false
        }

        # Escritura
        $tmpFile = Join-Path $DirPath "diag_rutas_test_$($Ctx.RunId).tmp"
        try {
            [System.IO.File]::WriteAllText($tmpFile, 'test', [System.Text.Encoding]::ASCII)
            Write-Build Green "  [$DirName] Escritura          : OK"
        } catch {
            Write-Build Red   "  [$DirName] Escritura          : FAIL - $_"
            $ok = $false
        }

        # Lectura
        if (Test-Path $tmpFile) {
            try {
                $contenido = [System.IO.File]::ReadAllText($tmpFile, [System.Text.Encoding]::ASCII)
                if ($contenido -eq 'test') {
                    Write-Build Green "  [$DirName] Lectura            : OK"
                } else {
                    Write-Build Red   "  [$DirName] Lectura            : contenido incorrecto"
                    $ok = $false
                }
            } catch {
                Write-Build Red   "  [$DirName] Lectura            : FAIL - $_"
                $ok = $false
            }

            # Eliminacion
            try {
                Remove-Item $tmpFile -Force -ErrorAction Stop
                Write-Build Green "  [$DirName] Eliminacion        : OK"
            } catch {
                Write-Build Red   "  [$DirName] Eliminacion        : FAIL - $_"
                $ok = $false
            }
        }

        # Espacio libre
        try {
            $drive = [System.IO.Path]::GetPathRoot($DirPath)
            $disco = Get-WmiObject -Class Win32_LogicalDisk `
                -Filter "DeviceID='$($drive.TrimEnd('\'))'" -ErrorAction SilentlyContinue
            if ($null -ne $disco) {
                $libreGB = [math]::Round($disco.FreeSpace / 1GB, 1)
                $totalGB = [math]::Round($disco.Size / 1GB, 1)
                $color   = if ($libreGB -lt 1) { 'Yellow' } else { 'Cyan' }
                Write-Build $color "  [$DirName] Espacio libre      : ${libreGB} GB de ${totalGB} GB"
                if ($libreGB -lt 0.5) {
                    Write-Build Red "  [$DirName] WARN: menos de 500 MB libres"
                    $ok = $false
                }
            }
        } catch {}

        return $ok
    }

    # ---- Carpetas de trabajo -----------------------------------------------
    $carpetas = [ordered]@{
        'INPUT'   = $ctx.Paths.Input
        'OUTPUT'  = $ctx.Paths.Output
        'REPORTS' = $ctx.Paths.Reports
        'LOGS'    = $ctx.Paths.Logs
    }

    foreach ($kv in $carpetas.GetEnumerator()) {
        Write-Build Cyan  ""
        Write-Build Cyan  "  --- $($kv.Key): $($kv.Value) ---"
        $resultado = & $testDir $kv.Value $kv.Key $ctx
        if (-not $resultado) { $errores++ }
    }

    # ---- Carpetas del motor (solo lectura necesaria) -----------------------
    Write-Build Cyan  ""
    Write-Build Cyan  "  --- Motor y librerias (solo lectura) ---"
    $carpetasMotor = [ordered]@{
        'ENGINE' = Join-Path $Root 'engine'
        'LIB'    = Join-Path $Root 'lib'
        'TASKS'  = Join-Path $Root 'tasks'
        'IB'     = Join-Path $Root 'tools\InvokeBuild'
    }

    foreach ($kv in $carpetasMotor.GetEnumerator()) {
        if (Test-Path $kv.Value) {
            $archivos = @(Get-ChildItem -Path $kv.Value -Filter '*.ps1' -ErrorAction SilentlyContinue)
            Write-Build Green "  [$($kv.Key)] $($kv.Value) : OK ($($archivos.Count) .ps1)"
        } else {
            Write-Build Red   "  [$($kv.Key)] FALTA: $($kv.Value)"
            $errores++
        }
    }

    # ---- Verificar config --------------------------------------------------
    Write-Build Cyan  ""
    $cfgFile = Join-Path $Root 'engine.config.json'
    if (Test-Path $cfgFile) {
        try {
            $cfg = Get-Content $cfgFile -Raw | ConvertFrom-Json
            Write-Build Green "  [CFG]  engine.config.json  : OK (logLevel=$($cfg.engine.logLevel))"
        } catch {
            Write-Build Red   "  [CFG]  engine.config.json  : JSON invalido - $_"
            $errores++
        }
    } else {
        Write-Build Yellow "  [CFG]  engine.config.json  : no existe (se usaran defaults)"
    }

    # ---- Resumen -----------------------------------------------------------
    Write-Build Cyan ""
    if ($errores -eq 0) {
        Write-Build Green "  RESULTADO: OK - rutas y permisos correctos"
    } else {
        Write-Build Red   "  RESULTADO: $errores problema(s) de rutas o permisos"
    }

    Write-BuildLog $ctx 'INFO' "diag_rutas completado. Errores=$errores"
    Write-RunResult -Context $ctx -Success ($errores -eq 0)

    if ($errores -gt 0) {
        throw "diag_rutas detecto $errores problema(s)"
    }
}
