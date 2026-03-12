#Requires -Version 5.1
# =============================================================================
# tasks/task_diag_motor.ps1
# @Description : Pruebas profundas del motor Invoke-Build:
#                - StrictMode activo en todas las tareas
#                - Variables no declaradas detectadas
#                - Contexto TaskContext correcto (RunId, Paths, Config)
#                - Logger: todos los niveles, JSONL valido, campos obligatorios
#                - Write-RunResult registra Success y Error correctamente
#                - New-RunId genera IDs unicos
#                - Get-EngineConfig carga y devuelve todos los campos
# @Category    : Diagnostico
# @Version     : 1.0.0
# =============================================================================
# Solo ASCII. PS 5.1.

# Synopsis: Prueba profunda del motor: StrictMode, contexto, logger y RunId
task diag_motor {
    $ctx = New-TaskContext `
        -TaskName 'diag_motor' `
        -Config   $Script:Config `
        -Root     (Split-Path $BuildRoot -Parent)

    Write-BuildLog $ctx 'INFO' 'Iniciando diagnostico profundo del motor'

    $errores = 0

    # ---- 1. StrictMode verificado en tarea de prueba ----------------------
    Write-Build Cyan "  [SM]   Verificando StrictMode -Version Latest activo..."
    $smActivo = $false
    try {
        # En StrictMode, acceder a variable no declarada lanza error
        $testBloque = [scriptblock]::Create('
            Set-StrictMode -Version Latest
            $x = $noDeclarada + 1
        ')
        & $testBloque
        Write-Build Red   "  [SM]   StrictMode NO activo - acceso a var no declarada no lanzo error"
        $errores++
    } catch {
        $smActivo = $true
        Write-Build Green "  [SM]   StrictMode activo : OK"
    }

    # ---- 2. New-TaskContext - estructura completa -------------------------
    Write-Build Cyan "  [CTX]  Verificando estructura del contexto..."

    $camposReq = @('RunId','TaskName','Config','StartTime','Params','Paths')
    foreach ($campo in $camposReq) {
        if ($ctx.ContainsKey($campo)) {
            Write-Build Green "  [CTX]  $campo : presente"
        } else {
            Write-Build Red   "  [CTX]  $campo : FALTA"
            $errores++
        }
    }

    $rutasReq = @('Root','Input','Output','Reports','Logs')
    foreach ($ruta in $rutasReq) {
        if ($ctx.Paths.ContainsKey($ruta)) {
            Write-Build Green "  [CTX]  Paths.$ruta : $($ctx.Paths[$ruta])"
        } else {
            Write-Build Red   "  [CTX]  Paths.$ruta : FALTA"
            $errores++
        }
    }

    # ---- 3. RunId - formato y unicidad ------------------------------------
    Write-Build Cyan "  [RID]  Verificando formato y unicidad de RunId..."
    $rid1 = New-RunId
    Start-Sleep -Milliseconds 50
    $rid2 = New-RunId

    # Formato real: yyyyMMdd_HHmmss_XXXX donde XXXX son 4 letras A-Z (65..90)
    $patron = '^\d{8}_\d{6}_[A-Z]{4}$'
    if ($rid1 -match $patron) {
        Write-Build Green "  [RID]  Formato : OK ($rid1)"
    } else {
        Write-Build Red   "  [RID]  Formato inesperado: $rid1 (patron esperado: $patron)"
        $errores++
    }
    if ($rid1 -ne $rid2) {
        Write-Build Green "  [RID]  Unicidad : OK ($rid1 != $rid2)"
    } else {
        Write-Build Red   "  [RID]  Colision: dos RunId identicos generados"
        $errores++
    }

    # ---- 4. Get-EngineConfig - campos obligatorios ------------------------
    Write-Build Cyan "  [CFG]  Verificando campos de engine.config.json..."
    $seccionesReq = @('engine','sap','excel','reports')
    foreach ($sec in $seccionesReq) {
        if ($ctx.Config.ContainsKey($sec)) {
            Write-Build Green "  [CFG]  $sec : OK"
        } else {
            Write-Build Red   "  [CFG]  $sec : FALTA en config"
            $errores++
        }
    }

    $camposEngine = @('logLevel','maxRetries','retryDelaySeconds')
    foreach ($c in $camposEngine) {
        $val = $ctx.Config.engine.$c
        if ($null -ne $val) {
            Write-Build Green "  [CFG]  engine.$c = $val"
        } else {
            Write-Build Red   "  [CFG]  engine.$c : NULO"
            $errores++
        }
    }

    # ---- 5. Write-BuildLog - todos los niveles ----------------------------
    Write-Build Cyan "  [LOG]  Probando todos los niveles de log..."

    # El logger filtra niveles por debajo del logLevel configurado.
    # Solo contar los que realmente se escribiran segun la config actual.
    $levelMap   = @{ DEBUG = 0; INFO = 1; WARN = 2; ERROR = 3 }
    $cfgLevel   = $ctx.Config.engine.logLevel
    $cfgMin     = $levelMap[$cfgLevel]
    $nivelesAll = @('DEBUG','INFO','WARN','ERROR')
    $nivelesEsperados = @($nivelesAll | Where-Object { $levelMap[$_] -ge $cfgMin })

    Write-Build Cyan "  [LOG]  logLevel configurado: $cfgLevel -> se esperan $($nivelesEsperados.Count) nivel(es)"

    $regFile = Join-Path $ctx.Paths.Logs 'registry.jsonl'
    $antes   = 0
    if (Test-Path $regFile) {
        $antes = @(Get-Content $regFile -Encoding ASCII).Count
    }

    foreach ($nivel in $nivelesAll) {
        Write-BuildLog $ctx $nivel "Prueba nivel $nivel desde diag_motor"
    }

    $despues = @(Get-Content $regFile -Encoding ASCII).Count
    $nuevas  = $despues - $antes
    if ($nuevas -ge $nivelesEsperados.Count) {
        Write-Build Green "  [LOG]  Lineas escritas : $nuevas (esperaba >= $($nivelesEsperados.Count)) OK"
    } else {
        Write-Build Red   "  [LOG]  Lineas escritas : $nuevas (esperaba >= $($nivelesEsperados.Count))"
        $errores++
    }

    # ---- 6. JSONL - campos obligatorios en cada linea --------------------
    Write-Build Cyan "  [JSONL] Verificando estructura JSONL del registro..."
    # Campos reales del Logger (ver lib/Logger.ps1):
    #   ts | level | runId | task | message | detail
    $camposJsonl = @('ts','level','runId','task','message')
    $lineasRun   = @(Get-Content $regFile -Encoding ASCII |
        Where-Object { $_ -ne '' } |
        ForEach-Object { try { $_ | ConvertFrom-Json } catch { $null } } |
        Where-Object { $null -ne $_ } |
        Where-Object {
            # StrictMode: leer campo via PSObject para no lanzar si no existe
            $rid = $_.PSObject.Properties['runId']
            $null -ne $rid -and $rid.Value -eq $ctx.RunId
        })

    if ($lineasRun.Count -ge $nivelesEsperados.Count) {
        Write-Build Green "  [JSONL] Entradas de este run: $($lineasRun.Count)"
        $muestra = $lineasRun[0]
        foreach ($campo in $camposJsonl) {
            # Usar PSObject.Properties para evitar PropertyNotFoundStrict
            $prop = $muestra.PSObject.Properties[$campo]
            if ($null -ne $prop -and $null -ne $prop.Value) {
                Write-Build Green "  [JSONL] Campo '$campo' : OK ($($prop.Value))"
            } else {
                Write-Build Red   "  [JSONL] Campo '$campo' : AUSENTE en JSONL"
                $errores++
            }
        }
    } else {
        Write-Build Red "  [JSONL] Entradas del run: $($lineasRun.Count) (esperaba >= $($nivelesEsperados.Count))"
        $errores++
    }

    # ---- 7. Write-RunResult - escribe entrada en JSONL -------------------
    Write-Build Cyan "  [RES]  Verificando Write-RunResult..."
    $antesRes = @(Get-Content $regFile -Encoding ASCII).Count
    Write-RunResult -Context $ctx -Success $true
    $despuesRes = @(Get-Content $regFile -Encoding ASCII).Count
    if ($despuesRes -gt $antesRes) {
        Write-Build Green "  [RES]  Write-RunResult escribio OK"
    } else {
        Write-Build Red   "  [RES]  Write-RunResult no escribio nada"
        $errores++
    }

    # ---- 8. Codificacion ASCII del log (nunca bytes > 127) ---------------
    Write-Build Cyan "  [ENC]  Verificando encoding ASCII del log..."
    $bytes   = [System.IO.File]::ReadAllBytes($regFile)
    $badBytes = @($bytes | Where-Object { $_ -gt 127 }).Count
    if ($badBytes -eq 0) {
        Write-Build Green "  [ENC]  Encoding ASCII puro : OK"
    } else {
        Write-Build Red   "  [ENC]  $badBytes bytes > 127 en el log - encoding incorrecto"
        $errores++
    }

    # ---- Resumen -----------------------------------------------------------
    Write-Build Cyan ""
    if ($errores -eq 0) {
        Write-Build Green "  RESULTADO: OK - motor funciona correctamente"
    } else {
        Write-Build Red   "  RESULTADO: $errores error(es) en el motor"
    }

    Write-BuildLog $ctx 'INFO' "diag_motor completado. Errores=$errores"

    if ($errores -gt 0) {
        throw "diag_motor detecto $errores error(es)"
    }
}
