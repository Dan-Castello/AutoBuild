#Requires -Version 5.1
# =============================================================================
# tasks/task_diag_log.ps1
# @Description : Diagnostico del sistema de logging: escritura JSONL,
#                rotacion, lectura y filtracion del registro de ejecuciones.
# @Category    : Utilidad
# @Version     : 1.0.0
# =============================================================================
# Solo ASCII. PS 5.1. Sin COM.

# Synopsis: Verifica escritura, lectura y formato del registro JSONL de logs
task diag_log {
    $ctx = New-TaskContext `
        -TaskName 'diag_log' `
        -Config   $Script:Config `
        -Root     (Split-Path $BuildRoot -Parent)

    Write-BuildLog $ctx 'INFO' 'Iniciando diagnostico del sistema de log'

    $errores = 0
    $logDir  = $ctx.Paths.Logs
    $regFile = Join-Path $logDir 'registry.jsonl'

    # ---- 1. Verificar carpeta de logs --------------------------------------
    if (Test-Path $logDir) {
        Write-Build Green "  [DIR]  logs\              : $logDir OK"
    } else {
        Write-Build Red   "  [DIR]  logs\              : FALTA - $logDir"
        $errores++
        throw "Carpeta logs no existe - ejecutar Setup.ps1"
    }

    # ---- 2. Prueba de escritura de todos los niveles -----------------------
    Write-Build Cyan  "  [WRITE] Probando Write-BuildLog en todos los niveles..."

    # Forzar nivel DEBUG temporalmente para verificar todos los niveles
    $nivelOriginal = $ctx.Config.engine.logLevel
    $ctx.Config.engine.logLevel = 'DEBUG'

    Write-BuildLog $ctx 'DEBUG' 'Prueba nivel DEBUG  - desde diag_log'
    Write-BuildLog $ctx 'INFO'  'Prueba nivel INFO   - desde diag_log'
    Write-BuildLog $ctx 'WARN'  'Prueba nivel WARN   - desde diag_log'
    Write-BuildLog $ctx 'ERROR' 'Prueba nivel ERROR  - desde diag_log'

    # Restaurar nivel original
    $ctx.Config.engine.logLevel = $nivelOriginal
    Write-Build Green "  [WRITE] Escritura de log   : OK (4 niveles registrados)"

    # ---- 3. Verificar que registry.jsonl existe y tiene contenido ----------
    if (Test-Path $regFile) {
        $lineas = @(Get-Content $regFile -Encoding ASCII).Count
        Write-Build Green "  [FILE] registry.jsonl     : $lineas lineas acumuladas"
    } else {
        Write-Build Red   "  [FILE] registry.jsonl     : no creado"
        $errores++
    }

    # ---- 4. Verificar formato JSONL (cada linea es JSON valido) ------------
    Write-Build Cyan  "  [FMT]  Verificando formato JSONL (ultimas 20 lineas)..."
    $jsonFail = 0
    try {
        $lineasRecientes = Get-Content $regFile -Encoding ASCII -Tail 20
        foreach ($linea in $lineasRecientes) {
            if ([string]::IsNullOrWhiteSpace($linea)) { continue }
            try {
                $obj = $linea | ConvertFrom-Json
                # Campos obligatorios
                $campos = @('ts','level','runId','task','message')
                foreach ($campo in $campos) {
                    if ($null -eq $obj.$campo) {
                        Write-Build Yellow "  [FMT]  Campo '$campo' faltante en: $linea"
                        $jsonFail++
                    }
                }
            } catch {
                Write-Build Red   "  [FMT]  Linea no es JSON valido: $linea"
                $jsonFail++
            }
        }
        if ($jsonFail -eq 0) {
            Write-Build Green "  [FMT]  Formato JSONL      : OK (20 lineas verificadas)"
        } else {
            Write-Build Red   "  [FMT]  $jsonFail lineas con formato incorrecto"
            $errores++
        }
    } catch {
        Write-Build Yellow "  [FMT]  WARN al verificar JSONL: $_"
    }

    # ---- 5. Prueba de filtracion del log (workflow tipico del operador) ----
    Write-Build Cyan  "  [QUERY] Probando consulta del registro..."
    try {
        $todos = @(Get-Content $regFile -Encoding ASCII |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            ForEach-Object {
                try { $_ | ConvertFrom-Json } catch { $null }
            } |
            Where-Object { $null -ne $_ })

        $porNivel = @($todos | Group-Object -Property level)
        Write-Build Green "  [QUERY] Registros totales  : $($todos.Count)"
        foreach ($g in $porNivel | Sort-Object Name) {
            Write-Build Cyan  "  [QUERY]   $($g.Name.PadRight(6)) : $($g.Count) entradas"
        }

        # Filtrar los de esta tarea (RunId actual) via PSObject para StrictMode
        $estaRun = @($todos | Where-Object {
            $p = $_.PSObject.Properties['runId']
            $null -ne $p -and $p.Value -eq $ctx.RunId
        })
        Write-Build Green "  [QUERY] Esta ejecucion     : $($estaRun.Count) entradas (RunId=$($ctx.RunId))"

    } catch {
        Write-Build Yellow "  [QUERY] WARN al consultar registro: $_"
    }

    # ---- 6. Verificar encoding del archivo (debe ser ASCII) ----------------
    Write-Build Cyan  "  [ENC]  Verificando encoding del archivo de log..."
    $bytes    = [System.IO.File]::ReadAllBytes($regFile)
    $badBytes = @($bytes | Where-Object { $_ -gt 127 }).Count
    if ($badBytes -eq 0) {
        Write-Build Green "  [ENC]  Encoding            : OK (ASCII puro)"
    } else {
        Write-Build Red   "  [ENC]  $badBytes bytes no-ASCII en registry.jsonl"
        Write-Build Red   "         Puede causar problemas de lectura en otros sistemas"
        $errores++
    }

    # ---- 7. Calcular tamano del archivo ------------------------------------
    $size = (Get-Item $regFile).Length
    $sizeKB = [math]::Round($size / 1KB, 1)
    Write-Build Cyan  "  [SIZE] Tamano registry.jsonl : ${sizeKB} KB"
    if ($size -gt 10MB) {
        Write-Build Yellow "  [SIZE] WARN: archivo >10 MB, considerar rotacion manual"
        $avisos++
    }

    # ---- Resumen -----------------------------------------------------------
    Write-Build Cyan  ""
    if ($errores -eq 0) {
        Write-Build Green "  RESULTADO: OK - sistema de logging funciona correctamente"
    } else {
        Write-Build Red   "  RESULTADO: $errores error(es) en el sistema de logging"
    }

    Write-BuildLog $ctx 'INFO' "diag_log completado. Errores=$errores"
    Write-RunResult -Context $ctx -Success ($errores -eq 0)

    if ($errores -gt 0) {
        throw "diag_log detecto $errores error(es)"
    }
}
