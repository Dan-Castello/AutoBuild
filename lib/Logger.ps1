#Requires -Version 5.1
# lib/Logger.ps1
# Logging estructurado a JSONL. Solo ASCII. PS 5.1.
# Entrada del registro: logs/registry.jsonl
#
# RESPONSABILIDADES DE ESTE ARCHIVO:
#   - Get-EngineConfig  : carga y validacion de configuracion
#   - New-RunId         : generacion de identificadores de ejecucion
#   - Write-BuildLog    : escritura atomica de entradas JSONL
#   - Write-RunResult   : escritura de resultado de ejecucion
#
# FIX-LOGGER-01 (Escritura atomica):
#   Add-Content no es atomico bajo escrituras concurrentes.
#   Se usa System.Threading.Mutex con nombre para serializar accesos
#   al archivo registry.jsonl. Compatibilidad: PS 5.1 / .NET 4.x.
#   El mutex se crea, se aguarda maximo 5 segundos y se libera en finally.
#   Nombre del mutex: 'Global\AutoBuildLogMutex'
#   El prefijo Global\ permite sincronizar entre sesiones distintas
#   de PowerShell en el mismo equipo (ej. Build-Parallel).

Set-StrictMode -Version Latest

# ---------------------------------------------------------------------------
# CONFIGURACION
# ---------------------------------------------------------------------------

function Get-EngineConfig {
    <#
    .SYNOPSIS
        Carga engine.config.json, combina con defaults y retorna
        un nuevo hashtable independiente cada vez que se invoca.
        Garantiza que existan las carpetas de trabajo.
    .NOTES
        FIX-CONFIG-01: Esta funcion siempre retorna un NUEVO hashtable.
        El llamante (engine/Main.build.ps1) almacena el resultado en
        $Script:EngineConfig. New-TaskContext copia las secciones
        relevantes al construir el contexto de tarea, de modo que
        mutaciones en el contexto no afectan $Script:EngineConfig.
    #>
    param([string]$Root)

    $cfgFile = Join-Path $Root 'engine.config.json'

    # Defaults internos — valores seguros para produccion
    $cfg = @{
        engine  = @{ logLevel = 'INFO'; maxRetries = 3; retryDelaySeconds = 5 }
        sap     = @{ systemId = 'PRD'; client = '800'; language = 'ES'; timeout = 180 }
        excel   = @{ visible = $false; screenUpdating = $false }
        reports = @{ defaultFormat = 'xlsx'; retentionDays = 30 }
    }

    # Merge desde archivo si existe
    if (Test-Path $cfgFile) {
        try {
            $raw = Get-Content $cfgFile -Raw | ConvertFrom-Json
            foreach ($key in @('engine','sap','excel','reports')) {
                if ($raw.$key) {
                    foreach ($prop in $raw.$key.PSObject.Properties) {
                        $cfg[$key][$prop.Name] = $prop.Value
                    }
                }
            }
        } catch {
            Write-Host "WARN: No se pudo leer engine.config.json: $_" -ForegroundColor Yellow
        }
    }

    # Garantizar existencia de carpetas de trabajo
    $workDirs = @(
        (Join-Path $Root 'logs'),
        (Join-Path $Root 'input'),
        (Join-Path $Root 'output'),
        (Join-Path $Root 'reports')
    )
    foreach ($dir in $workDirs) {
        if (-not (Test-Path $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
    }

    return $cfg
}

# ---------------------------------------------------------------------------
# RUN ID
# ---------------------------------------------------------------------------

function New-RunId {
    <#
    .SYNOPSIS
        Genera un identificador de ejecucion unico con formato yyyyMMdd_HHmmss_XXXX.
    #>
    $ts   = Get-Date -Format 'yyyyMMdd_HHmmss'
    $rand = -join ((65..90) | Get-Random -Count 4 | ForEach-Object { [char]$_ })
    return "${ts}_${rand}"
}

# ---------------------------------------------------------------------------
# LOGGING ATOMICO
# ---------------------------------------------------------------------------

function Write-BuildLog {
    <#
    .SYNOPSIS
        Escribe una entrada JSONL en registry.jsonl de forma atomica.
    .NOTES
        FIX-LOGGER-01: Usa System.Threading.Mutex para serializar escrituras.
        La ruta del log se lee desde $Context.Paths.Logs (fuente autoritativa
        unica derivada de $Script:EngineRoot en New-TaskContext).
        Fallback: si el mutex no puede obtenerse en 5 s, se escribe sin el
        mutex para no bloquear la tarea — preferimos un log ligeramente
        desincronizado a una tarea bloqueada.
    #>
    param(
        [hashtable]$Context,
        [ValidateSet('DEBUG','INFO','WARN','ERROR')]
        [string]$Level = 'INFO',
        [string]$Message,
        [string]$Detail = ''
    )

    # Filtro de nivel — lee del contexto (copia del config, mutable por tarea)
    $levelMap = @{ DEBUG = 0; INFO = 1; WARN = 2; ERROR = 3 }
    $cfgLevel = $Context.Config.engine.logLevel
    if ($levelMap[$Level] -lt $levelMap[$cfgLevel]) { return }

    $ts = Get-Date -Format 'yyyy-MM-ddTHH:mm:ss'

    # Salida en consola
    $color = switch ($Level) {
        'DEBUG' { 'Gray'   }
        'INFO'  { 'Cyan'   }
        'WARN'  { 'Yellow' }
        'ERROR' { 'Red'    }
    }
    Write-Host "[$ts][$Level] $Message" -ForegroundColor $color

    # Construir JSON — solo ASCII
    $entry = [ordered]@{
        ts      = $ts
        level   = $Level
        runId   = $Context.RunId
        task    = $Context.TaskName
        message = $Message
        detail  = $Detail
    }
    $json = $entry | ConvertTo-Json -Compress

    # FIX-CONFIG-01: leer ruta del log desde Paths.Logs, no desde Config.paths.logs
    $regFile = Join-Path $Context.Paths.Logs 'registry.jsonl'

    _Write-LogLine -FilePath $regFile -Line $json
}

function Write-RunResult {
    <#
    .SYNOPSIS
        Escribe el resultado final de una ejecucion de tarea en registry.jsonl.
    #>
    param(
        [hashtable]$Context,
        [bool]$Success,
        [string]$ErrorMsg = ''
    )

    $ts      = Get-Date -Format 'yyyy-MM-ddTHH:mm:ss'
    $status  = if ($Success) { 'OK' } else { 'ERROR' }
    $elapsed = ([datetime]::Now - $Context.StartTime).TotalSeconds

    $entry = [ordered]@{
        ts      = $ts
        level   = $status
        runId   = $Context.RunId
        task    = $Context.TaskName
        message = "Run $status"
        detail  = $ErrorMsg
        elapsed = $elapsed
    }
    $json = $entry | ConvertTo-Json -Compress

    $regFile = Join-Path $Context.Paths.Logs 'registry.jsonl'

    _Write-LogLine -FilePath $regFile -Line $json
}

function _Write-LogLine {
    <#
    .SYNOPSIS
        Escribe una linea en el archivo indicado usando un Mutex nombrado
        para garantizar atomicidad bajo concurrencia.
    .NOTES
        Funcion interna — no invocar directamente desde tareas.
        El nombre del mutex usa el prefijo 'Global\' para que sea visible
        entre sesiones de PS distintas en el mismo equipo (requisito para
        Build-Parallel).
        Timeout de adquisicion: 5000 ms. Si expira se escribe sin mutex
        (fallback) para evitar bloquear la tarea indefinidamente.
    #>
    param(
        [string]$FilePath,
        [string]$Line
    )

    $mutex  = $null
    $locked = $false
    try {
        $mutex  = New-Object System.Threading.Mutex($false, 'Global\AutoBuildLogMutex')
        $locked = $mutex.WaitOne(5000)
        Add-Content -Path $FilePath -Value $Line -Encoding ASCII
    } catch {
        # Fallback: intentar escribir sin mutex antes de perder la entrada
        try { Add-Content -Path $FilePath -Value $Line -Encoding ASCII } catch { }
    } finally {
        if ($locked -and $null -ne $mutex) {
            try { $mutex.ReleaseMutex() } catch { }
        }
        if ($null -ne $mutex) {
            try { $mutex.Dispose() } catch { }
        }
    }
}
