#Requires -Version 5.1
# lib/Context.ps1
# Fabrica de contextos de ejecucion para tareas.
# El contexto es el objeto que se pasa entre funciones de libreria.
#
# FIX-CONFIG-01 (Inmutabilidad de configuracion):
#   New-TaskContext recibe $Script:EngineConfig (hashtable autoritativo)
#   y crea una COPIA de las secciones que necesita. Las mutaciones del
#   contexto (ej. $ctx.Config.engine.logLevel = 'DEBUG' en diag_log)
#   afectan solo la copia local de la tarea y no contaminan
#   $Script:EngineConfig ni los contextos de otras tareas.
#
# FIX-PATH-01 (Ruta autoritativa unica):
#   $ctx.Paths.* se calcula EXCLUSIVAMENTE a partir del parametro $Root,
#   que en engine/Main.build.ps1 siempre es $Script:EngineRoot.
#   Se elimina la dependencia de $Config.paths.* que existia antes
#   y que podia divergir si el build se lanzaba desde un directorio
#   distinto al de la configuracion.
#
# CONTRATO DE COMPATIBILIDAD:
#   - La firma publica de New-TaskContext no cambia.
#   - Las claves de $ctx.Paths (Root, Input, Output, Reports, Logs)
#     mantienen los mismos nombres que antes.
#   - $ctx.Config sigue siendo accesible (es la copia, no el original).

Set-StrictMode -Version Latest

function New-TaskContext {
    <#
    .SYNOPSIS
        Crea el contexto de ejecucion de una tarea.
    .PARAMETER TaskName
        Nombre de la tarea (para logs y trazas).
    .PARAMETER Config
        Configuracion del motor ($Script:EngineConfig).
        New-TaskContext crea una copia defensiva — mutaciones no afectan
        al original.
    .PARAMETER Root
        Directorio raiz del proyecto AutoBuild ($Script:EngineRoot).
        Todas las rutas de trabajo se derivan de este parametro.
    .PARAMETER Params
        Parametros especificos de la tarea (opcional).
    #>
    param(
        [string]$TaskName,
        [hashtable]$Config,
        [string]$Root,
        [hashtable]$Params = @{}
    )

    # FIX-CONFIG-01: copia defensiva de las secciones usadas por la tarea.
    # Cada seccion se clona con .Clone() para obtener una nueva tabla hash
    # independiente. Las secciones anidadas (engine, excel, sap) contienen
    # solo valores escalares (string, bool, int), por lo que Clone() de nivel
    # 1 es suficiente — no hay hashtables anidadas adicionales en estos valores.
    $configSnapshot = @{
        engine  = $Config.engine.Clone()
        sap     = $Config.sap.Clone()
        excel   = $Config.excel.Clone()
        reports = $Config.reports.Clone()
    }

    # FIX-PATH-01: todas las rutas derivan de $Root.
    # Se elimina la lectura desde $Config.paths.* que era la fuente de
    # divergencia cuando Config y BuildRoot apuntaban a directorios distintos.
    $ctx = @{
        RunId     = New-RunId
        TaskName  = $TaskName
        Config    = $configSnapshot
        StartTime = [datetime]::Now
        Params    = $Params
        Paths     = @{
            Root    = $Root
            Input   = Join-Path $Root 'input'
            Output  = Join-Path $Root 'output'
            Reports = Join-Path $Root 'reports'
            Logs    = Join-Path $Root 'logs'
        }
    }

    return $ctx
}
