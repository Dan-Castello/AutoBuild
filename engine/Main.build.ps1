#Requires -Version 5.1
<#
.SYNOPSIS
    AutoBuild - Orquestador principal de tareas
    Motor basado en Invoke-Build 5.14.23 (portable, sin instalacion)
.NOTES
    Solo ASCII. PS 5.1. Sin dependencias externas.
    Entorno: Office 16, SAP GUI 800, Windows 10/11

    FIX-ROOT-01 (Ruta autoritativa unica):
        $Script:EngineRoot se establece UNA vez al inicio del script,
        antes de cargar libs o tareas. Es la unica fuente de verdad
        para rutas del proyecto. New-TaskContext recibe este valor
        como parametro -Root y deriva todas las rutas de trabajo
        desde el (input, output, reports, logs).
        Se elimina el patron '(Split-Path $BuildRoot -Parent)' que
        las tareas usaban localmente y que podia producir valores
        distintos si $BuildRoot cambiaba entre la carga y la ejecucion.

    FIX-CONFIG-02 (Configuracion inmutable en el motor):
        $Script:EngineConfig es el hashtable autoritativo de configuracion.
        Las tareas NUNCA reciben una referencia a este objeto.
        New-TaskContext crea una copia defensiva para cada tarea.
        Mutaciones en el contexto de la tarea (ej. $ctx.Config.engine.logLevel)
        afectan solo la copia local y no contaminan $Script:EngineConfig.

    FIX-LOAD-01 (Carga determinista de tareas):
        Get-ChildItem | Sort-Object -Property Name garantiza que los archivos
        de tareas se importan en orden alfabetico de nombre de archivo,
        independientemente del orden de enumeracion del sistema de archivos.
        Esto hace que la resolucion de dependencias entre tareas sea estable
        en todos los entornos.
#>

# ---- Parametros del script (visibles y persistibles en checkpoint) ---------
param(
    [string]$Centro  = '',
    [string]$Almacen = '',
    [string]$Fecha   = '',
    [string]$Extra   = ''
)

Set-StrictMode -Version Latest

# ============================================================================
# FIX-ROOT-01: Establecer la ruta raiz autoritativa del proyecto UNA sola vez.
# $BuildRoot apunta a engine/ (directorio del build script).
# $Script:EngineRoot apunta a la raiz del proyecto AutoBuild (un nivel arriba).
# Todas las libs, tareas y rutas de trabajo se derivan de $Script:EngineRoot.
# ============================================================================
$Script:EngineRoot = Split-Path $BuildRoot -Parent

# ---- Bootstrap: cargar librerias del motor ---------------------------------
$LibPath = Join-Path $Script:EngineRoot 'lib'

. (Join-Path $LibPath 'Logger.ps1')
. (Join-Path $LibPath 'Context.ps1')
. (Join-Path $LibPath 'ComHelper.ps1')
. (Join-Path $LibPath 'SapHelper.ps1')
. (Join-Path $LibPath 'ExcelHelper.ps1')
. (Join-Path $LibPath 'WordHelper.ps1')
. (Join-Path $LibPath 'Assertions.ps1')

# ============================================================================
# FIX-CONFIG-02: Cargar configuracion en la variable autoritativa del motor.
# $Script:EngineConfig es SOLO para lectura por parte del motor.
# Las tareas obtienen copias defensivas via New-TaskContext.
# ============================================================================
$Script:EngineConfig = Get-EngineConfig -Root $Script:EngineRoot

# ---- Alias de compatibilidad -----------------------------------------------
# $Script:Config se mantiene como alias de $Script:EngineConfig para que las
# tareas existentes que referencian '$Script:Config' continuen funcionando
# sin modificacion. Es una referencia al mismo objeto, por lo que si el motor
# necesitara reemplazar EngineConfig en el futuro, deberia actualizar ambas.
$Script:Config = $Script:EngineConfig

# ============================================================================
# FIX-LOAD-01: Carga determinista de tareas en orden alfabetico de nombre.
# Sort-Object -Property Name garantiza orden estable independientemente del
# sistema de archivos subyacente (NTFS, FAT, ReFS).
# ============================================================================
$TasksPath = Join-Path $Script:EngineRoot 'tasks'
if (Test-Path $TasksPath) {
    Get-ChildItem -Path $TasksPath -Filter 'task_*.ps1' |
        Sort-Object -Property Name |
        ForEach-Object { . $_.FullName }
}

# ---- Tareas integradas del motor -------------------------------------------

# Synopsis: Limpia procesos COM zombi (Excel/Word sin ventana)
task limpiar_com {
    $ctx = New-TaskContext `
        -TaskName 'limpiar_com' `
        -Config   $Script:EngineConfig `
        -Root     $Script:EngineRoot
    Write-BuildLog $ctx 'INFO' 'Buscando procesos COM zombi'
    $count = Remove-ZombieCom
    Write-Build Cyan "  Procesos eliminados: $count"
    Write-BuildLog $ctx 'INFO' "Limpieza COM completada. Eliminados: $count"
}

# Synopsis: Tarea vacia de ejemplo para probar el motor
task ejemplo {
    $ctx = New-TaskContext `
        -TaskName 'ejemplo' `
        -Config   $Script:EngineConfig `
        -Root     $Script:EngineRoot
    Write-BuildLog $ctx 'INFO' 'Tarea de ejemplo ejecutada'
    Write-Build Green '  Motor Invoke-Build funcionando correctamente'
}

# ---- Tarea por defecto -----------------------------------------------------
task . diag_completo
