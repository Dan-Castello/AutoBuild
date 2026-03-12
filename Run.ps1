#Requires -Version 5.1
<#
.SYNOPSIS
    AutoBuild - Lanzador de tareas de automatizacion
.DESCRIPTION
    Punto de entrada principal. Invoca Invoke-Build de forma portable
    sin necesidad de instalacion en el sistema.
.EXAMPLE
    .\Run.ps1 -List
    .\Run.ps1 sap_stock
    .\Run.ps1 sap_stock -Centro 1000
    .\Run.ps1 sap_stock -Checkpoint
#>
param(
    [Parameter(Position=0)]
    [string]$Task = '.',

    [switch]$List,
    [switch]$Checkpoint,
    [switch]$Resume,
    [switch]$WhatIf,

    # Parametros de tarea (pasados al build script)
    [string]$Centro,
    [string]$Almacen,
    [string]$Fecha,
    [string]$Extra
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---- Rutas base ------------------------------------------------------------
$Root   = $PSScriptRoot
$IB     = Join-Path $Root 'tools\InvokeBuild\Invoke-Build.ps1'
$IBChk  = Join-Path $Root 'tools\InvokeBuild\Build-Checkpoint.ps1'
$Build  = Join-Path $Root 'engine\Main.build.ps1'

# ---- Validaciones previas --------------------------------------------------
if (-not (Test-Path $IB)) {
    Write-Host 'ERROR: tools\InvokeBuild\Invoke-Build.ps1 no encontrado.' -ForegroundColor Red
    exit 1
}
if (-not (Test-Path $Build)) {
    Write-Host 'ERROR: engine\Main.build.ps1 no encontrado.' -ForegroundColor Red
    exit 1
}

# ---- Modo lista ------------------------------------------------------------
if ($List) {
    & $IB -File $Build -Task ?
    exit 0
}

# ---- Parametros del build script -------------------------------------------
$BuildParams = @{}
if ($Centro)  { $BuildParams['Centro']  = $Centro }
if ($Almacen) { $BuildParams['Almacen'] = $Almacen }
if ($Fecha)   { $BuildParams['Fecha']   = $Fecha }
if ($Extra)   { $BuildParams['Extra']   = $Extra }

# ---- Modo checkpoint/resume ------------------------------------------------
if ($Checkpoint -or $Resume) {
    $ChkFile = Join-Path $Root "logs\checkpoint_${Task}.clixml"
    $ChkArgs = @{
        Checkpoint = $ChkFile
        Build      = (@{ Task = $Task; File = $Build } + $BuildParams)
    }
    if ($Resume) { $ChkArgs['Resume'] = $true }
    & $IBChk @ChkArgs
    exit $LASTEXITCODE
}

# ---- Modo normal -----------------------------------------------------------
$InvokeArgs = @{
    Task    = $Task
    File    = $Build
    WhatIf  = $WhatIf
}
if ($BuildParams.Count -gt 0) { $InvokeArgs += $BuildParams }

& $IB @InvokeArgs
exit $LASTEXITCODE
