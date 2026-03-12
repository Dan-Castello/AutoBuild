#Requires -Version 5.1
<#
.SYNOPSIS
    AutoBuild UI Task Execution Helper
.DESCRIPTION
    Called by AutoBuild.UI.ps1 to run tasks via Run.ps1 in a clean environment.
    Captures exit codes and writes structured output for UI consumption.
    Do NOT modify task logic here - all execution goes through Run.ps1.
.NOTES
    ASCII-only. PS 5.1. Internal to the UI layer.
    Parameters beyond TaskName are passed through as -Key Value pairs.
.EXAMPLE
    .\Invoke-UITask.ps1 -EnginePath "C:\AutoBuild" -TaskName sap_stock -Centro 1000
    .\Invoke-UITask.ps1 -EnginePath "C:\AutoBuild" -TaskName diag_motor -WhatIf
#>
param(
    [Parameter(Mandatory=$true)]
    [string]$EnginePath,

    [Parameter(Mandatory=$true)]
    [string]$TaskName,

    # Standard Run.ps1 parameters
    [string]$Centro  = '',
    [string]$Almacen = '',
    [string]$Fecha   = '',
    [string]$Extra   = '',

    [switch]$WhatIf,
    [switch]$Checkpoint,
    [switch]$Resume
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$RunScript = Join-Path $EnginePath 'Run.ps1'

if (-not (Test-Path $RunScript)) {
    Write-Error "Run.ps1 not found: $RunScript"
    exit 1
}

# Build argument list
$runArgs = @($TaskName)
if (-not [string]::IsNullOrWhiteSpace($Centro))  { $runArgs += '-Centro';  $runArgs += $Centro }
if (-not [string]::IsNullOrWhiteSpace($Almacen)) { $runArgs += '-Almacen'; $runArgs += $Almacen }
if (-not [string]::IsNullOrWhiteSpace($Fecha))   { $runArgs += '-Fecha';   $runArgs += $Fecha }
if (-not [string]::IsNullOrWhiteSpace($Extra))   { $runArgs += '-Extra';   $runArgs += $Extra }
if ($WhatIf)     { $runArgs += '-WhatIf' }
if ($Checkpoint) { $runArgs += '-Checkpoint' }
if ($Resume)     { $runArgs += '-Resume' }

Write-Host "[UI-HELPER] Executing: Run.ps1 $($runArgs -join ' ')" -ForegroundColor DarkCyan

& $RunScript @runArgs
$exitCode = $LASTEXITCODE

if ($exitCode -eq 0) {
    Write-Host "[UI-HELPER] Task completed successfully." -ForegroundColor Green
} else {
    Write-Host "[UI-HELPER] Task failed with exit code: $exitCode" -ForegroundColor Red
}

exit $exitCode
