#Requires -Version 5.1
<#
.SYNOPSIS
    AutoBuild v3.0 - Single entry point for all task executions.

.DESCRIPTION
    Launches Invoke-Build tasks with a generic, extensible parameter model.
    Task-specific parameters are passed as a JSON dictionary via -Params.
    Adding new task parameters never requires modifying Run.ps1.

.NOTES
    AUDIT RESOLUTIONS:

    RUN-01 (HIGH) - Generic parameters:
        Replaced hardcoded $Centro/$Almacen/$Fecha/$Extra with a single
        -Params JSON string. Tasks accept any parameters without touching
        Run.ps1, AutoBuild.UI.ps1, or QueueRunner.psm1.

    RUN-02 (MED) - Checkpoint collision prevention:
        Reserved keys (Task, File, Checkpoint, Resume) are explicitly
        protected. A warning is emitted and the key is skipped rather
        than silently overwritten.

    RUN-03 (LOW) - Exit code guard:
        $LASTEXITCODE is null-checked. If null (Invoke-Build threw before
        setting it), exit code is derived from $? instead.

    CHK-01 (HIGH) - Timestamped checkpoint files:
        Checkpoint filename includes a timestamp so re-running the same
        task does not overwrite an in-progress checkpoint file.

.EXAMPLE
    .\Run.ps1 -List
    .\Run.ps1 sap_stock
    .\Run.ps1 sap_stock -Params '{"Centro":"1000","Almacen":"WH01"}'
    .\Run.ps1 sap_stock -Params '{"Centro":"1000"}' -Checkpoint
    .\Run.ps1 sap_stock -Checkpoint -Resume -CheckpointFile logs\checkpoint_sap_stock_20260312_143000.clixml
    .\Run.ps1 -WhatIf excel_reporte -Params '{"Fecha":"2026-03"}'
#>
param(
    [Parameter(Position = 0)]
    [string]$Task = '.',

    [switch]$List,
    [switch]$Checkpoint,
    [switch]$Resume,
    [switch]$WhatIf,

    # RUN-01 fix: all task parameters in one JSON dictionary.
    # Example: -Params '{"Centro":"1000","Almacen":"WH01","Fecha":"2026-03-12"}'
    [string]$Params = '{}',

    # CHK-01 fix: explicit checkpoint file allows resuming a specific run.
    # Auto-generated (timestamped) if not provided during -Checkpoint.
    [string]$CheckpointFile = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---- Paths -----------------------------------------------------------------
$Root  = $PSScriptRoot
$IB    = Join-Path $Root 'tools\InvokeBuild\Invoke-Build.ps1'
$IBChk = Join-Path $Root 'tools\InvokeBuild\Build-Checkpoint.ps1'
$Build = Join-Path $Root 'engine\Main.build.ps1'

# ---- Pre-flight validation -------------------------------------------------
if (-not (Test-Path $IB)) {
    Write-Host 'ERROR: tools\InvokeBuild\Invoke-Build.ps1 not found.' -ForegroundColor Red
    exit 1
}
if (-not (Test-Path $Build)) {
    Write-Host 'ERROR: engine\Main.build.ps1 not found.' -ForegroundColor Red
    exit 1
}

# ---- List mode -------------------------------------------------------------
if ($List) {
    & $IB -File $Build -Task ?
    exit 0
}

# ---- Deserialize task parameters -------------------------------------------
# RUN-01 fix: reserved keys cannot be overridden by task parameters.
$Script:ReservedKeys = @('Task','File','Checkpoint','Resume','_ParamsJson')

$TaskParams = @{}
if (-not [string]::IsNullOrWhiteSpace($Params) -and $Params -ne '{}') {
    try {
        $parsed = $Params | ConvertFrom-Json
        foreach ($prop in $parsed.PSObject.Properties) {
            if ($prop.Name -in $Script:ReservedKeys) {
                Write-Warning "Run.ps1: '$($prop.Name)' is a reserved key and was ignored."
                continue
            }
            $TaskParams[$prop.Name] = $prop.Value
        }
    } catch {
        Write-Host "ERROR: -Params is not valid JSON: $_" -ForegroundColor Red
        exit 1
    }
}

# ---- Checkpoint / Resume mode ----------------------------------------------
if ($Checkpoint -or $Resume) {
    if ([string]::IsNullOrWhiteSpace($CheckpointFile)) {
        # CHK-01 fix: timestamp prevents overwriting an in-progress checkpoint.
        $stamp         = Get-Date -Format 'yyyyMMdd_HHmmss'
        $CheckpointFile = Join-Path $Root "logs\checkpoint_${Task}_${stamp}.clixml"
    }

    $chkArgs = @{
        Checkpoint = $CheckpointFile
        Build      = @{ Task = $Task; File = $Build }
    }

    # RUN-02 fix: explicit collision detection for reserved keys.
    foreach ($k in $TaskParams.Keys) {
        if ($k -notin $Script:ReservedKeys) {
            $chkArgs.Build[$k] = $TaskParams[$k]
        }
    }

    if ($Resume) { $chkArgs['Resume'] = $true }

    & $IBChk @chkArgs

    # RUN-03 fix: null-safe exit code.
    $exitCode = if ($null -ne $LASTEXITCODE) { $LASTEXITCODE } else { if ($?) { 0 } else { 1 } }
    exit $exitCode
}

# ---- Normal mode -----------------------------------------------------------
$invokeArgs = @{
    Task   = $Task
    File   = $Build
    WhatIf = $WhatIf
}
if ($TaskParams.Count -gt 0) { $invokeArgs += $TaskParams }

& $IB @invokeArgs

# RUN-03 fix: null-safe exit code.
$exitCode = if ($null -ne $LASTEXITCODE) { $LASTEXITCODE } else { if ($?) { 0 } else { 1 } }
exit $exitCode
