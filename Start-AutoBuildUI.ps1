#Requires -Version 5.1
<#
.SYNOPSIS
    AutoBuild UI Launcher with role selection.
.DESCRIPTION
    Ensures STA apartment model for WPF, validates prerequisites,
    and launches the AutoBuild Automation Interface.
.PARAMETER EnginePath
    Path to the AutoBuild engine root. Defaults to the same folder as this script.
.PARAMETER Role
    RBAC role: Operator (default), Developer, or Admin.
.EXAMPLE
    .\Start-AutoBuildUI.ps1
    .\Start-AutoBuildUI.ps1 -Role Admin
    .\Start-AutoBuildUI.ps1 -EnginePath "C:\AutoBuild" -Role Developer
#>
param(
    [string]$EnginePath = '',
    [ValidateSet('Operator','Developer','Admin')]
    [string]$Role = 'Operator'
)

$ErrorActionPreference = 'Stop'

# ---- Resolve engine path ----
$ScriptDir = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($EnginePath)) {
    $EnginePath = $ScriptDir
}

# ---- Validate PS 5.1 ----
if ($PSVersionTable.PSVersion.Major -ne 5 -or $PSVersionTable.PSVersion.Minor -ne 1) {
    Write-Warning "PowerShell 5.1 recommended for full COM support. Current: $($PSVersionTable.PSVersion)"
}

# ---- Validate .NET WPF assemblies ----
try {
    Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
} catch {
    Write-Host 'ERROR: WPF assemblies not available. Windows .NET Framework 4.x required.' -ForegroundColor Red
    exit 1
}

# ---- Ensure STA thread (required for WPF) ----
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    Write-Host 'Relaunching in STA mode for WPF compatibility...' -ForegroundColor Cyan
    $scriptPath = Join-Path $ScriptDir 'AutoBuild.UI.ps1'
    Start-Process powershell.exe -ArgumentList @(
        '-NoProfile', '-NonInteractive',
        '-ExecutionPolicy', 'Bypass',
        '-STA',
        '-File', "`"$scriptPath`"",
        '-EnginePath', "`"$EnginePath`"",
        '-Role', $Role
    ) -Wait
    exit $LASTEXITCODE
}

# ---- Launch UI ----
$uiScript = Join-Path $ScriptDir 'AutoBuild.UI.ps1'
if (-not (Test-Path $uiScript)) {
    Write-Host "ERROR: AutoBuild.UI.ps1 not found at: $uiScript" -ForegroundColor Red
    exit 1
}

Write-Host "AutoBuild Automation Interface" -ForegroundColor Cyan
Write-Host "  Engine: $EnginePath" -ForegroundColor Gray
Write-Host "  Role:   $Role" -ForegroundColor Gray
Write-Host "  User:   $env:USERNAME" -ForegroundColor Gray
Write-Host ''

. $uiScript -EnginePath $EnginePath -Role $Role
