#Requires -Version 5.1
<#
.SYNOPSIS
    AutoBuild - Configuracion inicial del proyecto (ejecutar una sola vez)
.DESCRIPTION
    1. Verifica que PS sea 5.1
    2. Crea carpetas de trabajo si no existen
    3. Verifica que Invoke-Build este en tools/InvokeBuild/
    4. Verifica disponibilidad de Excel y SAP GUI (opcional)
    5. Configura politica de ejecucion para el usuario actual
#>
param(
    [switch]$SkipExecPolicy,
    [switch]$SkipComCheck
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$Root = $PSScriptRoot

Write-Host ''
Write-Host '===== AutoBuild Setup =====' -ForegroundColor Cyan

# ---- PowerShell version --------------------------------------------------
$psv = $PSVersionTable.PSVersion
Write-Host "PS Version   : $($psv.Major).$($psv.Minor)" -NoNewline
if ($psv.Major -eq 5) {
    Write-Host ' [OK]' -ForegroundColor Green
} else {
    Write-Host ' [WARN] Se recomienda PowerShell 5.1' -ForegroundColor Yellow
}

# ---- Politica de ejecucion -----------------------------------------------
if (-not $SkipExecPolicy) {
    $pol = Get-ExecutionPolicy -Scope CurrentUser
    if ($pol -notin @('RemoteSigned','Unrestricted','Bypass')) {
        Write-Host 'Estableciendo ExecutionPolicy RemoteSigned para el usuario...' -ForegroundColor Yellow
        Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force
        Write-Host '  ExecutionPolicy : RemoteSigned [OK]' -ForegroundColor Green
    } else {
        Write-Host "ExecutionPolicy : $pol [OK]" -ForegroundColor Green
    }
}

# ---- Carpetas de trabajo -------------------------------------------------
$dirs = @('input','output','reports','logs','tasks')
foreach ($dir in $dirs) {
    $p = Join-Path $Root $dir
    if (-not (Test-Path $p)) {
        New-Item -ItemType Directory -Path $p -Force | Out-Null
        Write-Host "  Creada carpeta : $dir" -ForegroundColor Green
    } else {
        Write-Host "  Carpeta OK     : $dir" -ForegroundColor Green
    }
}

# ---- Invoke-Build portable -----------------------------------------------
$ibFile = Join-Path $Root 'tools\InvokeBuild\Invoke-Build.ps1'
if (Test-Path $ibFile) {
    Write-Host 'Invoke-Build   : OK (portable)' -ForegroundColor Green
} else {
    Write-Host 'WARN: tools\InvokeBuild\Invoke-Build.ps1 no encontrado.' -ForegroundColor Red
    Write-Host '      Copiar manualmente desde el zip de Invoke-Build.' -ForegroundColor Red
}

# ---- Comprobacion COM (opcional) -----------------------------------------
if (-not $SkipComCheck) {
    Write-Host ''
    Write-Host 'Comprobando disponibilidad COM (puede tardar hasta 30s)...' -ForegroundColor Cyan

    foreach ($prog in @('Excel.Application', 'SapROTWr.SapROTWrapper')) {
        $job = $null
        try {
            $job = Start-Job -ScriptBlock {
                param($p)
                try {
                    $o = New-Object -ComObject $p -ErrorAction Stop
                    [Runtime.InteropServices.Marshal]::ReleaseComObject($o) | Out-Null
                    return $true
                } catch { return $false }
            } -ArgumentList $prog

            $done = Wait-Job -Job $job -Timeout 15
            $ok   = if ($null -ne $done) { Receive-Job -Job $job } else { $false }

            if ($ok) {
                Write-Host "  $prog : disponible" -ForegroundColor Green
            } else {
                Write-Host "  $prog : NO disponible (normal si no esta abierto)" -ForegroundColor Yellow
            }
        } finally {
            if ($null -ne $job) { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue }
        }
    }
}

Write-Host ''
Write-Host '===== Setup completado =====' -ForegroundColor Cyan
Write-Host "Ejecutar tareas : .\Run.ps1 -List" -ForegroundColor White
Write-Host "Diagnostico     : .\Run.ps1 diagnostico" -ForegroundColor White
Write-Host ''
