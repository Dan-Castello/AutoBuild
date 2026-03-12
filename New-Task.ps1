#Requires -Version 5.1
<#
.SYNOPSIS
    Genera un nuevo archivo de tarea a partir del template.
.EXAMPLE
    .\New-Task.ps1 -Name sap_ventas -Category SAP -Description "Ventas por periodo"
#>
param(
    [Parameter(Mandatory=$true)]
    [ValidatePattern('^[a-zA-Z0-9_-]+$')]
    [string]$Name,

    [ValidateSet('SAP','Excel','CSV','Reporte','Utilidad')]
    [string]$Category = 'Utilidad',

    [string]$Description = 'Descripcion pendiente'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$Root     = $PSScriptRoot
$Template = Join-Path $Root 'tasks\task_TEMPLATE.ps1'
$Output   = Join-Path $Root "tasks\task_${Name}.ps1"

if (-not (Test-Path $Template)) {
    Write-Host "ERROR: Template no encontrado: $Template" -ForegroundColor Red
    exit 1
}

if (Test-Path $Output) {
    Write-Host "ERROR: La tarea ya existe: $Output" -ForegroundColor Red
    exit 1
}

$content = Get-Content $Template -Raw -Encoding ASCII
$content = $content `
    -replace 'task_NOMBRE\.ps1', "task_${Name}.ps1" `
    -replace '@Description : Descripcion breve de la tarea', "@Description : $Description" `
    -replace '@Category    : SAP \| Excel \| CSV \| Reporte \| Utilidad', "@Category    : $Category" `
    -replace '# Synopsis: Descripcion breve que aparece en \.\\Run\.ps1 -List', "# Synopsis: $Description" `
    -replace 'task NOMBRE \{', "task $Name {" `
    -replace "TaskName 'NOMBRE'", "TaskName '$Name'"

# Forzar ASCII
[System.IO.File]::WriteAllText($Output, $content, [System.Text.Encoding]::ASCII)

Write-Host "Tarea creada: $Output" -ForegroundColor Green
Write-Host "Editar e implementar la logica dentro de la funcion task $Name { }" -ForegroundColor Cyan
