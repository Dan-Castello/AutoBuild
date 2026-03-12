#Requires -Version 5.1
<#
.SYNOPSIS
    AutoBuild v3.0 - Scaffold a new task file from the v3 template.
.EXAMPLE
    .\New-Task.ps1 -Name sap_stock -Category SAP -Description "SAP stock report"
#>
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^[a-zA-Z0-9_-]+$')]
    [string]$Name,

    [ValidateSet('SAP','Excel','CSV','Report','Utility')]
    [string]$Category = 'Utility',

    [string]$Description = 'Description pending',
    [string]$Author = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$Root     = $PSScriptRoot
$Template = Join-Path $Root 'tasks\task_TEMPLATE.ps1'
$Output   = Join-Path $Root "tasks\task_${Name}.ps1"

if (-not (Test-Path $Template)) { Write-Host "ERROR: Template not found: $Template" -ForegroundColor Red; exit 1 }
if (Test-Path $Output)          { Write-Host "ERROR: Task already exists: $Output"   -ForegroundColor Red; exit 1 }

if ([string]::IsNullOrWhiteSpace($Author)) {
    try { $Author = ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name -split '\\')[-1] }
    catch { $Author = $env:USERNAME }
}

$content = Get-Content $Template -Raw -Encoding ASCII
$content = $content `
    -replace '@Description : Brief task description',                         "@Description : $Description" `
    -replace '@Category    : SAP \| Excel \| CSV \| Report \| Utility',       "@Category    : $Category" `
    -replace '@Author      : Your Name',                                       "@Author      : $Author" `
    -replace '# Synopsis: Brief description shown in .\\Run.ps1 -List',       "# Synopsis: $Description" `
    -replace 'task NOMBRE \{',                                                 "task $Name {" `
    -replace "-TaskName 'NOMBRE'",                                             "-TaskName '$Name'"

[System.IO.File]::WriteAllText($Output, $content, [System.Text.Encoding]::ASCII)

Write-Host "Task created: $Output" -ForegroundColor Green
Write-Host "  Edit and implement: task $Name { }" -ForegroundColor Cyan
Write-Host "  Run with: .\Run.ps1 $Name -Params '{""key"":""value""}'" -ForegroundColor White

$hashFile = Join-Path $Root 'tasks\tasks.hash.json'
if (Test-Path $hashFile) {
    Write-Host "  Register hash: . .\lib\Integrity.ps1; Register-TaskHash -HashFile tasks\tasks.hash.json -FilePath `"$Output`"" -ForegroundColor Yellow
}
