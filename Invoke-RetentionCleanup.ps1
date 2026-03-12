#Requires -Version 5.1
<#
.SYNOPSIS
    AutoBuild Artifact Retention Cleanup
.DESCRIPTION
    Deletes output and report files older than retentionDays as configured in engine.config.json.
    Can be scheduled via Task Scheduler or run manually from the UI.
    All deletions are audit-logged to logs/audit.jsonl.
.NOTES
    ASCII-only. PS 5.1.
.EXAMPLE
    .\Invoke-RetentionCleanup.ps1 -EnginePath "C:\AutoBuild"
    .\Invoke-RetentionCleanup.ps1 -EnginePath "C:\AutoBuild" -WhatIf
#>
param(
    [Parameter(Mandatory=$true)]
    [string]$EnginePath,
    [switch]$WhatIf
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Read config
$cfgFile = Join-Path $EnginePath 'engine.config.json'
$retentionDays = 30  # default
if (Test-Path $cfgFile) {
    try {
        $cfg = Get-Content $cfgFile -Raw -Encoding ASCII | ConvertFrom-Json
        if ($cfg.reports.retentionDays) { $retentionDays = [int]$cfg.reports.retentionDays }
    } catch {}
}

$cutoff    = (Get-Date).AddDays(-$retentionDays)
$outputDir = Join-Path $EnginePath 'output'
$reportDir = Join-Path $EnginePath 'reports'
$auditFile = Join-Path $EnginePath 'logs\audit.jsonl'

$deleted = 0
$skipped = 0

foreach ($dir in @($outputDir, $reportDir)) {
    if (-not (Test-Path $dir)) { continue }
    $files = Get-ChildItem -Path $dir -File | Where-Object { $_.LastWriteTime -lt $cutoff }
    foreach ($f in $files) {
        if ($WhatIf) {
            Write-Host "[WhatIf] Would delete: $($f.FullName)" -ForegroundColor Yellow
            $skipped++
        } else {
            try {
                Remove-Item -Path $f.FullName -Force
                $entry = @{
                    ts     = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss')
                    user   = $env:USERNAME
                    role   = 'RETENTION'
                    action = 'DELETE_ARTIFACT'
                    target = $f.Name
                    detail = "RetentionCleanup cutoff=$($cutoff.ToString('yyyy-MM-dd'))"
                } | ConvertTo-Json -Compress
                Add-Content -Path $auditFile -Value $entry -Encoding ASCII
                Write-Host "Deleted: $($f.Name)" -ForegroundColor Gray
                $deleted++
            } catch {
                Write-Warning "Could not delete $($f.Name): $_"
            }
        }
    }
}

Write-Host "Retention cleanup complete. Deleted=$deleted$(if ($WhatIf) {" (WhatIf, skipped=$skipped)"})" -ForegroundColor Cyan
