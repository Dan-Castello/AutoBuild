#Requires -Version 5.1
<#
.SYNOPSIS
    AutoBuild v3.0 - Test runner helper.
.DESCRIPTION
    Installs Pester if needed, then runs the full test suite.
.EXAMPLE
    .\Run-Tests.ps1
    .\Run-Tests.ps1 -Tag 'Config','Logger'
    .\Run-Tests.ps1 -CI       # CI mode: exits with non-zero on failures
#>
param(
    [string[]]$Tag = @(),
    [switch]$CI
)

# Install Pester 5 if not present
$pester = Get-Module -ListAvailable -Name Pester | Where-Object { $_.Version.Major -ge 5 } | Select-Object -First 1
if ($null -eq $pester) {
    Write-Host 'Pester 5 not found. Installing...' -ForegroundColor Yellow
    Install-Module Pester -Force -SkipPublisherCheck -MinimumVersion 5.0.0 -Scope CurrentUser
}
Import-Module Pester -MinimumVersion 5.0.0

$config = New-PesterConfiguration
$config.Run.Path             = Join-Path $PSScriptRoot 'tests\AutoBuild.Tests.ps1'
$config.Output.Verbosity     = 'Detailed'
$config.TestResult.Enabled   = $true
$config.TestResult.OutputPath = Join-Path $PSScriptRoot 'tests\TestResults.xml'

if ($Tag.Count -gt 0) {
    $config.Filter.Tag = $Tag
}

$result = Invoke-Pester -Configuration $config

if ($CI -and $result.FailedCount -gt 0) {
    Write-Host "`nCI mode: $($result.FailedCount) test(s) failed. Exiting 1." -ForegroundColor Red
    exit 1
}
