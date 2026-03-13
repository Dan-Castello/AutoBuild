<#
.SYNOPSIS
    AutoBuild UI Quick Diagnostic
.DESCRIPTION
    Checks engine files, libraries, folders, permissions, and COM availability
    to diagnose why AutoBuild.UI.ps1 exits with error 2.
#>

param(
    [string]$EnginePath = "$PSScriptRoot\..",
    [string]$Role = 'Operator'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

$Report = [System.Collections.Generic.List[PSCustomObject]]::new()

function Add-Report {
    param($Category, $Item, $Status, $Message)
    $Report.Add([PSCustomObject]@{
        Category = $Category
        Item     = $Item
        Status   = $Status
        Message  = $Message
    })
}

# Paths
$EngineRoot  = Resolve-Path $EnginePath -ErrorAction SilentlyContinue
$UIRoot      = "$EngineRoot\ui"
$RunScript   = Join-Path $EngineRoot 'Run.ps1'
$ConfigFile  = Join-Path $EngineRoot 'engine.config.json'
$Libs        = @('Config.ps1','Logger.ps1','Context.ps1','Auth.ps1')

# Check Engine Root
if (-not (Test-Path $EngineRoot)) { Add-Report 'Paths' 'EngineRoot' 'ERROR' "Path not found: $EngineRoot" } else { Add-Report 'Paths' 'EngineRoot' 'OK' "Exists" }

# Check Run.ps1
if (-not (Test-Path $RunScript)) { Add-Report 'Engine' 'Run.ps1' 'ERROR' "Missing" } else { Add-Report 'Engine' 'Run.ps1' 'OK' "Found" }

# Check Config
if (-not (Test-Path $ConfigFile)) { Add-Report 'Engine' 'engine.config.json' 'ERROR' "Missing" } else {
    try {
        $cfg = Get-Content $ConfigFile -Raw -Encoding ASCII | ConvertFrom-Json
        Add-Report 'Engine' 'engine.config.json' 'OK' "Valid JSON"
    } catch { Add-Report 'Engine' 'engine.config.json' 'ERROR' "Invalid JSON: $_" }
}

# Check Libraries
$LibPath = Join-Path $EngineRoot 'lib'
foreach ($lib in $Libs) {
    $lp = Join-Path $LibPath $lib
    if (-not (Test-Path $lp)) { Add-Report 'Libraries' $lib 'ERROR' "Missing: $lp" } else { Add-Report 'Libraries' $lib 'OK' "Found" }
}

# Check UI XAML
$XamlFile = Join-Path $UIRoot 'AutoBuild.xaml'
if (-not (Test-Path $XamlFile)) { Add-Report 'UI' 'AutoBuild.xaml' 'WARN' "Missing, inline fallback used" } else { Add-Report 'UI' 'AutoBuild.xaml' 'OK' "Found" }

# Check folder permissions
$Dirs = @('logs','output','reports','input')
foreach ($d in $Dirs) {
    $path = Join-Path $EngineRoot $d
    if (-not (Test-Path $path)) { Add-Report 'Folders' $d 'WARN' "Missing" } else {
        try {
            $tmp = Join-Path $path "._test_$(Get-Random)"
            [void](New-Item $tmp -ItemType File -Force)
            Remove-Item $tmp -Force
            Add-Report 'Folders' $d 'OK' "Writable"
        } catch { Add-Report 'Folders' $d 'WARN' "Exists but not writable" }
    }
}

# Check COM objects
$COMs = @(
    @{Name='Excel.Application'; Friendly='Microsoft Excel'},
    @{Name='Word.Application';  Friendly='Microsoft Word'},
    @{Name='SapROTWr.SapROTWrapper'; Friendly='SAP GUI'}
)

foreach ($c in $COMs) {
    try {
        $obj = New-Object -ComObject $c.Name -ErrorAction Stop
        if ($c.Name -eq 'Excel.Application' -or $c.Name -eq 'Word.Application') { $ver = $obj.Version; $obj.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null; Add-Report 'COM' $c.Friendly 'OK' "v$ver available" }
        else { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null; Add-Report 'COM' $c.Friendly 'OK' "Available" }
    } catch { Add-Report 'COM' $c.Friendly 'WARN' "Not available: $_" }
}

# Output report
$Report | Format-Table -AutoSize

# Optional: export to CSV
#$Report | Export-Csv -Path "$PSScriptRoot\AutoBuild_UI_Diagnostic.csv" -NoTypeInformation -Encoding UTF8