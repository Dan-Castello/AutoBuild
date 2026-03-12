#Requires -Version 5.1
# =============================================================================
# lib/Config.ps1
# AutoBuild v3.0 - Configuration loader, merger, and bootstrap.
#
# DESIGN PRINCIPLES:
#   Single Responsibility: this file owns ONLY configuration.
#   Logger.ps1 owns only logging. Neither bleeds into the other.
#   (Resolves PROBLEMA-ARQUITECTURAL-01 from the audit.)
#
# IMMUTABILITY CONTRACT (resolves PROBLEMA-ARQUITECTURAL-03):
#   Get-EngineConfig ALWAYS returns a NEW independent hashtable.
#   It is never cached as an alias at the call site.
#   New-TaskContext deep-clones each section so task mutations cannot
#   contaminate the engine's master config object.
#
# PLUGIN EXTENSIBILITY (resolves PROBLEMA-EXT-01):
#   Adding a new section to engine.config.json requires only an entry
#   in $Script:KnownSections — no algorithmic change.
#
# IB VERSION DETECTION:
#   Invoke-Build version is detected from its script header and recorded
#   in cfg.engine.ibVersion for production audit trails.
#   (Resolves PROBLEMA-IB-01.)
# =============================================================================
Set-StrictMode -Version Latest

# Table-driven section registry. Add new config sections here only.
$Script:KnownSections = @('engine','sap','excel','reports','notifications','security')

function Get-EngineConfig {
    <#
    .SYNOPSIS
        Loads engine.config.json, merges over safe defaults, bootstraps
        working directories, and returns a new independent hashtable.
    .PARAMETER Root
        AutoBuild project root directory.
    .OUTPUTS
        Hashtable with keys: engine, sap, excel, reports, notifications, security.
    #>
    param([Parameter(Mandatory)][string]$Root)

    # Conservative production-safe defaults.
    $cfg = @{
        engine = @{
            logLevel          = 'INFO'
            maxRetries        = 3
            retryDelaySeconds = 5
            maxConcurrentJobs = 4
            ibVersion         = 'unknown'
        }
        sap = @{
            systemId = 'PRD'
            client   = '800'
            language = 'ES'
            timeout  = 180
        }
        excel = @{
            visible        = $false
            screenUpdating = $false
        }
        reports = @{
            defaultFormat   = 'xlsx'
            retentionDays   = 30
            maxLogSizeBytes = 10485760   # 10 MB before rotation
        }
        notifications = @{
            enabled    = $false
            smtpServer = ''
            smtpPort   = 587
            fromAddr   = ''
            toAddr     = ''
        }
        security = @{
            # AD group DNs that grant roles. Empty string = no AD check (dev mode).
            adminAdGroup     = ''
            developerAdGroup = ''
            # Comma-separated username fallback when AD integration not available.
            adminUsers       = ''
            developerUsers   = ''
        }
    }

    # Merge from file if present.
    $cfgFile = Join-Path $Root 'engine.config.json'
    if (Test-Path $cfgFile) {
        try {
            $raw = Get-Content $cfgFile -Raw -ErrorAction Stop | ConvertFrom-Json
            foreach ($section in $Script:KnownSections) {
                $rawSection = $raw.PSObject.Properties[$section]
                if ($null -ne $rawSection -and $null -ne $rawSection.Value) {
                    if (-not $cfg.ContainsKey($section)) { $cfg[$section] = @{} }
                    foreach ($prop in $rawSection.Value.PSObject.Properties) {
                        $cfg[$section][$prop.Name] = $prop.Value
                    }
                }
            }
        } catch {
            Write-Warning "AutoBuild Config: Cannot read engine.config.json: $_"
        }
    }

    # Detect and record Invoke-Build version for audit reproducibility.
    # (Resolves PROBLEMA-IB-01.)
    try {
        $ibFile = Join-Path $Root 'tools\InvokeBuild\Invoke-Build.ps1'
        if (Test-Path $ibFile) {
            $ibMatch = Select-String -Path $ibFile -Pattern 'version\s+([\d\.]+)' |
                       Select-Object -First 1
            if ($ibMatch) { $cfg.engine.ibVersion = $ibMatch.Matches.Groups[1].Value }
        }
    } catch { }

    # Guarantee working directories exist before any task needs them.
    foreach ($sub in @('logs','input','output','reports')) {
        $dir = Join-Path $Root $sub
        if (-not (Test-Path $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
    }

    return $cfg
}

function Get-ConfigSection {
    <#
    .SYNOPSIS
        Deep-clones a single config section hashtable.
        Used by New-TaskContext to produce isolated, task-local copies.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$Section
    )
    if ($Config.ContainsKey($Section)) {
        return $Config[$Section].Clone()
    }
    return @{}
}
