#Requires -Version 5.1
# =============================================================================
# lib/Context.ps1
# AutoBuild v3.0 - Task execution context factory.
#
# AUDIT RESOLUTIONS:
#   FIX-CONFIG-01 (IMMUTABILITY): New-TaskContext creates a deep clone of
#     all config sections. Task mutations never contaminate $Script:EngineConfig.
#   FIX-PATH-01 (SINGLE ROOT): All Paths.* entries derive exclusively from
#     the $Root parameter. No reliance on config-embedded path strings.
#   F3-08 (ROBUSTNESS): Context now captures User (from WindowsIdentity),
#     Hostname, and SessionId for structured log enrichment.
#   PROBLEMA-ARQUITECTURAL-03: $Script:Config alias eliminated.
#     Tasks reference $Script:EngineConfig or their context copy only.
#
# CONTEXT SCHEMA:
#   RunId     - unique run identifier (collision-free GUID fragment)
#   TaskName  - name of the executing task
#   Config    - deep clone of relevant engine config sections (mutable by task)
#   StartTime - datetime of context creation
#   Params    - task-specific parameters from Run.ps1 -Params JSON
#   User      - Windows username (domain\user -> user only)
#   Hostname  - machine name for cross-host log correlation
#   SessionId - process ID for within-session correlation
#   Paths     - hashtable: Root, Input, Output, Reports, Logs
# =============================================================================
Set-StrictMode -Version Latest

function New-TaskContext {
    <#
    .SYNOPSIS
        Creates the execution context for a task.
    .PARAMETER TaskName
        Name of the task (appears in every log entry).
    .PARAMETER Config
        Engine master config ($Script:EngineConfig). A deep clone is made;
        the original is never mutated.
    .PARAMETER Root
        AutoBuild project root ($Script:EngineRoot). All paths derive from this.
    .PARAMETER Params
        Hashtable of task-specific parameters (from $Script:TaskParams).
    .OUTPUTS
        Hashtable representing the task execution context.
    #>
    param(
        [Parameter(Mandatory)][string]$TaskName,
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][string]$Root,
        [hashtable]$Params = @{}
    )

    # Deep clone of all config sections. Each section contains only scalar
    # values, so a single-level Clone() is sufficient. New config sections
    # added in Config.ps1 are automatically cloned by iterating all keys.
    $configSnapshot = @{}
    foreach ($key in $Config.Keys) {
        if ($Config[$key] -is [hashtable]) {
            $configSnapshot[$key] = $Config[$key].Clone()
        } else {
            $configSnapshot[$key] = $Config[$key]
        }
    }

    # Capture identity for log enrichment. (F3-08 fix)
    $user     = ''
    $hostname = $env:COMPUTERNAME
    try {
        $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        # Strip domain prefix: DOMAIN\user -> user
        $user = ($identity.Name -split '\\')[-1]
    } catch {
        $user = $env:USERNAME
    }

    $ctx = @{
        RunId     = New-RunId
        TaskName  = $TaskName
        Config    = $configSnapshot
        StartTime = [datetime]::Now
        Params    = $Params
        User      = $user
        Hostname  = $hostname
        SessionId = $PID
        Paths     = @{
            Root    = $Root
            Input   = Join-Path $Root 'input'
            Output  = Join-Path $Root 'output'
            Reports = Join-Path $Root 'reports'
            Logs    = Join-Path $Root 'logs'
        }
    }

    return $ctx
}
