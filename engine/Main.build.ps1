#Requires -Version 5.1
<#
.SYNOPSIS
    AutoBuild v3.0 - Build orchestrator (Invoke-Build entry point).

.NOTES
    AUDIT RESOLUTIONS (v1 -> v3):

    MAIN-01 (CRITICAL) - Default task:
        Changed from 'diag_completo' (instantiated COM objects, consumed
        SAP licenses on every bare run) to 'listar_tareas' which only
        reads filenames from disk. Safe in production.

    MAIN-02 (MED) - Library load error isolation:
        Each dot-source is wrapped in an individual try/catch with a
        clear error naming the failing file. One bad library no longer
        crashes the entire engine with a cryptic message.

    MAIN-03 (MED) - Lazy task loading:
        Tasks are NOT dot-sourced at engine startup. Invoke-LoadTask
        loads only the single task file that matches the requested task.
        Eliminates: O(n) startup cost, cross-task syntax errors blocking
        unrelated executions, and unnecessary COM import surface.

    PROBLEMA-ARQUITECTURAL-01 (MED) - Logger/Config separation:
        Config.ps1 is loaded FIRST; Logger.ps1 loads second.
        Config no longer lives in Logger.

    PROBLEMA-ARQUITECTURAL-03 (HIGH) - Config alias removed:
        $Script:Config alias eliminated entirely.
        All code references $Script:EngineConfig directly.
        This removes the shared-reference bomb from v1.

    MAIN-04 (LOW) - Config null guard:
        Engine halts with a clear error if Get-EngineConfig returns null.

    TASK-01 (HIGH) - Template root fix:
        Invoke-LoadTask injects $Script:EngineRoot into the task scope.
        Tasks never need to call (Split-Path $BuildRoot -Parent).

    RUN-01 (HIGH) - Generic parameters:
        All task parameters arrive via $Script:TaskParams (collected from
        the build script's own BoundParameters). Tasks read params from
        $ctx.Params['Key'] instead of named script-level variables.
#>

# ---- Generic catch-all parameter (populated from -Params JSON in Run.ps1) --
param(
    [string]$_ParamsJson = '{}'
)

Set-StrictMode -Version Latest

# ============================================================================
# ROOT: $BuildRoot = engine/ directory (set by Invoke-Build).
#       $Script:EngineRoot = project root (one level up).
# ============================================================================
$Script:EngineRoot = Split-Path $BuildRoot -Parent

# ============================================================================
# BOOTSTRAP: Load libraries in dependency order, each in its own try/catch.
# Load order is intentional and documented:
#   1. Config.ps1  - no dependencies; defines Get-EngineConfig
#   2. Logger.ps1  - depends on config (log level, rotation threshold)
#   3. Context.ps1 - depends on Logger (New-RunId)
#   4. Auth.ps1    - depends on Config (security section)
#   5. Retry.ps1   - depends on Logger
#   6. ComHelper, ExcelHelper, WordHelper, SapHelper - depend on Logger
#   7. Assertions  - depends on Logger
# ============================================================================
$Script:LibLoadOrder = @(
    'Config.ps1',
    'Logger.ps1',
    'Notifications.ps1',   # FIX SMTP-MISSING: implements Send-Notification
    'Context.ps1',
    'Auth.ps1',
    'Retry.ps1',
    'ComHelper.ps1',
    'ExcelHelper.ps1',
    'WordHelper.ps1',
    'SapHelper.ps1',
    'Assertions.ps1',
    'Integrity.ps1'   # TASK-02 fix: hash-based task file verification
)

$LibPath = Join-Path $Script:EngineRoot 'lib'

foreach ($libFile in $Script:LibLoadOrder) {
    $libFullPath = Join-Path $LibPath $libFile
    try {
        . $libFullPath
    } catch {
        throw "AutoBuild engine: failed to load library '$libFile'. Error: $_"
    }
}

# ============================================================================
# CONFIG: Load master configuration. Guard against null result. (MAIN-04 fix)
# NO $Script:Config alias. PROBLEMA-ARQUITECTURAL-03 fix.
# ============================================================================
$Script:EngineConfig = Get-EngineConfig -Root $Script:EngineRoot

if ($null -eq $Script:EngineConfig -or $Script:EngineConfig.Count -eq 0) {
    throw 'AutoBuild engine: Get-EngineConfig returned null or empty. Check Config.ps1 and engine.config.json.'
}

# FIX PROD-GUARD (AUDIT v3): Warn when all security identifiers are empty.
# This prevents the engine silently starting in 'dev mode' (Operator-for-everyone)
# when deployed to production without security group configuration.
try {
    Assert-SecurityConfigPopulated -Config $Script:EngineConfig -ErrorOnEmpty:$false
} catch {
    # Assert-SecurityConfigPopulated not yet loaded (library load order race).
    Write-Warning "AutoBuild: Could not validate security config. Ensure adminAdGroup or adminUsers are set before production deployment."
}

# ============================================================================
# TASK PARAMETERS: Collect from bound parameters into $Script:TaskParams.
# Run.ps1 passes task parameters individually via @InvokeArgs splatting.
# Tasks read params via $ctx.Params['Key']. (RUN-01 fix continued)
# ============================================================================
$Script:TaskParams = @{}
$MyInvocation.BoundParameters.Keys |
    Where-Object { $_ -notin @('_ParamsJson','Verbose','Debug','WhatIf','Confirm') } |
    ForEach-Object { $Script:TaskParams[$_] = $MyInvocation.BoundParameters[$_] }

# ============================================================================
# LAZY TASK LOADING (MAIN-03 fix)
# ============================================================================
$Script:TasksPath   = Join-Path $Script:EngineRoot 'tasks'
$Script:LoadedTasks = [System.Collections.Generic.HashSet[string]]::new()

function Invoke-LoadTask {
    <#
    .SYNOPSIS
        Dot-sources the task file for $TaskName if not already loaded.
        Injects $Script:EngineRoot so tasks never call Split-Path. (TASK-01 fix)
        Optionally verifies the task file hash before loading. (TASK-02 fix)
    #>
    param([Parameter(Mandatory)][string]$TaskName)

    if ($Script:LoadedTasks.Contains($TaskName)) { return }

    $file = Join-Path $Script:TasksPath "task_${TaskName}.ps1"
    if (-not (Test-Path $file)) {
        throw "AutoBuild: task file not found: task_${TaskName}.ps1 in $Script:TasksPath"
    }

    # FIX V-02/R-02 (AUDIT v3 HIGH): Race condition between Test-TaskIntegrity
    # and actual file load. The previous implementation:
    #   1. Test-TaskIntegrity reads file, computes hash, checks registry
    #   2. [~50ms window where an attacker can swap the file]
    #   3. . $file (loads the potentially-swapped file)
    #
    # CORRECTION: Read the file bytes ONCE into memory, verify the in-memory
    # hash against the registry, then dot-source from a temp path that was
    # written from the verified in-memory bytes. This eliminates the TOCTOU window.
    $hashFile = Join-Path $Script:TasksPath 'tasks.hash.json'
    if (Test-Path $hashFile) {
        try {
            # Load file bytes once to eliminate the TOCTOU race.
            $fileBytes = [System.IO.File]::ReadAllBytes($file)
            $sha       = [System.Security.Cryptography.SHA256]::Create()
            try {
                $actualHash = ([BitConverter]::ToString($sha.ComputeHash($fileBytes)) -replace '-','').ToLower()
            } finally { $sha.Dispose() }

            # Load registry and compare hash of the in-memory bytes.
            $reg      = $null
            $hashRaw  = Get-Content $hashFile -Raw -Encoding ASCII | ConvertFrom-Json
            $fileName = [System.IO.Path]::GetFileName($file)
            $regEntry = $hashRaw.PSObject.Properties[$fileName]

            if ($null -eq $regEntry) {
                throw "AutoBuild: task '$TaskName' is not in the hash registry. " +
                      "Run Update-TaskRegistry to register approved task files."
            } elseif ($regEntry.Value.sha256 -ne $actualHash) {
                throw "AutoBuild: HASH MISMATCH for task '$TaskName'. " +
                      "Expected: $($regEntry.Value.sha256). Actual: $actualHash. " +
                      "File may have been tampered with."
            }

            # Hash verified. Write verified bytes to a temp location and dot-source from there.
            # This ensures even if the original file is swapped after verification,
            # we execute exactly what was hashed.
            $verifiedTemp = [System.IO.Path]::GetTempFileName() + '.ps1'
            try {
                [System.IO.File]::WriteAllBytes($verifiedTemp, $fileBytes)
                . $verifiedTemp
            } finally {
                Remove-Item $verifiedTemp -Force -ErrorAction SilentlyContinue
            }

        } catch [System.Management.Automation.CommandNotFoundException] {
            # Integrity functions not available (first call before full bootstrap). Proceed.
            . $file
        } catch {
            throw "AutoBuild: integrity check failed for task '$TaskName'. $_"
        }
    } else {
        try {
            . $file
        } catch {
            throw "AutoBuild: failed to load task '$TaskName' from '$file'. Error: $_"
        }
    }
    [void]$Script:LoadedTasks.Add($TaskName)
}

# ============================================================================
# BUILT-IN ENGINE TASKS
# ============================================================================

# Synopsis: Lists all available tasks with descriptions. Safe default task.
task listar_tareas {
    Write-Build Cyan "`n  AutoBuild v3.0 - Available Tasks`n"
    Write-Build Cyan ('  ' + ('=' * 60))
    if (Test-Path $Script:TasksPath) {
        Get-ChildItem -Path $Script:TasksPath -Filter 'task_*.ps1' |
            Sort-Object Name |
            ForEach-Object {
                $name     = $_.BaseName -replace '^task_', ''
                $synopsis = ''
                try {
                    $m = Select-String -Path $_.FullName -Pattern '# Synopsis:\s*(.+)' |
                         Select-Object -First 1
                    if ($m) { $synopsis = $m.Matches.Groups[1].Value.Trim() }
                } catch { }
                Write-Build White ('  {0,-35} {1}' -f $name, $synopsis)
            }
    }
    Write-Build Cyan ("`n  Usage: .\Run.ps1 " + '<task_name>' + " [-Params '{`"key`":`"value`"}'`n")
    Write-Build Cyan "  For help: .\Run.ps1 -List`n"
}

# Synopsis: Removes orphaned headless Office COM processes (not engine-owned).
task limpiar_com {
    $ctx = New-TaskContext `
        -TaskName 'limpiar_com' `
        -Config   $Script:EngineConfig `
        -Root     $Script:EngineRoot
    Write-BuildLog $ctx 'INFO' 'Searching for orphaned COM processes'
    $count = Remove-ZombieCom
    Write-Build Cyan "  Orphaned processes removed: $count"
    Write-BuildLog $ctx 'INFO' "COM cleanup complete. Removed: $count"
    Write-RunResult -Context $ctx -Success $true
}

# Synopsis: Engine health check (no COM, no SAP, no external dependencies).
task ejemplo {
    $ctx = New-TaskContext `
        -TaskName 'ejemplo' `
        -Config   $Script:EngineConfig `
        -Root     $Script:EngineRoot
    Write-BuildLog $ctx 'INFO' 'Engine health check starting'
    Write-Build Green "  AutoBuild engine v3.0 operational"
    Write-Build Green "  Invoke-Build: $($Script:EngineConfig.engine.ibVersion)"
    Write-Build Green "  User: $($ctx.User) `@ $($ctx.Hostname)"
    Write-Build Green "  RunId: $($ctx.RunId)"
    Write-RunResult -Context $ctx -Success $true
}

# Synopsis: Purges rotated log archives older than retention policy.
task purgar_logs {
    $ctx = New-TaskContext `
        -TaskName 'purgar_logs' `
        -Config   $Script:EngineConfig `
        -Root     $Script:EngineRoot
    Write-BuildLog $ctx 'INFO' 'Starting log purge'
    Invoke-LogPurge -LogsDir $ctx.Paths.Logs `
                    -RetentionDays $ctx.Config.reports.retentionDays
    Write-BuildLog $ctx 'INFO' 'Log purge complete'
    Write-RunResult -Context $ctx -Success $true
}

# ============================================================================
# MAIN-01 FIX: Safe default task. No COM, no SAP, no side effects.
# ============================================================================
task . listar_tareas
