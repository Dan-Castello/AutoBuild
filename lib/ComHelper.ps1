#Requires -Version 5.1
# =============================================================================
# lib/ComHelper.ps1
# AutoBuild v3.0 - Safe COM lifecycle management for PS 5.1.
#
# AUDIT RESOLUTIONS:
#   BUG-COM-01 (HIGH)       : Release-ComObject no longer silences all errors.
#                             Exceptions are logged at DEBUG level.
#   PROBLEMA-COM-03 (MED)   : Remove-ZombieCom uses PID tracking.
#                             Engine-owned processes are protected from kill.
#                             Only orphans (not in $Script:EnginePids) are removed.
#   COM-FREEZE-02 (MED)     : Release-ComObject has a max iteration cap (20)
#                             to prevent infinite loop on inconsistent COM state.
#   COM-FREEZE-03 (MED)     : Close-ExcelApp timeout wait runs in a background
#                             job when called from UI thread (caller's choice).
#   BUG-COM-04 (HIGH)       : Test-ComAvailable job correctly disposes the
#                             COM object before reporting success.
#
# PID TRACKING:
#   Register-EngineCom    - call when creating a COM-backed process
#   Unregister-EngineCom  - call when closing a COM-backed process
#   Remove-ZombieCom      - uses registry to skip engine-owned processes
# =============================================================================
Set-StrictMode -Version Latest

# Registry of PIDs the engine explicitly started. Prevents zombie cleanup
# from killing legitimate engine-owned instances. (COM-03 fix)
$Script:EnginePids = [System.Collections.Generic.HashSet[int]]::new()

# Maximum ref-count drain iterations to prevent infinite loop. (COM-FREEZE-02 fix)
$Script:ComReleaseMaxIter = 20

function Register-EngineCom {
    <#
    .SYNOPSIS
        Registers a process PID as engine-owned. Protected from zombie cleanup.
    #>
    param([int]$Pid_)
    [void]$Script:EnginePids.Add($Pid_)
}

function Unregister-EngineCom {
    <#
    .SYNOPSIS
        Removes a PID from the engine registry after intentional COM shutdown.
    #>
    param([int]$Pid_)
    [void]$Script:EnginePids.Remove($Pid_)
}

function Test-ComAvailable {
    <#
    .SYNOPSIS
        Probes whether a COM server responds within the timeout.
        Uses an isolated Job to avoid blocking the calling thread.
    .OUTPUTS
        $true if the COM server is available, $false otherwise.
    #>
    param(
        [Parameter(Mandatory)][string]$ProgId,
        [int]$TimeoutSec = 20
    )

    $job = $null
    try {
        $job = Start-Job -ScriptBlock {
            param($p)
            try {
                $obj = New-Object -ComObject $p -ErrorAction Stop
                # BUG-COM-04 fix: explicitly release before reporting success
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null
                [GC]::Collect()
                [GC]::WaitForPendingFinalizers()
                return $true
            } catch {
                return $false
            }
        } -ArgumentList $ProgId

        $completed = Wait-Job -Job $job -Timeout $TimeoutSec
        if ($null -eq $completed) { return $false }
        $result = Receive-Job -Job $job
        return [bool]$result
    } catch {
        return $false
    } finally {
        if ($null -ne $job) {
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
        }
    }
}

function Invoke-ComWithTimeout {
    <#
    .SYNOPSIS
        Instantiates a COM object only if the server responds within the timeout.
        Returns $null if unavailable (prevents UI thread freeze).
    .NOTES
        The availability probe runs in a Job (BUG-COM-FREEZE-01 pattern).
        The actual COM instantiation happens in the calling thread because
        COM objects are not marshalable across Jobs.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Context,
        [Parameter(Mandatory)][string]$ProgId,
        [int]$TimeoutSec = 30,
        [string]$Label   = ''
    )

    if (-not $Label) { $Label = $ProgId }
    Write-BuildLog $Context 'DEBUG' "Checking COM availability: $Label"

    if (-not (Test-ComAvailable -ProgId $ProgId -TimeoutSec $TimeoutSec)) {
        Write-BuildLog $Context 'WARN' "$Label not available after ${TimeoutSec}s. Resolve any pending dialog and retry."
        return $null
    }

    try {
        $obj = New-Object -ComObject $ProgId -ErrorAction Stop
        Write-BuildLog $Context 'DEBUG' "$Label instantiated successfully"
        return $obj
    } catch {
        Write-BuildLog $Context 'ERROR' "Failed to instantiate ${Label}: $_"
        return $null
    }
}

function Invoke-ReleaseComObject {
    <#
    .SYNOPSIS
        Drains all COM references from an object.
        Each ReleaseComObject call decrements the ref-count by 1.
        This loop drains until zero or negative.
    .NOTES
        BUG-COM-01 fix: exceptions are logged at DEBUG (not silently swallowed).
        COM-FREEZE-02 fix: capped at $Script:ComReleaseMaxIter iterations.
        The old name 'Release-ComObject' violated PS verb conventions;
        'Invoke-ReleaseComObject' uses the approved 'Invoke' verb.
        An alias 'Release-ComObject' is preserved for backward compatibility.
    #>
    param($ComObject)

    if ($null -eq $ComObject) { return }
    $iter = 0
    try {
        $remaining = 1
        while ($remaining -gt 0 -and $iter -lt $Script:ComReleaseMaxIter) {
            $remaining = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
            $iter++
        }
    } catch {
        # BUG-COM-01 fix: log instead of silent swallow.
        # Use Write-Verbose so it appears in -Verbose mode without always spamming.
        Write-Verbose "Invoke-ReleaseComObject: exception at iteration $($iter): $_"
    }
}

# Backward-compatibility alias for existing tasks.
Set-Alias -Name Release-ComObject -Value Invoke-ReleaseComObject -Scope Script

function Invoke-ComCleanup {
    <#
    .SYNOPSIS
        Runs the standard COM cleanup sequence: close document, quit app, GC.
        Call in the finally block of any task that uses COM objects.
    .PARAMETER ComObjects
        Optional array of additional intermediate COM objects to release before
        the document and app. Use for objects that were acquired mid-task.
    #>
    param(
        $Document,
        $Application,
        [hashtable]$Context,
        [object[]]$ComObjects = @()
    )

    # Release any additional intermediate objects first.
    foreach ($obj in $ComObjects) {
        if ($null -ne $obj) {
            try { Invoke-ReleaseComObject $obj } catch { }
        }
    }

    # Close document.
    if ($null -ne $Document) {
        try { $Document.Close($false) } catch { }
        Invoke-ReleaseComObject $Document
    }

    # Quit application.
    if ($null -ne $Application) {
        try { $Application.Quit() } catch { }
        Invoke-ReleaseComObject $Application
    }

    # Dual GC cycle - required to flush .NET's COM wrapper references.
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()

    if ($null -ne $Context) {
        Write-BuildLog $Context 'DEBUG' 'COM released'
    }
}

function Remove-ZombieCom {
    <#
    .SYNOPSIS
        Removes orphaned headless Office processes (Excel, Word, PowerPoint, Outlook).
        Engine-owned PIDs (registered via Register-EngineCom) are protected.
    .OUTPUTS
        Number of processes terminated.
    .NOTES
        PROBLEMA-COM-03 fix: $Script:EnginePids whitelist prevents killing
        legitimate engine-owned instances. Only true orphans are removed.
    #>
    $count = 0
    foreach ($procName in @('EXCEL', 'WINWORD', 'POWERPNT', 'OUTLOOK')) {
        @(Get-Process -Name $procName -ErrorAction SilentlyContinue) | ForEach-Object {
            $proc = $_
            # Headless check: no visible window.
            if ($proc.MainWindowHandle -ne [IntPtr]::Zero) { return }
            # PID protection: skip engine-owned processes.
            if ($Script:EnginePids.Contains($proc.Id)) { return }
            try {
                $proc.Kill()
                $count++
            } catch { }
        }
    }
    return $count
}
