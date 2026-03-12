#Requires -Version 5.1
# =============================================================================
# queue/QueueManager.psm1  v2.1
# Data model, CRUD, thread-safe via Mutex, JSON persistence.
# ASCII only. PS 5.1 compatible.
# =============================================================================
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$Script:VALID_STATUSES = @('Pending','Queued','Running','Completed','Failed','Canceled','Skipped')
$Script:MUTEX_NAME     = 'Global\AutoBuildQueueMutex'
$Script:MUTEX_TIMEOUT  = 5000
$Script:QueueData      = $null
$Script:QueueIdCounter = 0

function Invoke-WithQueueLock {
    param([scriptblock]$Action)
    $mutex  = New-Object System.Threading.Mutex($false, $Script:MUTEX_NAME)
    $locked = $false
    try {
        $locked = $mutex.WaitOne($Script:MUTEX_TIMEOUT)
        if (-not $locked) { throw "QueueManager: mutex timeout after ${Script:MUTEX_TIMEOUT}ms" }
        return (& $Action)
    } finally {
        if ($locked) { try { $mutex.ReleaseMutex() } catch {} }
        $mutex.Dispose()
    }
}

function New-TaskId {
    try   { return [System.Guid]::NewGuid().ToString('N').ToUpper() }
    catch { $Script:QueueIdCounter++; return ('TASK{0:D8}' -f $Script:QueueIdCounter) }
}

function Assert-QueueReady {
    if ($null -eq $Script:QueueData) {
        throw 'QueueManager: queue not initialized. Call Initialize-Queue first.'
    }
}

function Get-SortedItems {
    return @($Script:QueueData | Sort-Object { [int]$_.Order })
}

function Invoke-RecalcOrders {
    $i = 1
    foreach ($item in @($Script:QueueData | Sort-Object { [int]$_.Order })) {
        $item.Order = $i; $i++
    }
}

function Initialize-Queue {
    param([switch]$Reset)
    if ($null -eq $Script:QueueData -or $Reset) {
        $Script:QueueData = [System.Collections.Generic.List[hashtable]]::new()
    }
    return $Script:QueueData
}

function New-QueueTask {
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Name,
        [string]$Description   = '',
        [string]$ScriptPath    = '',
        [string]$TaskReference = '',
        [hashtable]$Parameters = @{},
        [ValidateRange(1,10)][int]$Priority     = 5,
        [ValidateRange(0,5)][int]$MaxRetries    = 0,
        [int]$TimeoutSeconds   = 0
    )
    return @{
        TaskId         = New-TaskId
        Name           = $Name.Trim()
        Description    = $Description
        ScriptPath     = $ScriptPath
        TaskReference  = if ([string]::IsNullOrWhiteSpace($TaskReference)) { $Name.Trim() } else { $TaskReference.Trim() }
        Parameters     = if ($null -eq $Parameters) { @{} } else { $Parameters.Clone() }
        Status         = 'Pending'
        Priority       = $Priority
        Order          = 0
        CreatedAt      = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss')
        StartedAt      = $null
        CompletedAt    = $null
        Result         = $null
        ErrorMessage   = $null
        RetryCount     = 0
        MaxRetries     = $MaxRetries
        TimeoutSeconds = $TimeoutSeconds
    }
}

function Add-QueueTask {
    param(
        [Parameter(Mandatory)][hashtable]$Task,
        [int]$AtPosition = 0
    )
    Assert-QueueReady
    Invoke-WithQueueLock {
        $item = $Task.Clone()
        $item.Parameters = if ($null -ne $Task.Parameters) { $Task.Parameters.Clone() } else { @{} }
        if ($item.Status -notin $Script:VALID_STATUSES) { $item.Status = 'Pending' }
        $count = $Script:QueueData.Count
        if ($AtPosition -gt 0 -and $AtPosition -le $count) {
            foreach ($e in $Script:QueueData) {
                if ([int]$e.Order -ge $AtPosition) { $e.Order++ }
            }
            $item.Order = $AtPosition
        } else {
            $item.Order = $count + 1
        }
        $Script:QueueData.Add($item)
        Invoke-RecalcOrders
        return $item
    }
}

function Remove-QueueTask {
    param([Parameter(Mandatory)][string]$TaskId)
    Assert-QueueReady
    Invoke-WithQueueLock {
        $item = $Script:QueueData | Where-Object { $_.TaskId -eq $TaskId } | Select-Object -First 1
        if ($null -eq $item) { return $false }
        if ($item.Status -eq 'Running') { Write-Warning "Cannot remove Running task '$($item.Name)'."; return $false }
        [void]$Script:QueueData.Remove($item)
        Invoke-RecalcOrders
        return $true
    }
}

function Edit-QueueTask {
    param(
        [Parameter(Mandatory)][string]$TaskId,
        [string]$Description    = $null,
        [hashtable]$Parameters  = $null,
        [int]$Priority          = 0,
        [int]$MaxRetries        = -1,
        [int]$TimeoutSeconds    = -1
    )
    Assert-QueueReady
    Invoke-WithQueueLock {
        $item = $Script:QueueData | Where-Object { $_.TaskId -eq $TaskId } | Select-Object -First 1
        if ($null -eq $item) { Write-Warning "Task '$TaskId' not found."; return $false }
        if ($item.Status -eq 'Running') { Write-Warning "Cannot edit Running task."; return $false }
        if ($null -ne $Description)   { $item.Description    = $Description }
        if ($null -ne $Parameters)    { $item.Parameters     = $Parameters.Clone() }
        if ($Priority    -gt 0)       { $item.Priority       = $Priority }
        if ($MaxRetries  -ge 0)       { $item.MaxRetries     = $MaxRetries }
        if ($TimeoutSeconds -ge 0)    { $item.TimeoutSeconds = $TimeoutSeconds }
        return $true
    }
}

function Move-QueueTaskUp {
    param([Parameter(Mandatory)][string]$TaskId)
    Assert-QueueReady
    Invoke-WithQueueLock {
        $item = $Script:QueueData | Where-Object { $_.TaskId -eq $TaskId } | Select-Object -First 1
        if ($null -eq $item -or [int]$item.Order -le 1) { return $false }
        $prev = $Script:QueueData | Where-Object { [int]$_.Order -eq ([int]$item.Order - 1) } | Select-Object -First 1
        if ($null -ne $prev) { $prev.Order = $item.Order }
        $item.Order--
        Invoke-RecalcOrders; return $true
    }
}

function Move-QueueTaskDown {
    param([Parameter(Mandatory)][string]$TaskId)
    Assert-QueueReady
    Invoke-WithQueueLock {
        $item   = $Script:QueueData | Where-Object { $_.TaskId -eq $TaskId } | Select-Object -First 1
        $maxOrd = ($Script:QueueData | Measure-Object -Property Order -Maximum).Maximum
        if ($null -eq $item -or [int]$item.Order -ge [int]$maxOrd) { return $false }
        $next = $Script:QueueData | Where-Object { [int]$_.Order -eq ([int]$item.Order + 1) } | Select-Object -First 1
        if ($null -ne $next) { $next.Order = $item.Order }
        $item.Order++
        Invoke-RecalcOrders; return $true
    }
}

function Move-QueueTaskToPosition {
    param(
        [Parameter(Mandatory)][string]$TaskId,
        [Parameter(Mandatory)][int]$Position
    )
    Assert-QueueReady
    Invoke-WithQueueLock {
        $item  = $Script:QueueData | Where-Object { $_.TaskId -eq $TaskId } | Select-Object -First 1
        $count = $Script:QueueData.Count
        if ($null -eq $item) { return $false }
        if ($Position -lt 1)      { $Position = 1 }
        if ($Position -gt $count) { $Position = $count }
        $old = [int]$item.Order; $new = $Position
        if ($old -eq $new) { return $true }
        if ($new -lt $old) {
            foreach ($o in $Script:QueueData) {
                if ($o.TaskId -ne $TaskId) {
                    $v = [int]$o.Order
                    if ($v -ge $new -and $v -lt $old) { $o.Order = $v + 1 }
                }
            }
        } else {
            foreach ($o in $Script:QueueData) {
                if ($o.TaskId -ne $TaskId) {
                    $v = [int]$o.Order
                    if ($v -gt $old -and $v -le $new) { $o.Order = $v - 1 }
                }
            }
        }
        $item.Order = $new
        Invoke-RecalcOrders; return $true
    }
}

function Clear-TaskQueue {
    param([switch]$IncludeRunning)
    Assert-QueueReady
    Invoke-WithQueueLock {
        $before = $Script:QueueData.Count
        if ($IncludeRunning) {
            $Script:QueueData.Clear()
        } else {
            $toRemove = @($Script:QueueData | Where-Object { $_.Status -ne 'Running' })
            foreach ($item in $toRemove) { [void]$Script:QueueData.Remove($item) }
        }
        Invoke-RecalcOrders
        return ($before - $Script:QueueData.Count)
    }
}

function Get-QueueTask {
    param([Parameter(Mandatory)][string]$TaskId)
    Assert-QueueReady
    return ($Script:QueueData | Where-Object { $_.TaskId -eq $TaskId } | Select-Object -First 1)
}

function Get-QueueSnapshot {
    Assert-QueueReady
    return @(
        Get-SortedItems | ForEach-Object {
            $pStr = ''
            if ($null -ne $_.Parameters -and $_.Parameters.Count -gt 0) {
                $pStr = ($_.Parameters.GetEnumerator() | Sort-Object Key |
                         ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '
            }
            [PSCustomObject]@{
                Order          = $_.Order
                TaskId         = $_.TaskId
                Name           = $_.Name
                Description    = $_.Description
                Status         = $_.Status
                Priority       = $_.Priority
                Parameters     = $pStr
                ParametersRaw  = if ($null -ne $_.Parameters) { $_.Parameters.Clone() } else { @{} }
                CreatedAt      = $_.CreatedAt
                StartedAt      = if ($_.StartedAt)    { $_.StartedAt    } else { '' }
                CompletedAt    = if ($_.CompletedAt)  { $_.CompletedAt  } else { '' }
                Result         = if ($_.Result)       { $_.Result       } else { '' }
                ErrorMessage   = if ($_.ErrorMessage) { $_.ErrorMessage } else { '' }
                RetryCount     = $_.RetryCount
                MaxRetries     = $_.MaxRetries
                TimeoutSeconds = $_.TimeoutSeconds
                ScriptPath     = $_.ScriptPath
                TaskReference  = $_.TaskReference
            }
        }
    )
}

function Get-NextPendingTask {
    Assert-QueueReady
    return (
        $Script:QueueData |
        Where-Object { $_.Status -in @('Pending','Queued') } |
        Sort-Object Priority, Order |
        Select-Object -First 1
    )
}

function Set-QueueTaskStatus {
    param(
        [Parameter(Mandatory)][string]$TaskId,
        [Parameter(Mandatory)][ValidateSet('Pending','Queued','Running','Completed','Failed','Canceled','Skipped')]
        [string]$Status,
        [string]$Result       = $null,
        [string]$ErrorMessage = $null,
        [switch]$IncrementRetry
    )
    Assert-QueueReady
    Invoke-WithQueueLock {
        $item = $Script:QueueData | Where-Object { $_.TaskId -eq $TaskId } | Select-Object -First 1
        if ($null -eq $item) { return $false }
        $item.Status = $Status
        $now = Get-Date -Format 'yyyy-MM-ddTHH:mm:ss'
        switch ($Status) {
            'Running'   { $item.StartedAt   = $now }
            'Completed' { $item.CompletedAt = $now }
            'Failed'    { $item.CompletedAt = $now }
            'Canceled'  { $item.CompletedAt = $now }
            'Skipped'   { $item.CompletedAt = $now }
        }
        if ($null -ne $Result)       { $item.Result       = $Result }
        if ($null -ne $ErrorMessage) { $item.ErrorMessage = $ErrorMessage }
        if ($IncrementRetry)         { $item.RetryCount++ }
        return $true
    }
}

function Save-TaskQueue {
    param([Parameter(Mandatory)][string]$Path)
    Assert-QueueReady
    try {
        $data = @(Get-SortedItems)
        $json = $data | ConvertTo-Json -Depth 5
        $dir  = Split-Path $Path -Parent
        if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
        [System.IO.File]::WriteAllText($Path, $json, [System.Text.Encoding]::UTF8)
        return $true
    } catch { Write-Warning "Save-TaskQueue: $_"; return $false }
}

function Import-TaskQueue {
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$Replace,
        [switch]$OnlyPending
    )
    Assert-QueueReady
    if (-not (Test-Path $Path)) { Write-Warning "Import-TaskQueue: file not found '$Path'"; return 0 }
    try {
        $items = @(Get-Content $Path -Raw -Encoding UTF8 | ConvertFrom-Json)
        if ($OnlyPending) { $items = @($items | Where-Object { $_.Status -in @('Pending','Queued') }) }
        if ($Replace) { $Script:QueueData.Clear() }
        $loaded = 0
        foreach ($obj in $items) {
            try {
                $params = @{}
                if ($null -ne $obj.Parameters) {
                    if ($obj.Parameters -is [hashtable]) { $params = $obj.Parameters }
                    elseif ($obj.Parameters.PSObject.Properties) {
                        foreach ($p in $obj.Parameters.PSObject.Properties) { $params[$p.Name] = $p.Value }
                    }
                }
                $t = New-QueueTask -Name ($obj.Name -as [string]) `
                                   -Description ($obj.Description -as [string]) `
                                   -ScriptPath ($obj.ScriptPath -as [string]) `
                                   -TaskReference ($obj.TaskReference -as [string]) `
                                   -Parameters $params `
                                   -Priority ([int]($obj.Priority)) `
                                   -MaxRetries ([int]($obj.MaxRetries)) `
                                   -TimeoutSeconds ([int]($obj.TimeoutSeconds))
                if ($obj.Status -in $Script:VALID_STATUSES) { $t.Status = $obj.Status }
                Add-QueueTask -Task $t | Out-Null; $loaded++
            } catch { Write-Warning "Import-TaskQueue: skip '$($obj.Name)': $_" }
        }
        return $loaded
    } catch { Write-Warning "Import-TaskQueue: read error: $_"; return 0 }
}

Export-ModuleMember -Function @(
    'Initialize-Queue','New-QueueTask','Add-QueueTask','Remove-QueueTask',
    'Edit-QueueTask','Move-QueueTaskUp','Move-QueueTaskDown','Move-QueueTaskToPosition',
    'Clear-TaskQueue','Get-QueueTask','Get-QueueSnapshot','Get-NextPendingTask',
    'Set-QueueTaskStatus','Save-TaskQueue','Import-TaskQueue'
)
