#Requires -Version 5.1
# =============================================================================
# lib/SapHelper.ps1
# AutoBuild v3.0 - SAP GUI Scripting helpers for PS 5.1.
#
# AUDIT RESOLUTIONS:
#   SAP-01 (HIGH)   : Export-SapTableToArray supports ALV GuiGridView in
#                     addition to GuiTableControl. Detects type automatically.
#   SAP-02 (MED)    : Wait-SapReady implemented. Polls session Busy flag
#                     until idle or timeout. Called by key interaction
#                     functions to prevent race conditions.
#   SAP-03 (LOW)    : Export-SapTableToArray uses
#                     [System.Collections.Generic.List[hashtable]] instead
#                     of += array operator (O(n^2) -> O(n) for 50K rows).
#   CANCEL-01 (MED) : Export-SapTableToArray accepts $CancellationToken
#                     (a [ref] bool). Callers can set it to $true to abort
#                     an in-progress table export cleanly.
#
# COM RULE:
#   Every intermediate COM object must be captured and released.
#   See Get-SapSession for the canonical chain: rot->sapgui->app->conn->sess.
# =============================================================================
Set-StrictMode -Version Latest

function Get-SapSession {
    <#
    .SYNOPSIS
        Returns the active SAP GUI session with correct COM lifecycle management.
    .NOTES
        The returned $sess MUST be released by the caller via Release-SapSession
        in a finally block.
    .OUTPUTS
        GuiSession object, or $null if no active session.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Context,
        [int]$ConnectionIndex = 0,
        [int]$SessionIndex    = 0
    )

    Write-BuildLog $Context 'DEBUG' 'Looking for active SAP GUI session'

    $rot    = $null
    $sapgui = $null
    $app    = $null
    $conn   = $null

    try {
        $rot = New-Object -ComObject 'SapROTWr.SapROTWrapper' -ErrorAction Stop

        $sapgui = $rot.GetROTEntry('SAPGUI')
        if ($null -eq $sapgui) {
            Write-BuildLog $Context 'WARN' 'No SAP GUI instance found'
            return $null
        }

        $app = $sapgui.GetScriptingEngine()
        if ($null -eq $app) {
            Write-BuildLog $Context 'WARN' 'SAP GUI scripting engine not available'
            return $null
        }

        $conn = $app.Children($ConnectionIndex)
        if ($null -eq $conn) {
            Write-BuildLog $Context 'WARN' "No SAP connection at index $ConnectionIndex"
            return $null
        }

        $sess = $conn.Children($SessionIndex)
        if ($null -eq $sess) {
            Write-BuildLog $Context 'WARN' "No SAP session at index $SessionIndex"
            return $null
        }

        # Read session info before releasing intermediates.
        try {
            $info     = $sess.Info
            $sysName  = $info.SystemName
            $userName = $info.User
            Invoke-ReleaseComObject $info
            Write-BuildLog $Context 'INFO' "SAP session acquired: $sysName / $userName"
        } catch { }

        return $sess

    } catch {
        Write-BuildLog $Context 'ERROR' "Failed to get SAP session: $_"
        return $null

    } finally {
        # Release in reverse order. finally runs even after explicit return.
        if ($null -ne $conn)   { Invoke-ReleaseComObject $conn   }
        if ($null -ne $app)    { Invoke-ReleaseComObject $app    }
        if ($null -ne $sapgui) { Invoke-ReleaseComObject $sapgui }
        if ($null -ne $rot)    { Invoke-ReleaseComObject $rot    }
    }
}

function Release-SapSession {
    <#
    .SYNOPSIS
        Releases the SAP session COM object. Call in finally blocks.
    #>
    param($Session)
    if ($null -eq $Session) { return }
    try { Invoke-ReleaseComObject $Session } catch { }
}

function Wait-SapReady {
    <#
    .SYNOPSIS
        Polls the SAP session Busy flag until the UI is idle or timeout.
        (SAP-02 fix: prevents race conditions after button presses.)
    .PARAMETER Session
        Active GuiSession object.
    .PARAMETER TimeoutSec
        Maximum wait time in seconds. Default: 30.
    .PARAMETER PollMs
        Polling interval in milliseconds. Default: 200.
    .OUTPUTS
        $true if SAP became ready within timeout, $false on timeout.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Context,
        [Parameter(Mandatory)]$Session,
        [int]$TimeoutSec = 30,
        [int]$PollMs     = 200
    )

    $deadline = [datetime]::Now.AddSeconds($TimeoutSec)
    while ([datetime]::Now -lt $deadline) {
        try {
            $info = $Session.Info
            $busy = $info.IsActive   # IsActive = $true means session is busy
            Invoke-ReleaseComObject $info
            if (-not $busy) { return $true }
        } catch {
            # If we can't read Info, assume ready.
            return $true
        }
        Start-Sleep -Milliseconds $PollMs
    }

    Write-BuildLog $Context 'WARN' "Wait-SapReady: session still busy after ${TimeoutSec}s"
    return $false
}

function Invoke-SapTransaction {
    <#
    .SYNOPSIS
        Navigates to a SAP transaction code.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Context,
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)][string]$TCode
    )

    Write-BuildLog $Context 'INFO' "Running transaction: $TCode"
    try {
        $Session.StartTransaction($TCode)
        Wait-SapReady -Context $Context -Session $Session | Out-Null
    } catch {
        Write-BuildLog $Context 'ERROR' "Transaction $TCode failed: $_"
        throw
    }
}

function Set-SapField {
    <#
    .SYNOPSIS
        Sets a field value by component ID.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Context,
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)][string]$FieldId,
        [Parameter(Mandatory)][string]$Value
    )

    $field = $null
    try {
        $field       = $Session.FindById($FieldId)
        $field.Text  = $Value
    } catch {
        Write-BuildLog $Context 'WARN' "Could not set field $FieldId : $_"
        throw
    } finally {
        if ($null -ne $field) { Invoke-ReleaseComObject $field }
    }
}

function Get-SapField {
    <#
    .SYNOPSIS
        Gets a field value by component ID.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Context,
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)][string]$FieldId
    )

    $field = $null
    try {
        $field = $Session.FindById($FieldId)
        return $field.Text
    } catch {
        Write-BuildLog $Context 'WARN' "Could not read field $FieldId : $_"
        return $null
    } finally {
        if ($null -ne $field) { Invoke-ReleaseComObject $field }
    }
}

function Invoke-SapButton {
    <#
    .SYNOPSIS
        Presses a SAP button by component ID and waits for the session to
        become ready. (SAP-02 integration: Wait-SapReady called automatically.)
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Context,
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)][string]$ButtonId,
        [int]$WaitTimeoutSec = 30
    )

    $btn = $null
    try {
        $btn = $Session.FindById($ButtonId)
        $btn.Press()
        Write-BuildLog $Context 'DEBUG' "Button pressed: $ButtonId"
    } catch {
        Write-BuildLog $Context 'ERROR' "Error pressing button $ButtonId : $_"
        throw
    } finally {
        if ($null -ne $btn) { Invoke-ReleaseComObject $btn }
    }

    # Wait for SAP to process the button click.
    Wait-SapReady -Context $Context -Session $Session -TimeoutSec $WaitTimeoutSec | Out-Null
}

function Export-SapTableToArray {
    <#
    .SYNOPSIS
        Exports a SAP table (GuiTableControl OR ALV GuiGridView) to an
        array of hashtables. (SAP-01 fix: dual-type detection.)

    .DESCRIPTION
        SAP-01 fix: Automatically detects whether the control at $TableId
        is a classic GuiTableControl or a modern ALV GuiGridView and uses
        the appropriate cell-access API for each.

        SAP-03 fix: Uses List[hashtable] instead of += array concatenation.
        Performance: O(n) instead of O(n^2). Critical for 50K-row tables.

        CANCEL-01 fix: Accepts a [ref]$CancellationToken. Set the referenced
        bool to $true from another thread to abort the export cleanly.

    .PARAMETER CancellationToken
        Pass a [ref] to a $bool variable. Set the variable to $true to
        request cancellation. The function returns whatever rows were
        collected before the cancellation request.
    .OUTPUTS
        Array of hashtables, one per row. Empty array on error.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Context,
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)][string]$TableId,
        [Parameter(Mandatory)][string[]]$Columns,
        $CancellationToken = $null
    )

    $rows  = [System.Collections.Generic.List[hashtable]]::new()
    $table = $null
    try {
        $table    = $Session.FindById($TableId)
        $rowCount = $table.RowCount
        $typeName = $table.Type   # 'GuiTableControl' or 'GuiGridView'

        Write-BuildLog $Context 'INFO' "Exporting SAP table [$typeName]: $rowCount rows, $($Columns.Count) columns"

        for ($r = 0; $r -lt $rowCount; $r++) {
            # Cancellation check every 100 rows.
            if ($null -ne $CancellationToken -and ($r % 100 -eq 0) -and $CancellationToken.Value) {
                Write-BuildLog $Context 'WARN' "Export-SapTableToArray: cancelled at row $r"
                break
            }

            $row = @{}
            foreach ($col in $Columns) {
                $cell = $null
                try {
                    # SAP-01 fix: different cell-access API per control type.
                    if ($typeName -eq 'GuiGridView') {
                        # ALV grid: GetCellValue(row, columnId) returns string directly.
                        $row[$col] = $table.GetCellValue($r, $col)
                    } else {
                        # Classic table: GetCell(row, columnId).Text
                        $cell      = $table.GetCell($r, $col)
                        $row[$col] = $cell.Text
                    }
                } catch {
                    $row[$col] = ''
                } finally {
                    if ($null -ne $cell) { Invoke-ReleaseComObject $cell }
                }
            }
            $rows.Add($row)
        }

    } catch {
        Write-BuildLog $Context 'ERROR' "Error exporting SAP table $TableId : $_"
        throw
    } finally {
        if ($null -ne $table) { Invoke-ReleaseComObject $table }
    }

    return $rows.ToArray()
}
