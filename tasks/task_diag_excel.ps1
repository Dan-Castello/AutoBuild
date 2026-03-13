#Requires -Version 5.1
# =============================================================================
# tasks/task_diag_excel.ps1
# @Description : Excel COM diagnostics - lifecycle, write performance, formats, cleanup
# @Category    : Excel
# @Version     : 1.1.0
# @Author      : AutoBuild QA
# =============================================================================
# Synopsis: Diagnostics - Excel COM (availability, write speed, formats, PID tracking, zombie)
# Params: {"RowCount":"500","TestFormats":"true"}
#
# REMEDIATION v3.1 - FIX-EXCEL-NULL:
#   Root cause of "No se puede llamar a un método en una expresión con valor NULL":
#   Get-ExcelSheet and Add-ExcelSheet were called unconditionally after
#   New-ExcelWorkbook, even when $wb was $null (workbook creation failed).
#   Calling $null.Worksheets triggers the NullReference error reported.
#
#   Fix: Guard all worksheet operations with ($null -ne $wb) checks.
#   Any step that depends on a null object is recorded as FAIL with a
#   descriptive detail, and the block is skipped — the task continues to
#   the next diagnostic section rather than throwing from inside the try.
#
#   Additional hardening:
#   - New-ExcelApp null check before workbook creation
#   - All COM method calls guarded against null objects
#   - Explicit null-initialisation for all COM variables before try block

task diag_excel {

    $ctx = New-TaskContext `
        -TaskName 'diag_excel' `
        -Config   $Script:EngineConfig `
        -Root     $Script:EngineRoot `
        -Params   $Script:TaskParams

    Write-BuildLog $ctx 'INFO' 'DIAG EXCEL: Starting Excel COM diagnostics'

    $rowCount    = try { $rv = $ctx.Params['RowCount'];    if ($null -eq $rv -or $rv -eq '') { 500 } else { [int]$rv }    } catch { 500 }
    $testFormats = try { $tv = $ctx.Params['TestFormats']; if ($null -eq $tv -or $tv -eq '') { $true } else { [bool]::Parse($tv) } } catch { $true }
    if ($rowCount -le 0) { $rowCount = 500 }

    $results  = [System.Collections.Generic.List[hashtable]]::new()
    $failures = 0
    $script:failures = 0

    function Add-Result {
        param([string]$Check, [bool]$Pass, [string]$Detail = '')
        $results.Add(@{ Check = $Check; Pass = $Pass; Detail = $Detail })
        $sym   = if ($Pass) { '[OK]' } else { '[FAIL]' }
        $lvl   = if ($Pass) { 'INFO' } else { 'ERROR' }
        $color = if ($Pass) { 'Green' } else { 'Red' }
        Write-BuildLog $ctx $lvl "$sym $Check" -Detail $Detail
        Write-Build $color "  $sym  $Check$(if($Detail){' -- '+$Detail})"
        if (-not $Pass) { $script:failures++ }
    }

    # ── 1. AVAILABILITY CHECK ────────────────────────────────────────────────
    Write-Build Cyan "`n  [1/6] Excel COM Availability"
    $available = Test-ComAvailable -ProgId 'Excel.Application' -TimeoutSec 20
    Add-Result 'Excel.Application COM available' $available
    if (-not $available) {
        Write-Build Yellow '  Excel not available. Skipping remaining Excel checks.'
        Write-BuildLog $ctx 'WARN' 'Excel not available -- skipping Excel COM tests'
        Write-RunResult -Context $ctx -Success $false -ErrorMsg 'Excel COM not available'
        return
    }

    # All COM handles — initialise to $null before the try so finally block is safe.
    $xl  = $null
    $wb  = $null
    $ws  = $null
    $wb2 = $null
    $ws3 = $null

    try {

        # ── 2. INSTANTIATION AND PID TRACKING ────────────────────────────────
        Write-Build Cyan "`n  [2/6] Instantiation and PID Tracking"
        $pidsBefore = @(Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
        $xl = New-ExcelApp -Context $ctx -TimeoutSec 30

        # FIX-EXCEL-NULL (partial): Guard ExcelApp before any downstream use.
        if ($null -eq $xl) {
            Add-Result 'New-ExcelApp returns non-null' $false 'COM instantiation failed — cannot continue Excel tests'
            Write-BuildLog $ctx 'ERROR' 'New-ExcelApp returned null. COM layer failed.'
            # Cannot proceed without Excel app — exit try cleanly.
            throw 'Excel COM application could not be instantiated'
        }
        Add-Result 'New-ExcelApp returns non-null' $true

        $pidsAfter = @(Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
        $newPids   = @($pidsAfter | Where-Object { $pidsBefore -notcontains $_ })
        Add-Result 'New Excel PID detected by set-difference' ($newPids.Count -gt 0) `
            "NewPIDs=$($newPids -join ',')"

        $pidProtected = $false
        if ($newPids.Count -gt 0) {
            Remove-ZombieCom | Out-Null
            $stillRunning = $null -ne (Get-Process -Id $newPids[0] -ErrorAction SilentlyContinue)
            $pidProtected = $stillRunning
        }
        Add-Result 'Excel PID protected from zombie cleanup' $pidProtected

        # ── 3. WORKBOOK AND SHEET OPERATIONS ─────────────────────────────────
        Write-Build Cyan "`n  [3/6] Workbook and Sheet Operations"

        # FIX-EXCEL-NULL (primary fix): New-ExcelWorkbook is the previously-failing
        # call site. If Add() throws internally (COM error, macro alert, etc.),
        # New-ExcelWorkbook re-throws and $wb remains $null.
        # Guard: all downstream sheet operations check ($null -ne $wb).
        try {
            $wb = New-ExcelWorkbook -Context $ctx -ExcelApp $xl
        } catch {
            Write-BuildLog $ctx 'ERROR' "New-ExcelWorkbook threw: $_"
            $wb = $null
        }
        Add-Result 'New-ExcelWorkbook creates workbook' ($null -ne $wb) `
            $(if($null -eq $wb){'Workbook is null — COM Add() failed. Check for pending Excel dialogs or macro alerts.'})

        # FIX-EXCEL-NULL (downstream guard): Skip sheet operations if $wb is null.
        # Without this guard, Get-ExcelSheet calls $null.Worksheets -> NullReference.
        if ($null -ne $wb) {
            try {
                $ws = Get-ExcelSheet -Workbook $wb -Index 1
            } catch {
                Write-BuildLog $ctx 'ERROR' "Get-ExcelSheet threw: $_"
                $ws = $null
            }
            Add-Result 'Get-ExcelSheet returns sheet' ($null -ne $ws) `
                $(if($null -eq $ws){'Sheet is null — Worksheets.Item(1) failed'})

            # FIX-EXCEL-NULL: Add-ExcelSheet also requires non-null $wb.
            $ws2 = $null
            try {
                $ws2 = Add-ExcelSheet -Workbook $wb -Name 'Sheet2'
            } catch {
                Write-BuildLog $ctx 'WARN' "Add-ExcelSheet threw: $_"
                $ws2 = $null
            }
            Add-Result 'Add-ExcelSheet creates named sheet' ($null -ne $ws2)
            if ($null -ne $ws2) { Invoke-ReleaseComObject $ws2; $ws2 = $null }
        } else {
            # Record downstream checks as failed (workbook prerequisite missing).
            Add-Result 'Get-ExcelSheet returns sheet'      $false 'Skipped: workbook is null'
            Add-Result 'Add-ExcelSheet creates named sheet' $false 'Skipped: workbook is null'
        }

        # ── 4. BULK WRITE PERFORMANCE ─────────────────────────────────────────
        Write-Build Cyan "`n  [4/6] Bulk Write Performance ($rowCount rows)"

        if ($null -ne $ws) {
            $data = [System.Collections.Generic.List[hashtable]]::new()
            for ($i = 1; $i -le $rowCount; $i++) {
                $data.Add(@{
                    ID        = $i
                    Name      = "Item_$i"
                    Category  = @('A','B','C','D')[$i % 4]
                    Value     = [math]::Round([math]::Sin($i) * 1000, 2)
                    Timestamp = (Get-Date).AddMinutes(-$i).ToString('yyyy-MM-dd HH:mm:ss')
                    Status    = @('OK','WARN','ERROR','PENDING')[$i % 4]
                })
            }

            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            Write-ExcelRange -Context $ctx -Sheet $ws -Data $data.ToArray() -WriteHeaders $true
            $sw.Stop()

            $msPerRow = if ($rowCount -gt 0) { [math]::Round($sw.ElapsedMilliseconds / $rowCount, 2) } else { 0 }
            Add-Result "Write-ExcelRange: $rowCount rows in $($sw.ElapsedMilliseconds)ms" `
                ($sw.ElapsedMilliseconds -lt 30000) "ms/row=$msPerRow"
            Write-BuildLog $ctx 'INFO' "Bulk write: $rowCount rows in $($sw.ElapsedMilliseconds)ms ($msPerRow ms/row)"

            $used = Get-ExcelUsedRange -Sheet $ws
            $expectedRows = $rowCount + 1
            Add-Result 'Used range row count matches data' ($used.Rows -eq $expectedRows) `
                "Expected=$expectedRows Actual=$($used.Rows)"

            Invoke-ExcelAutoFit -Sheet $ws
        } else {
            Add-Result "Write-ExcelRange: $rowCount rows" $false 'Skipped: sheet is null'
            Add-Result 'Used range row count matches data' $false 'Skipped: sheet is null'
        }

        # ── 5. SAVE IN MULTIPLE FORMATS ───────────────────────────────────────
        Write-Build Cyan "`n  [5/6] Save Formats"
        $stamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
        $xlsxFile = $null

        if ($null -ne $wb) {
            $xlsxFile = Join-Path $ctx.Paths.Reports "diag_excel_${stamp}.xlsx"
            try {
                Save-ExcelWorkbook -Context $ctx -Workbook $wb -Path $xlsxFile -Format 'xlsx'
            } catch {
                Write-BuildLog $ctx 'ERROR' "Save xlsx failed: $_"
                $xlsxFile = $null
            }
            Add-Result 'Save as .xlsx' ($null -ne $xlsxFile -and (Test-Path $xlsxFile)) "$xlsxFile"

            if ($testFormats -and $null -ne $xlsxFile) {
                $csvFile = Join-Path $ctx.Paths.Reports "diag_excel_${stamp}.csv"
                try {
                    Save-ExcelWorkbook -Context $ctx -Workbook $wb -Path $csvFile -Format 'csv'
                } catch {
                    Write-BuildLog $ctx 'WARN' "Save csv failed: $_"
                    $csvFile = $null
                }
                Add-Result 'Save as .csv' ($null -ne $csvFile -and (Test-Path $csvFile)) "$csvFile"

                if ($null -ne $csvFile -and (Test-Path $csvFile)) {
                    $csvRows = @(Import-Csv $csvFile -ErrorAction SilentlyContinue)
                    Add-Result 'CSV row count matches source data' ($csvRows.Count -eq $rowCount) `
                        "Expected=$rowCount CSV=$($csvRows.Count)"
                    Add-Result 'CSV has ID column' ($csvRows.Count -gt 0 -and $null -ne $csvRows[0].ID)
                }
            }
        } else {
            Add-Result 'Save as .xlsx' $false 'Skipped: workbook is null'
        }

        # ── 6. READ-BACK VALIDATION ───────────────────────────────────────────
        Write-Build Cyan "`n  [6/6] Read-back Validation"
        if ($null -ne $xlsxFile -and (Test-Path $xlsxFile)) {
            try {
                $wb2 = Open-ExcelWorkbook -Context $ctx -ExcelApp $xl -Path $xlsxFile -ReadOnly $true
                Add-Result 'Open-ExcelWorkbook opens saved file' ($null -ne $wb2)

                if ($null -ne $wb2) {
                    $ws3 = Get-ExcelSheet -Workbook $wb2 -Index 1
                    if ($null -ne $ws3) {
                        $used2 = Get-ExcelUsedRange -Sheet $ws3
                        $expectedRows = $rowCount + 1
                        Add-Result 'Read-back: row count matches' ($used2.Rows -eq $expectedRows) `
                            "Expected=$expectedRows ReadBack=$($used2.Rows)"
                    } else {
                        Add-Result 'Read-back: row count matches' $false 'Sheet null on read-back'
                    }
                }
            } catch {
                Write-BuildLog $ctx 'ERROR' "Read-back failed: $_"
                Add-Result 'Open-ExcelWorkbook opens saved file' $false "$_"
            } finally {
                if ($null -ne $ws3) { Invoke-ReleaseComObject $ws3; $ws3 = $null }
                if ($null -ne $wb2) { try { $wb2.Close($false) } catch {}; Invoke-ReleaseComObject $wb2; $wb2 = $null }
            }
        } else {
            Add-Result 'Read-back validation' $true 'Skipped: no xlsx file to read back' 
        }

    } catch {
        Write-BuildLog $ctx 'ERROR' "DIAG EXCEL failed: $_" -Detail $_.ScriptStackTrace
        $script:failures++
    } finally {
        # Release in reverse acquisition order: sheet -> workbook -> app
        if ($null -ne $ws3) { try { Invoke-ReleaseComObject $ws3 } catch {} }
        if ($null -ne $wb2) { try { $wb2.Close($false) } catch {}; try { Invoke-ReleaseComObject $wb2 } catch {} }
        if ($null -ne $ws)  { try { Invoke-ReleaseComObject $ws  } catch {} }
        if ($null -ne $wb)  { try { $wb.Close($false)  } catch {}; try { Invoke-ReleaseComObject $wb  } catch {} }
        if ($null -ne $xl)  { try { $xl.Quit()         } catch {}; try { Invoke-ReleaseComObject $xl  } catch {} }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
    }

    # ── SUMMARY ───────────────────────────────────────────────────────────────
    Write-Build Cyan "`n  ── DIAG EXCEL SUMMARY ───────────────────────────────────"
    $total  = $results.Count
    $passed = @($results | Where-Object { $_.Pass }).Count
    Write-Build White  "  Total  : $total  |  Passed: $passed  |  Failed: $failures"

    $outFile = Join-Path $ctx.Paths.Reports "diag_excel_results_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    try {
        @{ runId=$ctx.RunId; total=$total; passed=$passed; failed=$failures; checks=$results.ToArray() } |
            ConvertTo-Json -Depth 4 |
            ForEach-Object { [System.IO.File]::WriteAllText($outFile, $_, [System.Text.Encoding]::ASCII) }
        Write-Build Cyan "  Report: $outFile"
    } catch { Write-BuildLog $ctx 'WARN' "Report write failed: $_" }

    if ($failures -gt 0) {
        Write-RunResult -Context $ctx -Success $false -ErrorMsg "$failures check(s) failed"
        throw "DIAG EXCEL: $failures check(s) failed"
    }
    Write-RunResult -Context $ctx -Success $true
}
