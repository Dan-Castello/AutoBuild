#Requires -Version 5.1
# =============================================================================
# tasks/task_diag_notify.ps1
# @Description : Notifications + observability diagnostics - SMTP probe, log correlation, metrics
# @Category    : Utility
# @Version     : 1.1.0
# @Author      : AutoBuild QA
# =============================================================================
# Synopsis: Diagnostics - SMTP config, log correlation (RunId), run summaries, log rotation
# Params: {"SendTestEmail":"false"}
#
# REMEDIATION v3.1 - FIX-SMTP-OPTIONAL:
#   Root cause: 'SMTP server is configured' was treated as a REQUIRED check.
#   When smtpServer is empty the check failed, $failures was incremented, and
#   the task threw "DIAG NOTIFY: N check(s) failed" — even though SMTP is an
#   OPTIONAL feature. The diagnostic reported a phantom failure.
#
#   Fix: Call Test-EngineConfiguration to classify SMTP as OPTIONAL-MISSING.
#   When SmtpConfigured = $false, the SMTP section records an INFO-level pass
#   ("Notifications disabled: smtpServer not configured") instead of a failure.
#   All downstream SMTP sub-checks are skipped. The task passes unless a real
#   failure occurs in the log-correlation, rotation, or purge sections.

task diag_notify {

    $ctx = New-TaskContext `
        -TaskName 'diag_notify' `
        -Config   $Script:EngineConfig `
        -Root     $Script:EngineRoot `
        -Params   $Script:TaskParams

    $sendEmail = try { [bool]::Parse($ctx.Params['SendTestEmail']) } catch { $false }

    Write-BuildLog $ctx 'INFO' 'DIAG NOTIFY: Starting notifications + observability diagnostics'

    $results  = [System.Collections.Generic.List[hashtable]]::new()
    $failures = 0
    $script:failures = 0
    $stamp    = Get-Date -Format 'yyyyMMdd_HHmmss'

    function Add-Result {
        param([string]$Check, [bool]$Pass, [string]$Detail = '', [bool]$Optional = $false)
        $results.Add(@{ Check = $Check; Pass = $Pass; Detail = $Detail; Optional = $Optional })
        $sym   = if ($Pass) { '[OK]' } else { if ($Optional) { '[SKIP]' } else { '[FAIL]' } }
        $lvl   = if ($Pass) { 'INFO' } else { if ($Optional) { 'INFO' } else { 'ERROR' } }
        $color = if ($Pass) { 'Green' } else { if ($Optional) { 'Yellow' } else { 'Red' } }
        Write-BuildLog $ctx $lvl "$sym $Check" -Detail $Detail
        Write-Build $color "  $sym  $Check$(if($Detail){' -- '+$Detail})"
        # ONLY increment failures for non-optional, non-passing checks.
        if (-not $Pass -and -not $Optional) { $script:failures++ }
    }

    # ─── 1. SMTP CONFIGURATION (OPTIONAL) ────────────────────────────────────
    Write-Build Cyan "`n  [1/5] SMTP Configuration"

    # FIX-SMTP-OPTIONAL: Use Test-EngineConfiguration to classify SMTP correctly.
    $cfgValidation = Test-EngineConfiguration -Config $ctx.Config -Root $Script:EngineRoot

    if (-not $cfgValidation.SmtpConfigured) {
        # SMTP is intentionally unconfigured. Report as informational, NOT a failure.
        $smtpDisabledMsg = ($cfgValidation.OptionalIssues | Where-Object { $_ -match 'Notifications disabled' } | Select-Object -First 1)
        if (-not $smtpDisabledMsg) { $smtpDisabledMsg = 'Notifications disabled: smtpServer not configured' }

        Add-Result 'SMTP configuration' $true $smtpDisabledMsg -Optional $true
        Write-Build Yellow "  [INFO] $smtpDisabledMsg"
        Write-Build Yellow '  SMTP sub-checks skipped (SMTP is an optional feature).'
        Write-BuildLog $ctx 'INFO' "SMTP optional: $smtpDisabledMsg"

    } else {
        # SMTP IS configured — run full sub-checks.
        $smtpCfg = $ctx.Config.notifications
        Add-Result 'SMTP server is configured' $true $smtpCfg.smtpServer

        $port = try { [int]$smtpCfg.smtpPort } catch { 0 }
        Add-Result 'SMTP port is valid (1-65535)' ($port -ge 1 -and $port -le 65535) "Port=$port"

        $fromOk = -not [string]::IsNullOrWhiteSpace("$($smtpCfg.fromAddr)")
        $toOk   = -not [string]::IsNullOrWhiteSpace("$($smtpCfg.toAddr)")
        Add-Result 'SMTP fromAddr is configured' $fromOk "$($smtpCfg.fromAddr)"
        Add-Result 'SMTP toAddr is configured'   $toOk   "$($smtpCfg.toAddr)"

        # TCP reachability check (no actual email sent)
        try {
            $tcp = New-Object System.Net.Sockets.TcpClient
            $ar  = $tcp.BeginConnect($smtpCfg.smtpServer, $port, $null, $null)
            $ok  = $ar.AsyncWaitHandle.WaitOne(3000)
            if ($ok) { try { $tcp.EndConnect($ar) } catch {} }
            try { $tcp.Close() } catch {}
            Add-Result "TCP reachable: $($smtpCfg.smtpServer):$port" $ok `
                $(if(-not $ok){'Connection refused or timeout'}else{'Connected'})
        } catch {
            Add-Result "TCP reachable: $($smtpCfg.smtpServer):$port" $false "$_"
        }

        if ($sendEmail) {
            Write-Build Yellow '  Sending test email (SendTestEmail=true)...'
            $sent = Test-NotificationConfig -Config $ctx.Config
            Add-Result 'Test email delivered successfully' $sent `
                $(if(-not $sent){'Check SMTP credentials and relay config'})
        } else {
            Add-Result 'Test email send skipped (SendTestEmail=false)' $true 'Set -Params {"SendTestEmail":"true"} to send'
        }
    }

    # Report any additional optional issues from config validation
    foreach ($optIssue in $cfgValidation.OptionalIssues) {
        if ($optIssue -notmatch 'Notifications disabled') {
            Write-BuildLog $ctx 'INFO' "Config optional: $optIssue"
        }
    }
    foreach ($critIssue in $cfgValidation.CriticalIssues) {
        Add-Result "Config critical: $critIssue" $false $critIssue
    }

    # Send-Notification function exists check
    $fnExists = $null -ne (Get-Command 'Send-Notification' -ErrorAction SilentlyContinue)
    Add-Result 'Send-Notification function available' $fnExists `
        $(if(-not $fnExists){'Load lib/Notifications.ps1 in engine load order'})

    # ─── 2. LOG CORRELATION (RunId tracking) ─────────────────────────────────
    Write-Build Cyan "`n  [2/5] Log Correlation (RunId)"

    Write-BuildLog $ctx 'INFO'  'CORRELATION_TEST: entry 1 of 5'
    Write-BuildLog $ctx 'INFO'  'CORRELATION_TEST: entry 2 of 5'
    Write-BuildLog $ctx 'WARN'  'CORRELATION_TEST: simulated warning'
    Write-BuildLog $ctx 'DEBUG' 'CORRELATION_TEST: debug detail'
    Write-BuildLog $ctx 'INFO'  'CORRELATION_TEST: entry 5 of 5'

    $logFile = Join-Path $ctx.Paths.Logs 'registry.jsonl'
    if (Test-Path $logFile) {
        $allLines = @(Get-Content $logFile -Encoding ASCII -ErrorAction SilentlyContinue)
        $thisRunLines = @($allLines | ForEach-Object {
            try { $o = $_ | ConvertFrom-Json; if ($o.runId -eq $ctx.RunId) { $o } } catch {}
        })

        Add-Result 'All entries share this RunId' ($thisRunLines.Count -ge 5) `
            "Found=$($thisRunLines.Count) entries with RunId=$($ctx.RunId)"

        $levels = @($thisRunLines | Select-Object -ExpandProperty level) | Sort-Object -Unique
        Add-Result 'Multiple log levels recorded for this run' ($levels.Count -ge 2) `
            "Levels: $($levels -join ', ')"

        $badTs = @($thisRunLines | Where-Object {
            try { [datetime]::Parse($_.ts) | Out-Null; $false } catch { $true }
        })
        Add-Result 'All timestamps are valid ISO 8601' ($badTs.Count -eq 0) `
            "Invalid timestamps: $($badTs.Count)"

        $missingUser = @($thisRunLines | Where-Object { [string]::IsNullOrWhiteSpace($_.user) })
        $missingHost = @($thisRunLines | Where-Object { [string]::IsNullOrWhiteSpace($_.hostname) })
        Add-Result 'All entries have user field'     ($missingUser.Count -eq 0) "Missing=$($missingUser.Count)"
        Add-Result 'All entries have hostname field' ($missingHost.Count -eq 0) "Missing=$($missingHost.Count)"
    }

    # ─── 3. LOG ROTATION SIMULATION ──────────────────────────────────────────
    Write-Build Cyan "`n  [3/5] Log Rotation"
    $rotDir  = Join-Path $ctx.Paths.Logs 'rotation_test'
    $rotFile = Join-Path $rotDir 'registry.jsonl'
    New-Item $rotDir -ItemType Directory -Force | Out-Null

    try {
        $bigContent = [string]::new('A', 100)
        [System.IO.File]::WriteAllText($rotFile, $bigContent, [System.Text.Encoding]::ASCII)
        $sizeBefore = (Get-Item $rotFile).Length

        Invoke-LogRotationIfNeeded -FilePath $rotFile -MaxBytes 10
        $rotatedFiles = @(Get-ChildItem $rotDir -Filter 'registry_*.jsonl' -ErrorAction SilentlyContinue)

        Add-Result 'File > threshold triggers rotation' ($rotatedFiles.Count -ge 1) `
            "Size before=$sizeBefore rotated files=$($rotatedFiles.Count)"
        Add-Result 'Active log path vacated after rotation' (-not (Test-Path $rotFile)) `
            'registry.jsonl should be gone after rename'

    } catch {
        Add-Result 'Log rotation simulation' $false "$_"
    } finally {
        Remove-Item $rotDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    # ─── 4. LOG PURGE ─────────────────────────────────────────────────────────
    Write-Build Cyan "`n  [4/5] Log Purge (retentionDays)"
    $purgeDir = Join-Path $ctx.Paths.Logs 'purge_test'
    New-Item $purgeDir -ItemType Directory -Force | Out-Null
    try {
        $oldFile1 = Join-Path $purgeDir "registry_20200101_000000.jsonl"
        $oldFile2 = Join-Path $purgeDir "registry_20200601_000000.jsonl"
        $newFile  = Join-Path $purgeDir "registry_$(Get-Date -Format 'yyyyMMdd_HHmmss').jsonl"
        $active   = Join-Path $purgeDir "registry.jsonl"

        foreach ($f in @($oldFile1,$oldFile2,$newFile,$active)) {
            [System.IO.File]::WriteAllText($f, 'test', [System.Text.Encoding]::ASCII)
        }
        (Get-Item $oldFile1).LastWriteTime = [datetime]'2020-01-01'
        (Get-Item $oldFile2).LastWriteTime = [datetime]'2020-06-01'

        Invoke-LogPurge -LogsDir $purgeDir -RetentionDays 30

        Add-Result 'Old archive 1 purged (2020-01-01)'     (-not (Test-Path $oldFile1))
        Add-Result 'Old archive 2 purged (2020-06-01)'     (-not (Test-Path $oldFile2))
        Add-Result 'Recent archive preserved'              (Test-Path $newFile)
        Add-Result 'Active registry.jsonl never purged'    (Test-Path $active)

    } catch {
        Add-Result 'Log purge simulation' $false "$_"
    } finally {
        Remove-Item $purgeDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    # ─── 5. RUN SUMMARY GENERATION ───────────────────────────────────────────
    Write-Build Cyan "`n  [5/5] Run Summary Generation"
    $summaryCtx = New-TaskContext -TaskName 'summary_probe' -Config $Script:EngineConfig -Root $Script:EngineRoot
    Write-BuildLog $summaryCtx 'INFO'  'Summary probe: step 1'
    Write-BuildLog $summaryCtx 'INFO'  'Summary probe: step 2'
    Write-BuildLog $summaryCtx 'WARN'  'Summary probe: warning'
    Write-RunResult -Context $summaryCtx -Success $true

    $logLines200 = @(Get-Content $logFile -Last 200 -Encoding ASCII -ErrorAction SilentlyContinue)
    $grouped = @($logLines200 | ForEach-Object {
        try { $_ | ConvertFrom-Json } catch {}
    } | Group-Object runId | Where-Object { $_.Name -eq $summaryCtx.RunId })

    Add-Result 'Summary probe run found in last 200 entries' ($grouped.Count -gt 0) `
        "RunId=$($summaryCtx.RunId) entries=$($grouped | Select-Object -ExpandProperty Count -First 1)"

    $okEntry = @($grouped.Group | Where-Object { $_.level -eq 'OK' })
    Add-Result 'Summary probe has terminal OK entry' ($okEntry.Count -gt 0)

    $elapsed = @($grouped.Group | Where-Object {
        $null -ne $_.PSObject.Properties['elapsed'] -and $null -ne $_.elapsed
    })
    Add-Result 'Summary probe has elapsed time recorded' ($elapsed.Count -gt 0) `
        "elapsed=$($elapsed | Select-Object -ExpandProperty elapsed -First 1)s"

    # ─── SUMMARY ──────────────────────────────────────────────────────────────
    Write-Build Cyan "`n  ── DIAG NOTIFY SUMMARY ──────────────────────────────────"
    $total  = $results.Count
    $passed = @($results | Where-Object { $_.Pass }).Count
    Write-Build White "  Total: $total  |  Passed: $passed  |  Failed: $failures"

    $outFile = Join-Path $ctx.Paths.Reports "diag_notify_${stamp}.json"
    try {
        @{ runId=$ctx.RunId; total=$total; passed=$passed; failed=$failures; checks=$results.ToArray() } |
            ConvertTo-Json -Depth 4 |
            ForEach-Object { [System.IO.File]::WriteAllText($outFile, $_, [System.Text.Encoding]::ASCII) }
        Write-Build Cyan "  Report: $outFile"
    } catch {}

    if ($failures -gt 0) {
        Write-RunResult -Context $ctx -Success $false -ErrorMsg "$failures check(s) failed"
        throw "DIAG NOTIFY: $failures check(s) failed"
    }
    Write-RunResult -Context $ctx -Success $true
}
