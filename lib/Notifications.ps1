#Requires -Version 5.1
# =============================================================================
# lib/Notifications.ps1
# AutoBuild v3.1 - SMTP notification module.
#
# FIX SMTP-MISSING (AUDIT v3 MEDIUM):
#   engine.config.json had complete SMTP configuration (smtpServer, smtpPort,
#   fromAddr, toAddr) but no Send-Notification function existed anywhere in the
#   codebase. This created false expectations: operators assumed notifications
#   were sent on failure when they were not.
#
# DESIGN:
#   Uses System.Net.Mail.SmtpClient (built into .NET 4.x, no external deps).
#   Compatible with PS 5.1 and corporate environments without internet access.
#   Supports optional SMTP authentication and TLS (StartTLS on port 587).
#
# CONFIGURATION (engine.config.json):
#   "notifications": {
#     "smtpServer"   : "mailrelay.corp.local",
#     "smtpPort"     : 25,
#     "fromAddr"     : "autobuild@corp.local",
#     "toAddr"       : "ops-team@corp.local",
#     "smtpUser"     : "",        <- optional SMTP auth user
#     "smtpPassword" : "",        <- optional SMTP auth password (cleartext; use service account)
#     "useTls"       : false      <- set true for port 587/465 with STARTTLS
#   }
#
# SECURITY NOTE:
#   smtpPassword in config is cleartext. For production, either:
#   a) Use a relay that accepts unauthenticated connections from trusted IPs, or
#   b) Store the password encrypted with DPAPI and decrypt here.
#   Cleartext password is acceptable only for internal relay-only scenarios.
#
# ENGINE INTEGRATION:
#   Add to Main.build.ps1 lib load order after Logger.ps1:
#       'Notifications.ps1'
#   Then call at task completion / failure from task files or QueueRunner:
#       if (-not $success) {
#           Send-Notification -Context $ctx -Subject "FAILED: $($ctx.TaskName)" `
#               -Body "Run $($ctx.RunId) failed. See logs for details."
#       }
# =============================================================================
Set-StrictMode -Version Latest

function Send-Notification {
    <#
    .SYNOPSIS
        Sends an SMTP email notification using engine.config.json settings.
        Returns $true on success, $false on failure (non-fatal: errors are logged).
    .PARAMETER Context
        Task execution context (provides Config and logging).
    .PARAMETER Subject
        Email subject line. Prefixed with '[AutoBuild]' automatically.
    .PARAMETER Body
        Plain-text email body.
    .PARAMETER ToAddr
        Override recipient(s). Defaults to Config.notifications.toAddr.
    .PARAMETER Priority
        Normal | High | Low. Default: Normal.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Context,
        [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)][string]$Body,
        [string]$ToAddr   = '',
        [ValidateSet('Normal','High','Low')]
        [string]$Priority = 'Normal'
    )

    $cfg = $Context.Config.notifications
    if ($null -eq $cfg) {
        Write-BuildLog $Context 'WARN' 'Send-Notification: notifications section missing from config. Skipping.'
        return $false
    }

    $server = $cfg.smtpServer
    if ([string]::IsNullOrWhiteSpace($server)) {
        Write-BuildLog $Context 'WARN' 'Send-Notification: smtpServer not configured. Skipping notification.'
        return $false
    }

    $port = try { [int]$cfg.smtpPort } catch { 25 }
    if ($port -le 0) { $port = 25 }

    $from = if ([string]::IsNullOrWhiteSpace($cfg.fromAddr)) { "autobuild@$env:COMPUTERNAME" } else { $cfg.fromAddr.Trim() }
    $to   = if ([string]::IsNullOrWhiteSpace($ToAddr)) { $cfg.toAddr } else { $ToAddr }
    $to   = $to.Trim()

    if ([string]::IsNullOrWhiteSpace($to)) {
        Write-BuildLog $Context 'WARN' 'Send-Notification: toAddr not configured. Skipping notification.'
        return $false
    }

    $fullSubject = "[AutoBuild] $Subject"
    $fullBody    = @(
        $Body,
        '',
        '---',
        "Task     : $($Context.TaskName)",
        "RunId    : $($Context.RunId)",
        "User     : $($Context.User)",
        "Host     : $($Context.Hostname)",
        "Time     : $(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')"
    ) -join "`n"

    $smtp = $null
    try {
        $smtp = New-Object System.Net.Mail.SmtpClient($server, $port)
        $smtp.DeliveryMethod = [System.Net.Mail.SmtpDeliveryMethod]::Network

        # TLS support
        $useTls = try { [bool]$cfg.useTls } catch { $false }
        if ($useTls) { $smtp.EnableSsl = $true }

        # SMTP authentication (optional)
        $smtpUser = $cfg.smtpUser
        $smtpPass = $cfg.smtpPassword
        if (-not [string]::IsNullOrWhiteSpace($smtpUser) -and
            -not [string]::IsNullOrWhiteSpace($smtpPass)) {
            $smtp.Credentials = New-Object System.Net.NetworkCredential($smtpUser, $smtpPass)
        }

        $mail          = New-Object System.Net.Mail.MailMessage($from, $to, $fullSubject, $fullBody)
        $mail.Priority = [System.Net.Mail.MailPriority]$Priority

        $smtp.Send($mail)
        Write-BuildLog $Context 'INFO' "Notification sent to '$to': $fullSubject"
        return $true

    } catch {
        Write-BuildLog $Context 'WARN' "Send-Notification: SMTP delivery failed: $_"
        return $false
    } finally {
        if ($null -ne $smtp) { try { $smtp.Dispose() } catch { } }
    }
}

function Send-TaskFailureNotification {
    <#
    .SYNOPSIS
        Convenience wrapper: sends a standardised failure notification.
        Call from a task's catch block or from QueueRunner on task failure.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Context,
        [string]$ErrorMessage = ''
    )

    $subject    = "FAILED: $($Context.TaskName)"
    # PS 5.1: cannot use 'if' inline inside @() array literal — build body as string directly.
    $errorLine  = if ($ErrorMessage) { "Error: $ErrorMessage" } else { 'No error detail captured.' }
    $body       = "Task '$($Context.TaskName)' failed during automated execution.`n`n$errorLine"

    return Send-Notification -Context $Context -Subject $subject -Body $body -Priority High
}

function Test-NotificationConfig {
    <#
    .SYNOPSIS
        Validates SMTP configuration by sending a test message.
        Returns $true if test message was delivered.
    .PARAMETER Config
        Engine configuration hashtable.
    #>
    param([Parameter(Mandatory)][hashtable]$Config)

    # PS 5.1: hashtable values cannot contain try/catch expressions inline.
    # Resolve the user name before constructing the hashtable.
    $diagUser = try {
        ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name -split '\\')[-1]
    } catch { $env:USERNAME }

    $tempCtx = @{
        TaskName  = 'DiagNotification'
        RunId     = 'DIAG_TEST'
        User      = $diagUser
        Hostname  = $env:COMPUTERNAME
        Config    = $Config
        Paths     = @{ Logs = $env:TEMP }
        StartTime = [datetime]::Now
    }

    return Send-Notification -Context $tempCtx `
        -Subject 'Notification Test' `
        -Body 'This is an automated test of the AutoBuild notification system.'
}
