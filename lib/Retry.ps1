#Requires -Version 5.1
# =============================================================================
# lib/Retry.ps1
# AutoBuild v2.0 - Retry infrastructure with exponential backoff.
#
# RESOLVES:
#   PROBLEMA-EXT-01 / F3-03 : engine.config.json has maxRetries and
#     retryDelaySeconds, but v1 never implemented retry logic in the
#     actual operations. Invoke-WithRetry fills that gap.
#   PROBLEMA-SAP-01 (HIGH)  : SAP transient errors (session busy, network
#     blip) are now automatically retried instead of failing hard.
#   PROBLEMA-COM-01 (HIGH)  : COM errors that are transient (e.g. Excel
#     starting up) can be wrapped in Invoke-WithRetry at the call site.
#
# USAGE:
#   $result = Invoke-WithRetry -Context $ctx -Label 'Excel open' -ScriptBlock {
#       Open-ExcelWorkbook -Context $ctx -ExcelApp $xl -Path $file
#   }
# =============================================================================
Set-StrictMode -Version Latest

function Invoke-WithRetry {
    <#
    .SYNOPSIS
        Executes $ScriptBlock with automatic retry and exponential backoff.
    .PARAMETER Context
        Task context. MaxRetries and initial delay read from Config.engine.
    .PARAMETER ScriptBlock
        Block to execute. May throw; exceptions trigger retry.
    .PARAMETER Label
        Short label used in log messages.
    .PARAMETER MaxRetries
        Override the config default (optional).
    .PARAMETER BaseDelaySeconds
        Initial wait before first retry. Each subsequent retry doubles it.
    .OUTPUTS
        Whatever $ScriptBlock returns on success.
    .NOTES
        Jitter: +-25% random noise on each delay prevents thundering herd
        when multiple parallel tasks hit the same resource simultaneously.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Context,
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [string]$Label            = 'operation',
        [int]$MaxRetries          = -1,    # -1 = read from config
        [double]$BaseDelaySeconds = -1     # -1 = read from config
    )

    if ($MaxRetries   -lt 0) { $MaxRetries   = [int]$Context.Config.engine.maxRetries }
    if ($BaseDelaySeconds -lt 0) { $BaseDelaySeconds = [double]$Context.Config.engine.retryDelaySeconds }

    $attempt = 0
    $lastErr  = $null
    $delay    = $BaseDelaySeconds

    while ($attempt -le $MaxRetries) {
        try {
            return (& $ScriptBlock)
        } catch {
            $lastErr = $_
            if ($attempt -ge $MaxRetries) { break }

            # Exponential backoff with +-25% jitter
            $jitter    = $delay * (0.75 + (Get-Random -Minimum 0 -Maximum 50) / 100.0)
            $waitMs    = [int]($jitter * 1000)
            $jitterStr = '{0:F1}' -f $jitter
            $attempt++
            Write-BuildLog $Context 'WARN' "[$Label] attempt $attempt/$MaxRetries failed: $($lastErr.Exception.Message). Retrying in ${jitterStr}s."
            Start-Sleep -Milliseconds $waitMs
            $delay = $delay * 2   # double for next round
        }
    }

    Write-BuildLog $Context 'ERROR' "[$Label] exhausted $MaxRetries retries. Last error: $lastErr"
    throw $lastErr
}
