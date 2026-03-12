#Requires -Version 5.1
# =============================================================================
# lib/Integrity.ps1
# AutoBuild v3.0 - Task file integrity verification.
#
# RESOLVES: TASK-02 (HIGH) — No integrity verification for task files.
#   In environments with multiple operators, a task file can be replaced or
#   modified on disk without the engine detecting it. This is a vector for
#   arbitrary code execution in an automation environment.
#
# DESIGN:
#   Two complementary mechanisms (use one or both):
#
#   1. SHA-256 Hash Registry (no PKI required)
#      - Get-TaskHash      : computes the SHA-256 of a task file
#      - Register-TaskHash : stores approved hashes in tasks.hash.json
#      - Test-TaskIntegrity: verifies a file matches its registered hash
#      - Update-TaskRegistry: re-registers all task files (post-deploy step)
#
#   2. Authenticode Signature (requires Code Signing certificate)
#      - Invoke-SignTask    : signs a task file with a code-signing cert
#      - Test-TaskSignature : verifies the Authenticode signature
#      - Both modes can be combined: hash check + signature check
#
# HASH REGISTRY FILE: tasks/tasks.hash.json
#   Format:
#   {
#     "task_sap_stock.ps1": {
#       "sha256": "abc123...",
#       "registeredAt": "2026-03-12T14:00:00+01:00",
#       "registeredBy": "jsmith"
#     }
#   }
#
# ENGINE INTEGRATION:
#   Call Invoke-LoadTask (Main.build.ps1) already handles lazy loading.
#   Add this before dot-sourcing in Invoke-LoadTask:
#     if ($Script:IntegrityEnabled) {
#         if (-not (Test-TaskIntegrity -FilePath $file -HashFile $Script:TaskHashFile)) {
#             throw "AutoBuild: integrity check failed for '$TaskName'. File may have been modified."
#         }
#     }
#
# SETUP WORKFLOW:
#   # After deploying approved tasks:
#   $cfg  = Get-EngineConfig -Root $EngineRoot
#   Update-TaskRegistry -TasksDir (Join-Path $EngineRoot 'tasks') `
#                       -HashFile  (Join-Path $EngineRoot 'tasks\tasks.hash.json')
#   # Protect the hash file: restrict write access to Admin-only accounts.
# =============================================================================
Set-StrictMode -Version Latest

# ============================================================================
# HASH-BASED VERIFICATION
# ============================================================================

function Get-TaskHash {
    <#
    .SYNOPSIS
        Computes the SHA-256 hash of a task file.
    .OUTPUTS
        Lowercase hex string (64 characters).
    #>
    param([Parameter(Mandatory)][string]$FilePath)

    if (-not (Test-Path $FilePath)) {
        throw "Integrity: file not found: $FilePath"
    }
    $bytes = [System.IO.File]::ReadAllBytes($FilePath)
    $sha   = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hash = $sha.ComputeHash($bytes)
        return ([BitConverter]::ToString($hash) -replace '-','').ToLower()
    } finally {
        $sha.Dispose()
    }
}

function Get-TaskHashRegistry {
    <#
    .SYNOPSIS
        Loads the task hash registry from disk.
    .OUTPUTS
        Hashtable: filename -> @{ sha256, registeredAt, registeredBy }
        Returns empty hashtable if file not found.
    #>
    param([Parameter(Mandatory)][string]$HashFile)

    if (-not (Test-Path $HashFile)) { return @{} }
    try {
        $raw = Get-Content $HashFile -Raw -Encoding ASCII | ConvertFrom-Json
        $reg = @{}
        foreach ($prop in $raw.PSObject.Properties) {
            $reg[$prop.Name] = @{
                sha256        = $prop.Value.sha256
                registeredAt  = $prop.Value.registeredAt
                registeredBy  = $prop.Value.registeredBy
            }
        }
        return $reg
    } catch {
        Write-Warning "Integrity: Cannot read hash registry '$HashFile': $_"
        return @{}
    }
}

function Register-TaskHash {
    <#
    .SYNOPSIS
        Adds or updates a single task file in the hash registry.
        Uses an atomic write (temp + rename) to prevent partial writes.
    .PARAMETER HashFile
        Path to tasks.hash.json.
    .PARAMETER FilePath
        Absolute path to the task .ps1 file.
    #>
    param(
        [Parameter(Mandatory)][string]$HashFile,
        [Parameter(Mandatory)][string]$FilePath
    )

    $fileName = [System.IO.Path]::GetFileName($FilePath)
    $hash     = Get-TaskHash -FilePath $FilePath
    $reg      = Get-TaskHashRegistry -HashFile $HashFile

    $reg[$fileName] = @{
        sha256       = $hash
        registeredAt = (Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')
        registeredBy = try { ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name -split '\\')[-1] } catch { $env:USERNAME }
    }

    Save-TaskHashRegistry -HashFile $HashFile -Registry $reg
    Write-Verbose "Integrity: registered '$fileName' (sha256=$hash)"
}

function Save-TaskHashRegistry {
    <#
    .SYNOPSIS
        Writes the registry hashtable to disk atomically.
    #>
    param(
        [Parameter(Mandatory)][string]$HashFile,
        [Parameter(Mandatory)][hashtable]$Registry
    )

    $ordered = [ordered]@{}
    foreach ($k in ($Registry.Keys | Sort-Object)) {
        $ordered[$k] = [ordered]@{
            sha256       = $Registry[$k].sha256
            registeredAt = $Registry[$k].registeredAt
            registeredBy = $Registry[$k].registeredBy
        }
    }
    $json = $ordered | ConvertTo-Json -Depth 3
    $tmp  = "$HashFile.tmp"
    try {
        [System.IO.File]::WriteAllText($tmp, $json, [System.Text.Encoding]::ASCII)
        Move-Item -Path $tmp -Destination $HashFile -Force -ErrorAction Stop
    } catch {
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        throw "Integrity: Cannot write hash registry: $_"
    }
}

function Test-TaskIntegrity {
    <#
    .SYNOPSIS
        Verifies a task file's hash matches the approved registry entry.
    .OUTPUTS
        $true if hash matches (or file not in registry and -AllowUnregistered).
        $false if file is unregistered (default: deny-unknown).
        Throws if file is registered but hash has changed.
    .PARAMETER AllowUnregistered
        If $true, files not in the registry are allowed (dev/migration mode).
        Default is $false (deny-unknown = production mode).
    #>
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string]$HashFile,
        [bool]$AllowUnregistered = $false
    )

    $fileName = [System.IO.Path]::GetFileName($FilePath)
    $reg      = Get-TaskHashRegistry -HashFile $HashFile

    if (-not $reg.ContainsKey($fileName)) {
        if ($AllowUnregistered) {
            Write-Warning "Integrity: '$fileName' is not in the hash registry (AllowUnregistered=true)."
            return $true
        }
        # Fail-safe: deny unknown tasks in production mode
        Write-Warning "Integrity: '$fileName' is NOT registered. Add it with Register-TaskHash."
        return $false
    }

    $expected = $reg[$fileName].sha256
    $actual   = Get-TaskHash -FilePath $FilePath

    if ($expected -ne $actual) {
        throw "Integrity: HASH MISMATCH for '$fileName'. " +
              "Expected: $expected. " +
              "Actual: $actual. " +
              "File may have been tampered with. Registration date: $($reg[$fileName].registeredAt)."
    }

    return $true
}

function Update-TaskRegistry {
    <#
    .SYNOPSIS
        Re-registers all task_*.ps1 files in $TasksDir.
        Excludes task_TEMPLATE.ps1. Run this after deploying approved task changes.
    .PARAMETER TasksDir
        Path to the tasks/ directory.
    .PARAMETER HashFile
        Path to tasks.hash.json. Created if not present.
    .OUTPUTS
        Count of files registered.
    #>
    param(
        [Parameter(Mandatory)][string]$TasksDir,
        [Parameter(Mandatory)][string]$HashFile
    )

    if (-not (Test-Path $TasksDir)) { throw "Integrity: tasks directory not found: $TasksDir" }

    $files = @(
        Get-ChildItem $TasksDir -Filter 'task_*.ps1' |
        Where-Object { $_.Name -ne 'task_TEMPLATE.ps1' }
    )
    $reg = @{}
    foreach ($f in $files) {
        $hash = Get-TaskHash -FilePath $f.FullName
        $reg[$f.Name] = @{
            sha256       = $hash
            registeredAt = (Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')
            registeredBy = try { ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name -split '\\')[-1] } catch { $env:USERNAME }
        }
        Write-Verbose "Integrity: registered $($f.Name) ($hash)"
    }
    Save-TaskHashRegistry -HashFile $HashFile -Registry $reg
    Write-Host "Integrity: registered $($files.Count) task files in '$HashFile'" -ForegroundColor Cyan
    return $files.Count
}

function Get-TaskRegistryReport {
    <#
    .SYNOPSIS
        Returns a status report comparing registry to actual files on disk.
        Useful for detecting: unregistered new files, deleted tasks still in registry,
        and modified (hash-changed) tasks.
    .OUTPUTS
        Array of PSCustomObject with: FileName, Status (OK|MODIFIED|UNREGISTERED|MISSING), Hash
    #>
    param(
        [Parameter(Mandatory)][string]$TasksDir,
        [Parameter(Mandatory)][string]$HashFile
    )

    $reg     = Get-TaskHashRegistry -HashFile $HashFile
    $onDisk  = @(Get-ChildItem $TasksDir -Filter 'task_*.ps1' |
                 Where-Object { $_.Name -ne 'task_TEMPLATE.ps1' })
    $results = [System.Collections.Generic.List[PSCustomObject]]::new()

    # Check registered files
    foreach ($f in $onDisk) {
        $status = 'UNREGISTERED'
        $actual = Get-TaskHash -FilePath $f.FullName
        if ($reg.ContainsKey($f.Name)) {
            $status = if ($reg[$f.Name].sha256 -eq $actual) { 'OK' } else { 'MODIFIED' }
        }
        $results.Add([PSCustomObject]@{
            FileName     = $f.Name
            Status       = $status
            ActualHash   = $actual
            RegisteredAt = if ($reg.ContainsKey($f.Name)) { $reg[$f.Name].registeredAt } else { '' }
            RegisteredBy = if ($reg.ContainsKey($f.Name)) { $reg[$f.Name].registeredBy } else { '' }
        })
    }

    # Check for registry entries with no corresponding file (deleted tasks)
    $onDiskNames = @($onDisk.Name)
    foreach ($name in $reg.Keys) {
        if ($onDiskNames -notcontains $name) {
            $results.Add([PSCustomObject]@{
                FileName     = $name
                Status       = 'MISSING'
                ActualHash   = ''
                RegisteredAt = $reg[$name].registeredAt
                RegisteredBy = $reg[$name].registeredBy
            })
        }
    }

    return $results.ToArray()
}

# ============================================================================
# AUTHENTICODE SIGNATURE-BASED VERIFICATION (requires Code Signing cert)
# ============================================================================

function Invoke-SignTask {
    <#
    .SYNOPSIS
        Signs a task file with an Authenticode code-signing certificate.
    .DESCRIPTION
        Requires a valid code-signing certificate in the current user's
        certificate store. The thumbprint can be specified; if omitted,
        the first valid code-signing cert is used.
    .PARAMETER FilePath
        Task .ps1 file to sign.
    .PARAMETER Thumbprint
        Certificate thumbprint. If empty, uses the first valid code-signing cert.
    .OUTPUTS
        The SignatureStatus string ('Valid' on success).
    #>
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [string]$Thumbprint = ''
    )

    if (-not (Test-Path $FilePath)) { throw "Integrity: file not found: $FilePath" }

    $cert = $null
    if ($Thumbprint) {
        $cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Thumbprint -eq $Thumbprint }
    } else {
        $cert = Get-ChildItem Cert:\CurrentUser\My |
                Where-Object { $_.EnhancedKeyUsageList.ObjectId -contains '1.3.6.1.5.5.7.3.3' } |
                Select-Object -First 1
    }

    if ($null -eq $cert) {
        throw "Integrity: no code-signing certificate found. " +
              "Install a certificate in Cert:\CurrentUser\My or specify -Thumbprint."
    }

    $result = Set-AuthenticodeSignature -FilePath $FilePath -Certificate $cert
    if ($result.Status -ne 'Valid') {
        throw "Integrity: signing failed for '$FilePath'. Status: $($result.Status)"
    }
    Write-Verbose "Integrity: signed '$FilePath' with cert $($cert.Thumbprint)"
    return $result.Status
}

function Test-TaskSignature {
    <#
    .SYNOPSIS
        Verifies the Authenticode signature of a task file.
    .OUTPUTS
        $true if signature is valid.
        $false if unsigned (when AllowUnsigned = $true in dev mode).
        Throws if signature is invalid or tampered.
    .PARAMETER AllowUnsigned
        If $true, unsigned files are allowed (dev mode only).
        Default $false: all tasks must be signed (production mode).
    #>
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [bool]$AllowUnsigned = $false
    )

    if (-not (Test-Path $FilePath)) { throw "Integrity: file not found: $FilePath" }

    $sig = Get-AuthenticodeSignature -FilePath $FilePath

    switch ($sig.Status) {
        'Valid'           { return $true }
        'NotSigned'       {
            if ($AllowUnsigned) {
                Write-Warning "Integrity: '$([System.IO.Path]::GetFileName($FilePath))' is not signed (AllowUnsigned=true)."
                return $true
            }
            throw "Integrity: '$([System.IO.Path]::GetFileName($FilePath))' is not signed. Sign it with Invoke-SignTask before deployment."
        }
        'HashMismatch'    { throw "Integrity: '$([System.IO.Path]::GetFileName($FilePath))' has a HASH MISMATCH - file was modified after signing." }
        'NotTrusted'      { throw "Integrity: '$([System.IO.Path]::GetFileName($FilePath))' signature is NOT TRUSTED - certificate is not in TrustedPublisher store." }
        'UnknownError'    { throw "Integrity: '$([System.IO.Path]::GetFileName($FilePath))' unknown signature error." }
        default           { throw "Integrity: '$([System.IO.Path]::GetFileName($FilePath))' unexpected signature status: $($sig.Status)" }
    }
}
