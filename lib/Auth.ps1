#Requires -Version 5.1
# =============================================================================
# lib/Auth.ps1
# AutoBuild v3.1 - Role-Based Access Control with AD group validation.
#
# AUDIT v3 FIXES:
#   FIX R-03 (HIGH)  : Resolve-UserRole no longer silently degrades all users
#                      to Operator when the DC is unreachable. A DPAPI-encrypted
#                      role cache (user-scoped) is used as fallback with TTL.
#   FIX V-04 (HIGH)  : Assert-SecurityConfigPopulated added. Call on startup to
#                      detect dev-mode deployments with empty security groups.
#   FIX PROD-GUARD   : Added Assert-SecurityConfigPopulated for engine startup guard.
# =============================================================================

# DPAPI is in System.Security. Available in .NET 4.x on PS 5.1 Desktop edition.
Add-Type -AssemblyName System.Security
# RESOLVES:
#   PROBLEMA-SEC-01 (CRITICAL) : RBAC was purely decorative - any user could
#     claim -Role Admin on the CLI. Auth.ps1 validates the claimed role
#     against Active Directory group membership (or a config-file allowlist
#     when AD integration is not configured).
#   PROBLEMA-RBAC-01 (HIGH)    : Fn_TestPermission defaulted to $true for
#     unknown actions (fail-open). This module defaults to $false (fail-safe).
#   PROBLEMA-SEC-02 (HIGH)     : User identity is read from WindowsIdentity,
#     not $env:USERNAME.
#
# HOW IT WORKS:
#   1. Resolve-UserRole is called at UI startup (or Run.ps1 launch).
#      It returns the HIGHEST role the current user qualifies for.
#   2. The role assignment is authoritative: the -Role CLI parameter is
#      treated only as a REQUESTED role, not a granted one.
#      If the user requests Admin but only qualifies for Operator, they
#      get Operator.
#   3. Test-Permission evaluates a requested action against the resolved role.
#      Unknown actions return $false (fail-safe).
#
# AD CONFIGURATION (engine.config.json):
#   "security": {
#     "adminAdGroup"     : "CN=AutoBuild-Admins,OU=Groups,DC=corp,DC=local",
#     "developerAdGroup" : "CN=AutoBuild-Devs,OU=Groups,DC=corp,DC=local",
#     "adminUsers"       : "jsmith,mgarcia",   <- fallback whitelist
#     "developerUsers"   : "alopez,bchen"
#   }
#   If both AD group and user whitelist are empty for a level, that level
#   is unreachable (only Operator is available to any domain user).
# =============================================================================
Set-StrictMode -Version Latest

# Action -> minimum required role (fail-safe: unlisted actions = denied)
$Script:ActionRoleMap = @{
    'RunTask'            = 'Operator'
    'ViewHistory'        = 'Operator'
    'ViewArtifacts'      = 'Operator'
    'ViewMetrics'        = 'Operator'
    'ViewDiag'           = 'Operator'
    'CreateTask'         = 'Developer'
    'EditTask'           = 'Developer'
    'EditConfig'         = 'Admin'
    'DeleteArtifact'     = 'Admin'
    'ViewAudit'          = 'Admin'
    'ManageCheckpoints'  = 'Admin'
    'PurgeOldLogs'       = 'Admin'
}

$Script:RoleHierarchy = @{ Operator = 0; Developer = 1; Admin = 2 }

# ---------------------------------------------------------------------------
# FIX R-03: Role cache with DPAPI encryption and configurable TTL.
# When AD is unreachable, the last resolved role is used instead of
# silently degrading all users to Operator.
# ---------------------------------------------------------------------------
$Script:RoleCache = @{
    FilePath   = ''     # Populated in Initialize-RoleCache
    TTLMinutes = 480    # Default: 8 hours (override via Config.security.roleCacheTTLMinutes)
}

function Initialize-RoleCache {
    <#
    .SYNOPSIS Initialises the role cache file path and TTL from config. #>
    param([hashtable]$Config, [string]$CacheDir)
    $Script:RoleCache.FilePath = Join-Path $CacheDir 'rolecache.dpapi'
    $ttl = try { [int]$Config.security.roleCacheTTLMinutes } catch { 0 }
    if ($ttl -gt 0) { $Script:RoleCache.TTLMinutes = $ttl }
}

function Read-RoleCache {
    <#
    .SYNOPSIS
        Reads and decrypts the role cache for the current user.
        Returns $null if cache is missing, expired, or decryption fails.
    #>
    param([string]$UserName)
    if ([string]::IsNullOrEmpty($Script:RoleCache.FilePath)) { return $null }
    if (-not (Test-Path $Script:RoleCache.FilePath)) { return $null }
    try {
        $raw     = [System.IO.File]::ReadAllBytes($Script:RoleCache.FilePath)
        $plain   = [System.Security.Cryptography.ProtectedData]::Unprotect(
                       $raw, $null,
                       [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
        $json    = [System.Text.Encoding]::UTF8.GetString($plain)
        $obj     = $json | ConvertFrom-Json
        # Validate: same user, not expired
        if ($obj.user -ne $UserName) { return $null }
        $age = ([datetime]::Now - [datetime]$obj.cachedAt).TotalMinutes
        if ($age -gt $Script:RoleCache.TTLMinutes) { return $null }
        return $obj.role
    } catch { return $null }
}

function Write-RoleCache {
    <#
    .SYNOPSIS Encrypts and persists the resolved role for the current user. #>
    param([string]$UserName, [string]$Role)
    if ([string]::IsNullOrEmpty($Script:RoleCache.FilePath)) { return }
    try {
        $obj   = [ordered]@{ user = $UserName; role = $Role; cachedAt = (Get-Date -Format 'o') }
        $json  = $obj | ConvertTo-Json -Compress
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
        $enc   = [System.Security.Cryptography.ProtectedData]::Protect(
                     $bytes, $null,
                     [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
        $dir   = Split-Path $Script:RoleCache.FilePath -Parent
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory $dir -Force | Out-Null }
        [System.IO.File]::WriteAllBytes($Script:RoleCache.FilePath, $enc)
    } catch {
        Write-Verbose "Auth: Could not write role cache: $_"
    }
}

function Assert-SecurityConfigPopulated {
    <#
    .SYNOPSIS
        FIX PROD-GUARD: Emits a warning (or error) when the engine starts
        with all security identifiers empty — effectively 'dev mode' for all users.
        Call this after loading config in Main.build.ps1 and AutoBuild.UI.ps1.
    .PARAMETER Config       Engine configuration hashtable.
    .PARAMETER ErrorOnEmpty Throw instead of warn (use in CI / strict production).
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [switch]$ErrorOnEmpty
    )
    $sec = $Config.security
    $adminEmpty = [string]::IsNullOrWhiteSpace($sec.adminAdGroup) -and
                  [string]::IsNullOrWhiteSpace($sec.adminUsers)
    $devEmpty   = [string]::IsNullOrWhiteSpace($sec.developerAdGroup) -and
                  [string]::IsNullOrWhiteSpace($sec.developerUsers)

    if ($adminEmpty -or $devEmpty) {
        $msg = @(
            '',
            '  ************************************************************',
            '  *  AutoBuild SECURITY WARNING                               *',
            '  *  engine.config.json security groups are empty.            *',
            '  *  ALL users will be resolved as Operator (no Admin/Dev).   *',
            '  *  Populate adminAdGroup or adminUsers before production.    *',
            '  ************************************************************',
            ''
        ) -join "`n"
        if ($ErrorOnEmpty) { throw $msg }
        Write-Warning $msg
    }
}

function Resolve-UserRole {
    <#
    .SYNOPSIS
        Determines the highest role the current Windows user qualifies for.
    .PARAMETER Config
        Engine configuration (must contain 'security' section).
    .PARAMETER RequestedRole
        The role the user CLAIMED via -Role parameter. Acts as a ceiling:
        even if AD qualifies for Admin, passing -Role Operator returns Operator.
    .OUTPUTS
        String: 'Admin' | 'Developer' | 'Operator'
    .NOTES
        FIX R-03: When AD (DC) is unreachable and IsInRole() fails, the
        function now uses a DPAPI-encrypted role cache instead of silently
        downgrading all users to Operator.
        Cache TTL is controlled by Config.security.roleCacheTTLMinutes (default 480).
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [ValidateSet('Operator','Developer','Admin')]
        [string]$RequestedRole = 'Operator'
    )

    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $userName = ($identity.Name -split '\\')[-1].ToLower()  # strip domain prefix
    $sec      = $Config.security

    # Helper: check AD group membership
    function Test-InAdGroup {
        param([string]$GroupDn)
        if ([string]::IsNullOrWhiteSpace($GroupDn)) { return $false }
        try {
            $principal = [System.Security.Principal.WindowsPrincipal]$identity
            return $principal.IsInRole($GroupDn)
        } catch {
            return $null   # FIX R-03: return $null to signal AD failure vs. genuine non-member
        }
    }

    # Helper: check whitelist
    function Test-InWhitelist {
        param([string]$Whitelist)
        if ([string]::IsNullOrWhiteSpace($Whitelist)) { return $false }
        $list = $Whitelist -split '[,;\s]+' | ForEach-Object { $_.Trim().ToLower() } | Where-Object { $_ }
        return $list -contains $userName
    }

    # Determine the highest qualified role, tracking AD availability
    $qualifiedRole = 'Operator'   # Everyone is at minimum an Operator
    $adUnavailable = $false

    $isDevAd  = Test-InAdGroup $sec.developerAdGroup
    $isAdmAd  = Test-InAdGroup $sec.adminAdGroup
    if ($null -eq $isDevAd -or $null -eq $isAdmAd) { $adUnavailable = $true }

    if (($isDevAd -eq $true) -or (Test-InWhitelist $sec.developerUsers)) {
        $qualifiedRole = 'Developer'
    }
    if (($isAdmAd -eq $true) -or (Test-InWhitelist $sec.adminUsers)) {
        $qualifiedRole = 'Admin'
    }

    # FIX R-03: If AD was unreachable and only whitelist resolution was used,
    # check whether the whitelist alone resolved a privileged role.
    # If not, try the role cache before degrading to Operator.
    if ($adUnavailable -and $qualifiedRole -eq 'Operator') {
        $cached = Read-RoleCache -UserName $userName
        if ($null -ne $cached) {
            Write-Warning "Auth: AD unreachable. Using cached role '$cached' for '$userName' (expires in $($Script:RoleCache.TTLMinutes)min)."
            $qualifiedRole = $cached
        } else {
            Write-Warning "Auth: AD unreachable and no valid role cache for '$userName'. Defaulting to Operator. Administrators may need to re-authenticate when AD recovers."
        }
    } elseif (-not $adUnavailable) {
        # AD was reachable: refresh the cache with the resolved role.
        try { Write-RoleCache -UserName $userName -Role $qualifiedRole } catch { }
    }

    # Apply RequestedRole as a ceiling (user cannot escalate above their qualification)
    if ($Script:RoleHierarchy[$RequestedRole] -lt $Script:RoleHierarchy[$qualifiedRole]) {
        return $RequestedRole
    }
    return $qualifiedRole
}

function Test-Permission {
    <#
    .SYNOPSIS
        Returns $true if $Role is allowed to perform $Action.
        FAIL-SAFE: unknown actions return $false (deny by default).
    .PARAMETER Role
        The caller's resolved role string.
    .PARAMETER Action
        The action string (must exist in $ActionRoleMap to be allowed).
    #>
    param(
        [Parameter(Mandatory)][ValidateSet('Operator','Developer','Admin')][string]$Role,
        [Parameter(Mandatory)][string]$Action
    )

    if (-not $Script:ActionRoleMap.ContainsKey($Action)) {
        return $false   # fail-safe: unknown action = denied
    }

    $required = $Script:ActionRoleMap[$Action]
    return ($Script:RoleHierarchy[$Role] -ge $Script:RoleHierarchy[$required])
}

function Assert-Permission {
    <#
    .SYNOPSIS
        Throws a descriptive error if the current role cannot perform $Action.
        Use at the start of any sensitive operation instead of silent denials.
    #>
    param(
        [Parameter(Mandatory)][string]$Role,
        [Parameter(Mandatory)][string]$Action
    )

    if (-not (Test-Permission -Role $Role -Action $Action)) {
        throw "Access denied: role '$Role' is not authorized to perform '$Action'."
    }
}
