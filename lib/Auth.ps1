#Requires -Version 5.1
# =============================================================================
# lib/Auth.ps1
# AutoBuild v2.0 - Role-Based Access Control with AD group validation.
#
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
            # WindowsPrincipal.IsInRole accepts either SAM name or DN on domain-joined machines
            return $principal.IsInRole($GroupDn)
        } catch {
            return $false
        }
    }

    # Helper: check whitelist
    function Test-InWhitelist {
        param([string]$Whitelist)
        if ([string]::IsNullOrWhiteSpace($Whitelist)) { return $false }
        $list = $Whitelist -split '[,;\s]+' | ForEach-Object { $_.Trim().ToLower() } | Where-Object { $_ }
        return $list -contains $userName
    }

    # Determine the highest qualified role
    $qualifiedRole = 'Operator'   # Everyone is at minimum an Operator

    if ((Test-InAdGroup $sec.developerAdGroup) -or (Test-InWhitelist $sec.developerUsers)) {
        $qualifiedRole = 'Developer'
    }

    if ((Test-InAdGroup $sec.adminAdGroup) -or (Test-InWhitelist $sec.adminUsers)) {
        $qualifiedRole = 'Admin'
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
