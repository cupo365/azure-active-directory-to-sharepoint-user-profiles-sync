<#
    .SYNOPSIS
        Script to grant service principal permissions to a managed identity.

    .DESCRIPTION
        This script allows you to grant one or more service principal permissions to a managed identity.

    .FUNCTIONALITY
        - Grant-ServicePrincipalPermissionsToMsi

    .PARAMETER MsiObjectId <string> [required]
        The object id of the managed identity you want to grant permissions for.
        Caution: object id and app ip are different fields/values!

    .PARAMETER ServicePrincipalId <string> [required]
        The app id of the service principal from which you want to grant permissions to the managed identity.
        Caution: object id and app ip are different fields/values!
        Some examples:
        - The Graph API has the following id: 00000003-0000-0000-c000-000000000000
        - The SharePoint service principal has the following id: 00000003-0000-0ff1-ce00-000000000000

    .PARAMETER PermissionScopes <array> [required]
        Specify one or more permissions to grant to the managed identity.
        For example: "Directory.Read.All", "Group.Read.All", "GroupMember.Read.All", "User.Read.All"
        or "Sites.FullControl.All"

    .LINK
        Inspired by: https://aztoso.com/security/microsoft-graph-permissions-managed-identity/

        Developers:
            - Lennart de Waart
                - GitHub: https://github.com/cupo365
                - LinkedIn: https://www.linkedin.com/in/lennart-dewaart/
#>

<# ------------------ Requirements ------------------ #>
#Requires -Version 5.1
##Requires -Modules @{ModuleName='AzureAD';ModuleVersion='2.0.2.140'}

<# --------------- Runtime variables ---------------- #>
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'Continue'
$VerbosePreference = "Continue"
$DebugPreference = 'SilentlyContinue' # 'Continue'

Remove-Module -Name Grant-ServicePrincipalPermissionsToMsi -WarningAction SilentlyContinue -ErrorAction SilentlyContinue -Verbose:$false -Force
Import-Module -Name (Join-Path (Split-Path -parent $PSCommandPath) "\Modules\Grant-ServicePrincipalPermissionsToMsi.psm1") -Scope Local -Force -ErrorAction Stop -WarningAction SilentlyContinue -Verbose:$false | Out-Null

$MsiObjectId = "{MSI_OBJECT_ID}"
$ServicePrincipalId = "00000003-0000-0ff1-ce00-000000000000" #"00000003-0000-0000-c000-000000000000"
$PermissionScopes = "Sites.FullControl.All", "User.Read.All" #"User.Read.All"

Grant-ServicePrincipalPermissionsToMsi -MsiObjectId $MsiObjectId -ServicePrincipalId $ServicePrincipalId -PermissionScopes $PermissionScopes -Verbose