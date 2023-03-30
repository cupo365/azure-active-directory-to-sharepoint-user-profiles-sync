<#
    .SYNOPSIS
        Script to reindex SharePoint users.

    .DESCRIPTION
        This script allows you to reindex SharePoint users so that the SharePoint Search crawl will pick up the changes in the User Profiles faster.
        The script uses a self-proclaimed 'clever' exploit to trigger the SharePoint Search crawl by changing the value of a user property,
        and placing back the original value, therefore updating the user profile property "modified", based on which the crawl decides whether or not
        to scan the user profile for new property values.

    .PARAMETER tenantName <string> [required]
        The name of the tenant.    

    .LINK
        Source of insipration: https://github.com/wobba/SPO-Trigger-Reindex
#>

<# --------------- Runtime variables ---------------- #>
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'Continue'
$VerbosePreference = "Continue"
$DebugPreference = 'SilentlyContinue' # 'Continue'

Remove-Module -Name Reindex-Users -WarningAction SilentlyContinue -ErrorAction SilentlyContinue -Verbose:$false -Force
Import-Module -Name (Join-Path (Split-Path -parent $PSCommandPath) "\Modules\Reindex-Users.psm1") -Scope Local -Force -ErrorAction Stop -WarningAction SilentlyContinue -Verbose:$false | Out-Null

$tenantName = ''

Reindex-Users -tenantName $tenantName