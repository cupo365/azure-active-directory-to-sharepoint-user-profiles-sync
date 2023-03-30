<#
    .SYNOPSIS
        Module with helper function to validate the PowerShell version.

    .DESCRIPTION
        This module allows you to call a function that takes care of validating the PowerShell version.
        If the script is not running the required version, it will be re-launched as the required version if available on the users machine.
    
    .FUNCTIONALITY
        - Validate-PSVersion

    .PARAMETER RequiredPSVersion <string> [required]
        The required PowerShell version.
        Examples: "5.1", "7.0"

    .PARAMETER CommandPath <string> [optional]
        The path to the script that is calling this module.
        Example: "./Call-Add-MSIPermissions.ps1"

    .OUTPUTS
        None

    .LINK
        Developers:
            - Lennart de Waart
                - GitHub: https://github.com/cupo365
                - LinkedIn: https://www.linkedin.com/in/lennart-dewaart/
#>

<# --------------- Runtime variables ---------------- #>
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'Continue'
$VerbosePreference = "Continue"
$DebugPreference = 'SilentlyContinue' # 'Continue'

Remove-Module -Name Validate-PSVersion -WarningAction SilentlyContinue -ErrorAction SilentlyContinue -Verbose:$false -Force
Import-Module -Name (Join-Path (Split-Path -parent $PSCommandPath) "\Modules\Validate-PSVersion.psm1") -Scope Local -Force -ErrorAction Stop -WarningAction SilentlyContinue -Verbose:$false | Out-Null

$RequiredPSVersion = ''

Validate-PSVersion -RequiredPSVersion $RequiredPSVersion -Verbose