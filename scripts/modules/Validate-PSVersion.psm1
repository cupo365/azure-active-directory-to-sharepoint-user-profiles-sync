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

<# ---------------- Program execution ---------------- #>
Function Validate-PSVersion {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $RequiredPSVersion,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $CommandPath
    )

    Try {
        # There is are some import issue for some modules like AzureAD while using PowerShell 7.
        # More info here: https://github.com/PowerShell/PowerShell/issues/10473
        If ($false -eq $PSVersionTable.PSVersion.ToString().StartsWith($RequiredPSVersion)) {
            If ($false -eq [string]::IsNullOrEmpty($CommandPath)) {
                # Re-launch as configured PS version if the script is not running the required version
                powershell -Version $RequiredPSVersion -File $CommandPath
                exit
            }
            Else {
                # Re-launch as configured PS version if the script is not running the required version
                powershell -Version $RequiredPSVersion -File $MyInvocation.PSCommandPath
                exit
            }
            
        }
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to validate the PowerShell version. Message: $($_.Exception.Message)."

        Write-Verbose -Message "Terminating program. Reason: could not validate the PowerShell version."
        exit 1
    }
}


Export-ModuleMember -Function Validate-PSVersion