<#
    .SYNOPSIS
        Module with helper functions to verify the authorization of an Azure connection.

    .DESCRIPTION
        This module allows you to verify the authorization of an Azure connection. If the connection
        is not authorized, the module will try to authorize the connection by presenting the user with
        an authorization dialog.
    
    .FUNCTIONALITY
        - Verify-ConnectionAuthorization

    .PARAMETER ResourceGroupName <string> [required]
        The resource group name in which the connection to verify resides.
    
    .PARAMETER ConnectionName <string> [required]
        The name of the connection to verify.

    .OUTPUTS
        None

    .LINK
        Inspired by:
            - https://github.com/logicappsio/LogicAppConnectionAuth/blob/master/LogicAppConnectionAuth.ps1
        
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

Remove-Module -Name Verify-ConnectionAuthorization -WarningAction SilentlyContinue -ErrorAction SilentlyContinue -Verbose:$false -Force
Import-Module -Name (Join-Path (Split-Path -parent $PSCommandPath) "\Modules\Verify-ConnectionAuthorization.psm1") -Scope Local -Force -ErrorAction Stop -WarningAction SilentlyContinue -Verbose:$false | Out-Null

$ResourceGroupName = 'rg-aad2spup'
$ConnectionName = 'office365'

Verify-ConnectionAuthorization -ResourceGroupName $ResourceGroupName -ConnectionName $ConnectionName -Verbose