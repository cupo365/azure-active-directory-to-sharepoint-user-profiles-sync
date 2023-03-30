<#
    .SYNOPSIS
        Module with helper functions to import dependend PowerShell modules.

    .DESCRIPTION
        This module allows you to call one function that takes care of importing a dependend PowerShell module.
        If the module is not installed, it will be installed first.
    
    .FUNCTIONALITY
        - Import-Dependency
        - Install-Dependency

    .PARAMETER ModuleName <string> [required]
        The name of the dependend PowerShell module to import or install.

    .PARAMETER AttemptedBefore <boolean> [optional]
        Whether or not the import of the dependend module has already been attempted before.

    .OUTPUTS
        None

    .LINK
        Developers:
        - Lennart de Waart (https://linkedin.com/in/lennart-dewaart)
#>

<# --------------- Runtime variables ---------------- #>
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'Continue'
$VerbosePreference = "Continue"
$DebugPreference = 'SilentlyContinue' # 'Continue'

Remove-Module -Name Import-Dependency -WarningAction SilentlyContinue -ErrorAction SilentlyContinue -Verbose:$false -Force
Import-Module -Name (Join-Path (Split-Path -parent $PSCommandPath) "\Modules\Import-Dependency.psm1") -Scope Local -Force -ErrorAction Stop -WarningAction SilentlyContinue -Verbose:$false | Out-Null

$ModuleName = ''

Import-Dependency -ModuleName $ModuleName -Verbose