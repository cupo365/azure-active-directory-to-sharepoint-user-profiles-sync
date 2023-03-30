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
            - Lennart de Waart
                - GitHub: https://github.com/cupo365
                - LinkedIn: https://www.linkedin.com/in/lennart-dewaart/
#>

<# --------------- Runtime variables ---------------- #>
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'Continue'
$VerbosePreference = "Continue"
$DebugPreference = 'SilentlyContinue' # 'Continue'

<# ---------------- Global variables ---------------- #>
$GLOBAL:psModulePath = "$($HOME)\Documents\WindowsPowerShell\Modules"

<# ---------------- Program execution ---------------- #>
Function Import-Dependency {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $ModuleName,
        [Parameter(Mandatory = $false)] [boolean] $AttemptedBefore
    )

    Try {
        # Ensure env module path is set correctly (bug in VS Code)
        If ($env:PSModulePath -NotContains $GLOBAL:psModulePath) {
            Write-Verbose -Message "Configuring PS module path..." 
            $env:PSModulePath = $env:PSModulePath + ";$($GLOBAL:psModulePath)"
        }

        Write-Verbose -Message "Importing module '$($ModuleName)'..."
        $VerbosePreference = "SilentlyContinue" # Ignore verbose import messages that significantly slow down the Runbook
        Import-Module -Name $ModuleName -Scope Local -Force -ErrorAction Stop -Verbose:$false | Out-Null
        $VerbosePreference = "Continue" # Restore verbose preference
    }
    Catch [Exception] {
        $VerbosePreference = "Continue" # Restore verbose preference
        Write-Warning -Message "$($ModuleName) was not found."

        If ($true -eq $AttemptedBefore) {
            Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to import the dependend module '$($ModuleName)'. Message: $($_.Exception.Message)."

            Write-Verbose -Message "Terminating program. Reason: could not import dependend modules."
            exit 1
        }
        Else {
            Install-Dependency -ModuleName $ModuleName
        }
    }
}

Export-ModuleMember -Function Import-Dependency

<# ---------------- Helper functions ---------------- #>

Function Install-Dependency {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $ModuleName
    )

    Try {
        # Install module
        Write-Verbose -Message "Installing $($ModuleName) module..."
        $VerbosePreference = "SilentlyContinue" # Ignore verbose import messages that significantly slow down the Runbook
        Install-Module -Name $ModuleName -Scope CurrentUser -AllowClobber -Force -WarningAction SilentlyContinue -ErrorAction Stop -Verbose:$false | Out-Null
        $VerbosePreference = "Continue" # Restore verbose preference

        #Import-PSModule -ModuleName $ModuleName -AttemptedBefore $true
    }
    Catch [Exception] {
        $VerbosePreference = "Continue" # Restore verbose preference
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to install the dependend module '$($ModuleName)'. Message: $($_.Exception.Message)."

        Write-Verbose -Message "Terminating program. Reason: could not install dependend modules."
        exit 1
    }
}