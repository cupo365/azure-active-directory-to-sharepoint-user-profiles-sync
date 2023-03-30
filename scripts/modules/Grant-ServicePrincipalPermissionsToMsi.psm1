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
$ErrorActionPreference = 'Continue'
$ProgressPreference = 'Continue'
$VerbosePreference = "Continue"
$DebugPreference = 'Continue' # 'Continue'

<# ---------------- Global variables ---------------- #>
$GLOBAL:requiredPSVersion = "5.1"
$GLOBAL:psCommandPath = $MyInvocation.PSCommandPath

<# ---------------- Program execution ---------------- #>

Function Grant-ServicePrincipalPermissionsToMsi {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $MsiObjectId,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $ServicePrincipalId,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [array] $PermissionScopes
    )

    Try {
        Initialize

        Write-Verbose -Message "Fetching MSI..."
        $LoopCount = 0
        $Msi = $null

        Do {
            $LoopCount++

            Try {
                $Msi = Get-AzureADServicePrincipal -Filter "objectId eq '$($MsiObjectId)'"
                Write-Debug -Message "MSI: $($Msi)"
                Write-Debug -Message "Loop count: $($LoopCount)"
            }
            Catch [Exception] {
                # There is a delay between creating the MSI and being able to fetch it.
                Write-Verbose -Message "MSI does not exist yet. Waiting 10 seconds to overcome a potential delay..."
                Start-Sleep -Seconds 10
                Write-Verbose -Message "Resuming. Checking again (attempt $($LoopCount + 1) out of maximum 6)..."                
            }
            
        } Until ($null -ne $Msi -or 5 -lt $LoopCount)

        If ($null -eq $Msi) {
            throw "Could not fetch MSI with ID '$($MsiObjectId)'"
        }

        Write-Verbose -Message "Fetched MSI after $($LoopCount) attempt(s)."

        Write-Verbose -Message "Fetching current MSI permission scopes..."
        $CurrentMsiPermissionScopes = Get-AzureADServiceAppRoleAssignedTo -ObjectId $MsiObjectId -All $true -Verbose:$false
        Write-Debug -Message "Current MSI permission scopes: $($CurrentMsiPermissionScopes)"

        Write-Verbose -Message "Fetching Service Principal..."
        $ServicePrincipal = Get-AzureADServicePrincipal -Filter "appId eq '$ServicePrincipalId'" -Verbose:$false

        If ($null -eq $ServicePrincipal) {
            throw "Could not fetch service principal with ID '$($ServicePrincipal)'."
        }

        Write-Debug -Message "Service principal: $($ServicePrincipal)"

        Foreach ($PermissionScope in $PermissionScopes) {
            Write-Verbose -Message "Fetching permission scope '$($PermissionScope)'..."
            $AppRole = $ServicePrincipal.AppRoles | Where-Object { $_.Value -eq $PermissionScope -and $_.AllowedMemberTypes -contains "Application" }

            If ($null -eq $AppRole) {
                throw "Permission scope '$($PermissionScope)' does not exist on service principal '$($ServicePrincipalId)'."
            }

            Write-Debug -Message "Permission scope: $($AppRole)"

            $ExistingMsiPermissionScope = $CurrentMsiPermissionScopes | Where-Object { $_.Id -eq $AppRole.Id -and $_.ResourceId -eq $ServicePrincipal.ObjectId }
            If ($null -ne $ExistingMsiPermissionScope) {
                Write-Verbose -Message "Permission scope '$($PermissionScope)' for service principal '$($ServicePrincipalId)' has already been granted to the MSI. Continuing."
                continue
            }

            Write-Verbose -Message "Granting MSI permission scope '$($PermissionScope)' for service principal '$($ServicePrincipalId)'..."
            $NewAppRole = New-AzureAdServiceAppRoleAssignment -ObjectId $Msi.ObjectId -PrincipalId $Msi.ObjectId -ResourceId $ServicePrincipal.ObjectId -Id $AppRole.Id -ErrorAction SilentlyContinue -ErrorVariable NewAppRoleError -Verbose:$false
            
            If ($false -eq [string]::IsNullOrEmpty($NewAppRoleError)) {
                throw "Could not grant MSI permission scope '$($PermissionScope)' for service principal '$($ServicePrincipalId)'. Reason: $($NewAppRoleError)."
            }

            Write-Verbose -Message "Successfully granted permission scope '$($PermissionScope)' for service principal '$($NewAppRole.ResourceDisplayName)' to MSI '$($NewAppRole.PrincipalDisplayName)'!"
        }
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to grant the MSI the requested service principal permissions. Message: $($_.Exception.Message)."

        Write-Verbose -Message "Terminating program. Reason: could not successfully grant the MSI the requested service principal permissions."
        Finish-Up -TerminateWithError $true
    }
}

Export-ModuleMember -Function Grant-ServicePrincipalPermissionsToMsi

<# ---------------- Helper functions ---------------- #>

Function Initialize {
    Try {
        Remove-Module -Name Validate-PSVersion -WarningAction SilentlyContinue -ErrorAction SilentlyContinue -Verbose:$false -Force
        Import-Module -Name ".\Modules\Validate-PSVersion.psm1" -Scope Local -Force -ErrorAction Stop -WarningAction SilentlyContinue -Verbose:$false | Out-Null

        Validate-PSVersion -RequiredPSVersion $GLOBAL:requiredPSVersion -CommandPath $GLOBAL:psCommandPath -Verbose
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while initializing. Message: $($_.Exception.Message)"

        Finish-Up -TerminateWithError $true
    }
}

Function Finish-Up {
    Param(
        [Parameter(Mandatory = $false)] [boolean] $TerminateWithError = $false
    )

    Try {
        Write-Verbose -Message "Disconnecting the session..."
        Disconnect-AzureAD -Verbose:$false | Out-Null
    }
    Catch [Exception] {
        If ($_.Exception.Message -ne 'No connection to disconnect') {
            Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while finishing up. Message: $($_.Exception.Message)"
        }
    }
    Finally {
        Write-Verbose -Message "Finished."
        
        If ($TerminateWithError) {
            exit 1 # Terminate the program as failed
        }
    }
}