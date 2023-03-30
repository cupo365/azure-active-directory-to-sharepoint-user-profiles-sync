<#
    .SYNOPSIS
        Script to deploy the aad2spup solution resources.

    .DESCRIPTION
        This script allows you to deploy all resources for the aad2spup solution.
        TODO
    
    .FUNCTIONALITY
        TODO
    
    .PARAMETER ResourceGroupName <string> [required]
        The name of the Azure resource group in which the Azure resources should be deployed.
        Default: rg-aad2spup
    
    .PARAMETER ResourceGroupLocation <string> [required]
        The location of the Azure resource group and resources.
        Default: West Europe
    
    .PARAMETER AzureResourceTemplateFilePath <string> [required]
        The (relative) file path to the Azure resource template to deploy.
        Default: ../arm/azuredeploy.bicep
    
    .PARAMETER AzureResourceTemplateParametersFilePath <string> [required]
        The (relative) file path to the Azure Resource Manager template parameters to deploy.
        Default: ../arm/azuredeploy.parameters.production.json

    .PARAMETER DatabaseSiteUrl <string> [required]
        The SharePoint Online site URL in which the required resources for data storage may be implemented.
        Example: https://contoso.sharepoint.com/sites/contoso

    .PARAMETER PnPSiteProvisioningTemplateFilePath <string> [required]
        The path to the PnP site template to apply to the SharePoint database site.
        Default: ../templates/pnp-site-template.xml

    .PARAMETER AuthenticationMethod <string> [optional]
        The authentication method to use to connect to Azure. 
        Either CredentialsWithoutMFA, CredentialsWithMFA (default) or CertificateThumbprint.
    
    .PARAMETER TenantId <string> [optional]
        The tenant ID of the Azure Active Directory tenant that is used to connect to Azure.
        Only required when using the authentication methods CertificateThumbprint.
        Example: 00000000-0000-0000-0000-000000000000
    
    .PARAMETER UserName <mailaddress> [optional]
        The username of the user that is used to connect to Azure.
        Only required when using the authentication methods CredentialsWithoutMFA.
        Example: johndoe@contoso.com
    
    .PARAMETER Pass <string> [optional]
        The password of the user that is used to connect to Azure.
        Only required when using the authentication methods CredentialsWithoutMFA.
    
    .PARAMETER ClientId <string> [optional]
        The client ID of the Azure AD app registration that is used to connect to Azure.
        Only required when using the authentication method CertificateThumbprint.
        Example: 00000000-0000-0000-0000-000000000000
    
    .PARAMETER CertificateThumbprint <string> [optional]
        The thumbprint of the certificate that is used to connect to Azure.
        Only required when using the authentication method CertificateThumbprint.
        Example: 0000000000000000000000000000000000000000
    
    .NOTES
        Version: 1.0.0
        Requires:
            - PnP.PowerShell v1.12.0 - https://www.powershellgallery.com/packages/PnP.PowerShell/1.12.0
            - PowerShell 5.1 - https://learn.microsoft.com/en-us/skypeforbusiness/set-up-your-computer-for-windows-powershell/download-and-install-windows-powershell-5-1
            - SharePoint Online User.Read.All and Sites.FullControl.All permissions 
            - Graph User.Read.All permissions
            TODO

    .LINK
        Developers:
            - Lennart de Waart
                - GitHub: https://github.com/cupo365
                - LinkedIn: https://www.linkedin.com/in/lennart-dewaart/

        Inspired by: 
        - https://aztoso.com/security/microsoft-graph-permissions-managed-identity/
        - https://github.com/logicappsio/LogicAppConnectionAuth/blob/master/LogicAppConnectionAuth.ps1
#>

Param(
    [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [string] $ResourceGroupName = "rg-aad2spup",
    [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [string] $ResourceGroupLocation = "West Europe",
    [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [string] $AzureResourceTemplateFilePath = "../arm/azuredeploy.bicep",
    [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [string] $AzureResourceTemplateParametersFilePath = "../arm/azuredeploy.parameters.production.json",
    [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [ValidatePattern('(http[s]?|[s]?ftp[s]?)(:\/\/)([^\s,]+)')] [string] $DatabaseSiteUrl = 'https://cupo365.sharepoint.com/sites/aad2spup',
    [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [string] $PnPSiteProvisioningTemplateFilePath = '../templates/pnp-site-template.xml',
    [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [ValidateSet("CredentialsWithoutMFA", "CredentialsWithMFA", "CertificateThumbprint")] [string] $AuthenticationMethod = 'CertificateThumbprint',
    [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [mailaddress] $UserName,
    [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $Pass,
    [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $TenantId = '{TENANT_ID}',
    [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $ClientId = '{CLIENT_ID}',
    [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $CertificateThumbprint = '{CERTIFICATE_THUMBPRINT}'
)


<# ------------------ Requirements ------------------ #>
#Requires -Version 5.1
##Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

<# --------------- Runtime variables ---------------- #>
$ErrorActionPreference = 'Continue'
$ProgressPreference = 'Continue'
$VerbosePreference = "Continue"
$DebugPreference = 'SilentlyContinue' # 'Continue'

<# ---------------- Global variables ---------------- #>
$GLOBAL:requiredPSVersion = "5.1"
$GLOBAL:requiredPSModules = @("Az.Accounts", "Az.Resources", "AzureAD", "PnP.PowerShell") # "AzureRm"
$Env:PNPPOWERSHELL_UPDATECHECK = 'Off'

<# ---------------- Helper functions ---------------- #>
Function Initialize {
    Try {
        Remove-Module -Name Validate-PSVersion -WarningAction SilentlyContinue -ErrorAction SilentlyContinue -Verbose:$false -Force
        Import-Module -Name (Join-Path (Split-Path -parent $PSCommandPath) "\Modules\Validate-PSVersion.psm1") -Scope Local -Force -ErrorAction Stop -WarningAction SilentlyContinue -Verbose:$false | Out-Null

        Validate-PSVersion -RequiredPSVersion $GLOBAL:requiredPSVersion -CommandPath $MyInvocation.PSCommandPath -Verbose

        Remove-Module -Name Import-Dependency -WarningAction SilentlyContinue -ErrorAction SilentlyContinue -Verbose:$false -Force
        Import-Module -Name (Join-Path (Split-Path -parent $PSCommandPath) "\Modules\Import-Dependency.psm1") -Scope Local -Force -ErrorAction Stop -WarningAction SilentlyContinue -Verbose:$false | Out-Null

        Write-Verbose -Message "Initializing..."

        Foreach ($Module in $GLOBAL:requiredPSModules) {
            Import-Dependency -ModuleName $Module
        }
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while initializing. Message: $($_.Exception.Message)"

        Write-Verbose -Message "Terminating program. Reason: could not initialize."
        Finish-Up -TerminateWithError $true
    }
}

Function Apply-PnPTemplate {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [ValidatePattern('(http[s]?|[s]?ftp[s]?)(:\/\/)([^\s,]+)')] [string] $SiteUrl,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $TemplateFilePath,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [ValidateSet("ManagedIdentity", "CredentialsWithoutMFA", "CredentialsWithMFA", "ClientSecret", "CertificateThumbprint", "CertificatePath")] [string] $AuthenticationMethod,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $TenantId,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [mailaddress] $UserName,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $Pass,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $ClientId,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $ClientSecret,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $CertificateThumbprint,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $CertificatePath,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $CertificatePass
    )

    Try {
        Remove-Module -Name Apply-PnPTemplate -WarningAction SilentlyContinue -ErrorAction SilentlyContinue -Verbose:$false -Force
        Import-Module -Name (Join-Path (Split-Path -parent $PSCommandPath) "\Modules\Apply-PnPTemplate.psm1") -Scope Local -Force -ErrorAction Stop -WarningAction SilentlyContinue -Verbose:$false | Out-Null

        Apply-PnPTemplate\Apply-PnPTemplate -SiteUrl $SiteUrl -TemplateFilePath $TemplateFilePath -AuthenticationMethod $AuthenticationMethod -TenantId $TenantId -UserName $UserName -Pass $Pass -ClientId $ClientId -ClientSecret $ClientSecret -CertificateThumbprint $CertificateThumbprint -CertificatePath $CertificatePath -CertificatePass $CertificatePass -Verbose | Out-Null
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while applying the PnP template from path '$($TemplateFilePath)' to site '$($SiteUrl)'. Message: $($_.Exception.Message)"

        Finish-Up -SharePointConnection $SharePointConnection -TerminateWithError $true
    }
}

Function ConnectTo-AzAccount {
    Param(
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [ValidateSet("CredentialsWithoutMFA", "CredentialsWithMFA", "CertificateThumbprint")] [string] $AuthenticationMethod = 'CredentialsWithMFA',
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [mailaddress] $UserName,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $Pass,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $TenantId,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $ClientId,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $CertificateThumbprint
    )

    Try {
        Remove-Module -Name ConnectTo-AzAccount -WarningAction SilentlyContinue -ErrorAction SilentlyContinue -Verbose:$false -Force
        Import-Module -Name (Join-Path (Split-Path -parent $PSCommandPath) "\Modules\ConnectTo-AzAccount.psm1") -Scope Local -Force -ErrorAction Stop -WarningAction SilentlyContinue -Verbose:$false | Out-Null

        ConnectTo-AzAccount\ConnectTo-AzAccount -AuthenticationMethod $AuthenticationMethod -TenantId $TenantId -UserName $UserName -Pass $Pass -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -ErrorVariable ConnectToAzAccountError -Verbose | Out-Null

        If ($false -eq [string]::IsNullOrEmpty($ConnectToAzAccountError) -and $ConnectToAzAccountError -notmatch 'Get-AutomationConnection' -and $ConnectToAzAccountError -notmatch 'Get-AutomationPSCredential') {
            throw "Could not connect to Azure. Reason: $($ConnectToAzAccountError)."
        }
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while connecting to Azure. Message: $($_.Exception.Message)"

        Write-Verbose -Message "Terminating program. Reason: could not connect to Azure."
        Finish-Up -TerminateWithError $true
    }
}

Function Verify-ResourceGroup {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $ResourceGroupName,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $ResourceGroupLocation
    )

    Try {
        Write-Verbose -Message "Verifying the existence of resource group '$($ResourceGroupName)'..."
        Get-AzResourceGroup -Name $ResourceGroupName -ErrorVariable notPresent -ErrorAction SilentlyContinue -Verbose:$false | Out-Null

        if ($notPresent) {
            Write-Verbose -Message "Resource group does not exist. Creating..."
            $NewResourceGroup = New-AzResourceGroup -Name $ResourceGroupName -Location $ResourceGroupLocation -Verbose:$false

            If ($NewResourceGroup.ProvisioningState -ne "Succeeded") {
                throw "Could not create resource group with name $($ResourceGroupName) and location $($ResourceGroupLocation). Provisioning state: $($NewResourceGroup.ProvisioningState)."
            }
        }
        else {
            Write-Verbose -Message "Resource group already exists."
        }
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while verifying the existence of the resource group '$($ResourceGroupName)' with location '$($ResourceGroupLocation)'. Message: $($_.Exception.Message)"

        Write-Verbose -Message "Terminating program. Reason: could not verify the existence of the resource group."
        Finish-Up -TerminateWithError $true
    }
}

Function Deploy-AzureResourceTemplate {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $ResourceGroupName,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $AzureResourceTemplateFilePath,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $AzureResourceTemplateParametersFilePath
    )

    Try {
        Write-Verbose -Message "Deploying Azure Resource Template. This could take a while..."
        If ($AzureResourceTemplateFilePath | Test-Path) {
            If ($AzureResourceTemplateParametersFilePath | Test-Path) {
                $Deployment = New-AzResourceGroupDeployment -Name "PSDeploy-$(Get-Date -Format "HHmmddMMyyyy")" -ResourceGroupName $ResourceGroupName -TemplateFile $AzureResourceTemplateFilePath -TemplateParameterFile $AzureResourceTemplateParametersFilePath -WarningAction SilentlyContinue -Verbose:$false
            }
            Else {
                throw "Could not find the Azure Resource Template parameters file at path '$($AzureResourceTemplateParametersFilePath)'."
            }
        }
        Else {
            throw "Could not find the Azure Resource Template file at path '$($AzureResourceTemplateFilePath)'."
        }

        Write-Verbose -Message "Verifying the deployment..."
        If ($Deployment.ProvisioningState -eq "Succeeded" -and $null -ne $Deployment.Outputs) {
            Write-Verbose -Message "Deployed!"
        }
        Else {
            throw "Failed to deploy the Azure Resource Template. Provisioning state: $($Deployment.ProvisioningState)."
        }

        return $Deployment.Outputs
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while deploying the Azure Resource Template. Message: $($_.Exception.Message)"

        Write-Verbose -Message "Terminating program. Reason: could not successfully deploy the Azure Resource Template."
        Finish-Up -TerminateWithError $true
    }
}

Function ConnectTo-AzureAD {
    Param(
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [ValidateSet("CredentialsWithoutMFA", "CredentialsWithMFA", "CertificateThumbprint")] [string] $AuthenticationMethod = 'CredentialsWithMFA',
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [mailaddress] $UserName,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $Pass,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $TenantId,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $ClientId,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $CertificateThumbprint
    )

    Try {
        Remove-Module -Name ConnectTo-AzureAD -WarningAction SilentlyContinue -ErrorAction SilentlyContinue -Verbose:$false -Force
        Import-Module -Name (Join-Path (Split-Path -parent $PSCommandPath) "\Modules\ConnectTo-AzureAD.psm1") -Scope Local -Force -ErrorAction Stop -WarningAction SilentlyContinue -Verbose:$false | Out-Null

        ConnectTo-AzureAD\ConnectTo-AzureAD -AuthenticationMethod $AuthenticationMethod -TenantId $TenantId -UserName $UserName -Pass $Pass -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -ErrorVariable ConnectToAzureAdError -Verbose | Out-Null

        If ($false -eq [string]::IsNullOrEmpty($ConnectToAzureAdError) -and $ConnectToAzureAdError -notmatch 'Get-AutomationConnection' -and $ConnectToAzureAdError -notmatch 'Get-AutomationPSCredential') {
            throw "Could not connect to Azure AD. Reason: $($ConnectToAzureAdError)."
        }
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while connecting to Azure AD. Message: $($_.Exception.Message)"

        Write-Verbose -Message "Terminating program. Reason: could not connect to Azure AD."
        Finish-Up -TerminateWithError $true
    }
}

Function Grant-ServicePrincipalPermissionsToMsi {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $MsiObjectId,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $ServicePrincipalId,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [array] $PermissionScopes
    )

    Try {
        Remove-Module -Name Grant-ServicePrincipalPermissionsToMsi -WarningAction SilentlyContinue -ErrorAction SilentlyContinue -Verbose:$false -Force
        Import-Module -Name (Join-Path (Split-Path -parent $PSCommandPath) "\Modules\Grant-ServicePrincipalPermissionsToMsi.psm1") -Scope Local -Force -ErrorAction Stop -WarningAction SilentlyContinue -Verbose:$false | Out-Null

        Grant-ServicePrincipalPermissionsToMsi\Grant-ServicePrincipalPermissionsToMsi -MsiObjectId $MsiObjectId -ServicePrincipalId $ServicePrincipalId -PermissionScopes $PermissionScopes -Verbose | Out-Null
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while granting the permission scopes '$($PermissionScopes)' for service principal '$($ServicePrincipalId)' to MSI '$($MsiObjectId)'. Message: $($_.Exception.Message)"

        Write-Verbose -Message "Terminating program. Reason: could not grant the service principal permission scopes to the MSI."
        Finish-Up -TerminateWithError $true
    }
}

Function ConnectTo-AzureRmAccount {
    Param(
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [ValidateSet("CredentialsWithoutMFA", "CredentialsWithMFA", "CertificateThumbprint")] [string] $AuthenticationMethod = 'CredentialsWithMFA',
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [mailaddress] $UserName,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $Pass,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $TenantId,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $ClientId,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $CertificateThumbprint
    )

    Try {
        Remove-Module -Name ConnectTo-AzureRmAccount -WarningAction SilentlyContinue -ErrorAction SilentlyContinue -Verbose:$false -Force
        Import-Module -Name (Join-Path (Split-Path -parent $PSCommandPath) "\Modules\ConnectTo-AzureRmAccount.psm1") -Scope Local -Force -ErrorAction Stop -WarningAction SilentlyContinue -Verbose:$false | Out-Null

        ConnectTo-AzureRmAccount\ConnectTo-AzureRmAccount -AuthenticationMethod $AuthenticationMethod -TenantId $TenantId -UserName $UserName -Pass $Pass -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -ErrorVariable ConnectToAzureRmAccountError -Verbose | Out-Null

        If ($false -eq [string]::IsNullOrEmpty($ConnectToAzureRmAccountError) -and $ConnectToAzureRmAccountError -notmatch 'Get-AutomationConnection' -and $ConnectToAzureRmAccountError -notmatch 'Get-AutomationPSCredential') {
            throw "Could not connect to Azure Remote. Reason: $($ConnectToAzureRmAccountError)."
        }
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while connecting to Azure Remote. Message: $($_.Exception.Message)"

        Write-Verbose -Message "Terminating program. Reason: could not connect to Azure Remote."
        Finish-Up -TerminateWithError $true
    }
}

Function Verify-ConnectionAuthorization {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $ResourceGroupName,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $ConnectionName
    )

    Try {       
        Remove-Module -Name Verify-ConnectionAuthorization -WarningAction SilentlyContinue -ErrorAction SilentlyContinue -Verbose:$false -Force
        Import-Module -Name (Join-Path (Split-Path -parent $PSCommandPath) "\Modules\Verify-ConnectionAuthorization.psm1") -Scope Local -Force -ErrorAction Stop -WarningAction SilentlyContinue -Verbose:$false | Out-Null

        Verify-ConnectionAuthorization\Verify-ConnectionAuthorization -ResourceGroupName $ResourceGroupName -ConnectionName $ConnectionName -Verbose | Out-Null
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while verifying the connection '$($ConnectionName)'. Message: $($_.Exception.Message)"

        Write-Verbose -Message "Terminating program. Reason: could not verify the connection."
        Finish-Up -TerminateWithError $true
    }
}

Function Finish-Up {
    Param(
        [Parameter(Mandatory = $false)] [AllowNull()] [PnP.PowerShell.Commands.Base.PnPConnection] $SharePointConnection,
        [Parameter(Mandatory = $false)] [boolean] $TerminateWithError = $false
    )

    Try {
        Write-Verbose -Message "Disconnecting the session..."
        If ($null -ne $SharePointConnection) {
            # Dispose the session according to https://pnp.github.io/powershell/cmdlets/Disconnect-PnPOnline.html#description
            $SharePointConnection = $null
        }

        Disconnect-AzAccount -Verbose:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
        Disconnect-AzureAD -Verbose:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
        Disconnect-AzureRmAccount -Verbose:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
    }
    Catch [Exception] {
        If ($_.Exception.Message -ne 'No connection to disconnect' -and $_.Exception.Message -ne 'Object reference not set to an instance of an object') {
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

<# ---------------- Program execution ---------------- #>

Try {
    Initialize

    Apply-PnPTemplate -SiteUrl $DatabaseSiteUrl -TemplateFilePath $PnPSiteProvisioningTemplateFilePath -AuthenticationMethod $AuthenticationMethod -TenantId $TenantId -UserName $UserName -Pass $Pass -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint

    ConnectTo-AzAccount -AuthenticationMethod $AuthenticationMethod -TenantId $TenantId -UserName $UserName -Pass $Pass -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint

    Verify-ResourceGroup -ResourceGroupName $ResourceGroupName -ResourceGroupLocation $ResourceGroupLocation

    $DeploymentOutput = Deploy-AzureResourceTemplate -ResourceGroupName $ResourceGroupName -AzureResourceTemplateFilePath $AzureResourceTemplateFilePath -AzureResourceTemplateParametersFilePath $AzureResourceTemplateParametersFilePath

    If ($null -ne $DeploymentOutput.logicAppIdentityObjectId.Value -or $null -ne $DeploymentOutput.automationAccountIdentityObjectId.Value) {
        # Application Administrator is required to grant permissions to the MSI. Therefore, an application registration cannot be used to grant permissions to an MSI.
        ConnectTo-AzureAD -AuthenticationMethod 'CredentialsWithMFA' -TenantId $TenantId -UserName $UserName -Pass $Pass -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
    }

    If ($null -ne $DeploymentOutput.automationAccountIdentityObjectId.Value) {
        # Grant SharePoint Sites.FullControl.All and User.Read.All permissions
        Grant-ServicePrincipalPermissionsToMsi -MsiObjectId $DeploymentOutput.automationAccountIdentityObjectId.Value -ServicePrincipalId "00000003-0000-0ff1-ce00-000000000000" -PermissionScopes @("Sites.FullControl.All", "User.Read.All")

        # Grant Graph User.Read.All permissions
        Grant-ServicePrincipalPermissionsToMsi -MsiObjectId $DeploymentOutput.automationAccountIdentityObjectId.Value -ServicePrincipalId "00000003-0000-0000-c000-000000000000" -PermissionScopes @("User.Read.All")
    }

    ConnectTo-AzureRmAccount -AuthenticationMethod $AuthenticationMethod -TenantId $TenantId -UserName $UserName -Pass $Pass -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint

    $DeployedConnections = @($DeploymentOutput.office365ConnectionName.Value) #$DeploymentOutput.sharePointConnectionName.Value, $DeploymentOutput.automationConnectionName.Value
    If (0 -eq $DeployedConnections.Length) {
        Write-Verbose -Message "No connections to verify. Continuing."
    }
    Else {
        Foreach ($DeployedConnection in $DeployedConnections) {
            Verify-ConnectionAuthorization -ResourceGroupName $ResourceGroupName -ConnectionName $DeployedConnection
        }
    }

    Finish-Up -SharePointConnection $SharePointConnection
}
Catch [Exception] {
    Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while executing the program. Message: $($_.Exception.Message)"

    Write-Verbose -Message "Terminating program. Reason: encountered an error before program could successfully finish."

    Finish-Up -TerminateWithError $true
}
