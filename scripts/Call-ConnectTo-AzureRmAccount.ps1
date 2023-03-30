<#
    .SYNOPSIS
        Module with helper functions to connect to Azure Remote.

    .DESCRIPTION
        This module allows you to connect to Azure Remote using various authentication methods, like:
            - ManagedIdentity
            - CredentialsWithoutMFA
            - CredentialsWithMFA
            - ClientSecret
            - CertificateThumbprint
    
    .FUNCTIONALITY
        - ConnectTo-AzureRmAccount

    .PARAMETER AuthenticationMethod <string> [optional]
        The authentication method to use to connect to Azure Remote. 
        Either ManagedIdentity, CredentialsWithoutMFA, CredentialsWithMFA, ClientSecret, CertificateThumbprint.
    
    .PARAMETER TenantId <string> [optional]
        The tenant ID of the Azure Active Directory tenant that is used to connect to Azure Remote.
        Only required when using the authentication methods ClientSecret, CertificateThumbprint.
        Example: 00000000-0000-0000-0000-000000000000
    
    .PARAMETER UserName <mailaddress> [optional]
        The username of the user that is used to connect to Azure Remote.
        Only required when using the authentication methods CredentialsWithoutMFA.
        Example: johndoe@contoso.com
    
    .PARAMETER Pass <string> [optional]
        The password of the user that is used to connect to Azure Remote.
        Only required when using the authentication methods CredentialsWithoutMFA.
    
    .PARAMETER ClientId <string> [optional]
        The client ID of the Azure AD app registration that is used to connect to Azure Remote.
        Only required when using the authentication methods ClientSecret, CertificateThumbprint.
        Example: 00000000-0000-0000-0000-000000000000
    
    .PARAMETER ClientSecret <string> [optional]
        The client secret of the Azure AD app registration that is used to connect to Azure Remote.
        Only required when using the authentication method ClientSecret.
    
    .PARAMETER CertificateThumbprint <string> [optional]
        The thumbprint of the certificate that is used to connect to Azure Remote.
        Only required when using the authentication method CertificateThumbprint.
        Example: 0000000000000000000000000000000000000000

    .OUTPUTS
        None

    .LINK
        Inspired by:
            - https://learn.microsoft.com/en-us/powershell/module/azurerm.profile/connect-azurermaccount?view=azurermps-6.13.0
        
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

Remove-Module -Name ConnectTo-AzureRmAccount -WarningAction SilentlyContinue -ErrorAction SilentlyContinue -Verbose:$false -Force
Import-Module -Name (Join-Path (Split-Path -parent $PSCommandPath) "\Modules\ConnectTo-AzureRmAccount.psm1") -Scope Local -Force -ErrorAction Stop -WarningAction SilentlyContinue -Verbose:$false | Out-Null

$AuthenticationMethod = 'CertificateThumbprint'
$UserName = $null
$Pass = $null
$TenantId = '{TENANT_ID}'
$ClientId = '{CLIENT_ID}'
$ClientSecret = $null
$CertificateThumbprint = '{CERTIFICATE_THUMBPRINT}'

ConnectTo-AzureRmAccount -AuthenticationMethod $AuthenticationMethod -TenantId $TenantId -UserName $UserName -Pass $Pass -ClientId $ClientId -ClientSecret $ClientSecret -CertificateThumbprint $CertificateThumbprint -Verbose | Out-Null