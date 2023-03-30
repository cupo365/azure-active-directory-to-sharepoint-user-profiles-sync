<#
    .SYNOPSIS
        Module with helper functions to connect to Azure AD.

    .DESCRIPTION
        This module allows you to connect to Azure AD using various authentication methods, like:
            - CredentialsWithoutMFA
            - CredentialsWithMFA
            - CertificateThumbprint
    
    .FUNCTIONALITY
        - ConnectTo-AzureAD

    .PARAMETER AuthenticationMethod <string> [optional]
        The authentication method to use to connect to Azure AD. 
        Either CredentialsWithoutMFA, CredentialsWithMFA, CertificateThumbprint.
    
    .PARAMETER TenantId <string> [optional]
        The tenant ID of the Azure Active Directory tenant that is used to connect to Azure AD.
        Only required when using the authentication method CertificateThumbprint.
        Example: 00000000-0000-0000-0000-000000000000
    
    .PARAMETER UserName <mailaddress> [optional]
        The username of the user that is used to connect to Azure AD.
        Only required when using the authentication method CredentialsWithoutMFA.
        Example: johndoe@contoso.com
    
    .PARAMETER Pass <string> [optional]
        The password of the user that is used to connect to Azure AD.
        Only required when using the authentication method CredentialsWithoutMFA.
    
    .PARAMETER ClientId <string> [optional]
        The client ID of the Azure AD app registration that is used to connect to Azure AD.
        Only required when using the authentication method CertificateThumbprint.
        Example: 00000000-0000-0000-0000-000000000000
    
    .PARAMETER CertificateThumbprint <string> [optional]
        The thumbprint of the certificate that is used to connect to Azure AD.
        Only required when using the authentication method CertificateThumbprint.
        Example: 0000000000000000000000000000000000000000

    .OUTPUTS
        None

    .LINK
        Inspired by:
            - https://learn.microsoft.com/en-us/powershell/module/azuread/connect-azuread
        
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

Remove-Module -Name ConnectTo-AzureAD -WarningAction SilentlyContinue -ErrorAction SilentlyContinue -Verbose:$false -Force
Import-Module -Name (Join-Path (Split-Path -parent $PSCommandPath) "\Modules\ConnectTo-AzureAD.psm1") -Scope Local -Force -ErrorAction Stop -WarningAction SilentlyContinue -Verbose:$false | Out-Null

$AuthenticationMethod = 'CredentialsWithMFA'
$UserName = $null
$Pass = $null
$TenantId = '{TENANT_ID}'
$ClientId = '{CLIENT_ID}'
$CertificateThumbprint = '{CERTIFICATE_THUMBPRINT}'

ConnectTo-AzureAD -AuthenticationMethod $AuthenticationMethod -TenantId $TenantId -UserName $UserName -Pass $Pass -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -Verbose | Out-Null