<#
    .SYNOPSIS
        Module with helper functions to connect to SharePoint Online with PnP.

    .DESCRIPTION
        This module allows you to connect to SharePoint Online with PnP using various authentication methods, like:
            - ManagedIdentity
            - CredentialsWithoutMFA
            - CredentialsWithMFA
            - ClientSecret
            - CertificateThumbprint
            - CertificatePath
    
    .FUNCTIONALITY
        - ConnectTo-SharePointOnline

    .PARAMETER SiteUrl <string> [required]
        The SharePoint Online site URL to connect to.
        Example: https://contoso.sharepoint.com/sites/contoso

    .PARAMETER AuthenticationMethod <string> [optional]
        The authentication method to use to connect to SharePoint Online. 
        Either ManagedIdentity, CredentialsWithoutMFA, CredentialsWithMFA, ClientSecret, CertificateThumbprint or CertificatePath.
    
    .PARAMETER TenantId <string> [optional]
        The tenant ID of the Azure Active Directory tenant that is used to connect to SharePoint Online.
        Only required when using the authentication methods CertificateThumbprint or CertificatePath.
        Example: 00000000-0000-0000-0000-000000000000
    
    .PARAMETER UserName <mailaddress> [optional]
        The username of the user that is used to connect to SharePoint Online.
        Only required when using the authentication methods CredentialsWithoutMFA.
        Example: johndoe@contoso.com
    
    .PARAMETER Pass <string> [optional]
        The password of the user that is used to connect to SharePoint Online.
        Only required when using the authentication methods CredentialsWithoutMFA.
    
    .PARAMETER ClientId <string> [optional]
        The client ID of the Azure AD app registration that is used to connect to SharePoint Online.
        Only required when using the authentication methods ClientSecret, CertificateThumbprint or CertificatePath.
        Example: 00000000-0000-0000-0000-000000000000
    
    .PARAMETER ClientSecret <string> [optional]
        The client secret of the Azure AD app registration that is used to connect to SharePoint Online.
        Only required when using the authentication method ClientSecret.
    
    .PARAMETER CertificateThumbprint <string> [optional]
        The thumbprint of the certificate that is used to connect to SharePoint Online.
        Only required when using the authentication method CertificateThumbprint.
        Example: 0000000000000000000000000000000000000000
    
    .PARAMETER CertificatePath <string> [optional]
        The path to the certificate that is used to connect to SharePoint Online.
        Only required when using the authentication method CertificatePath.
        Example: C:\certificates\contoso.pfx
    
    .PARAMETER CertificatePass <string> [optional]
        The password of the certificate that is used to connect to SharePoint Online.
        Only required when using the authentication method CertificatePath.

    .PARAMETER ReturnConnection <boolean> [optional]
        Whether or not to return the connection as an object.
        Default: false
        Note: ReturnConnection does not work when using the authentication method ManagedIdentity.

    .OUTPUTS
        None or a PnP.PowerShell.Commands.Base.PnPConnection object if ReturnConnection is set to true.

    .LINK
        Inspired by:
            - https://pnp.github.io/powershell/cmdlets/Connect-PnPOnline.html
        
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

Remove-Module -Name ConnectTo-SharePointOnline -WarningAction SilentlyContinue -ErrorAction SilentlyContinue -Verbose:$false -Force
Import-Module -Name (Join-Path (Split-Path -parent $PSCommandPath) "\Modules\ConnectTo-SharePointOnline.psm1") -Scope Local -Force -ErrorAction Stop -WarningAction SilentlyContinue -Verbose:$false | Out-Null

$SiteUrl = 'https://{SHAREPOINT_DOMAIN}.sharepoint.com/sites/{SITE_URL}'
$AuthenticationMethod = 'CertificateThumbprint'
$UserName = $null
$Pass = $null
$TenantId = '{TENANT_ID}'
$ClientId = '{CLIENT_ID}'
$ClientSecret = $null
$CertificateThumbprint = '{CERTIFICATE_THUMBPRINT}}'
$CertificatePath = $null
$CertificatePass = $null
$ReturnConnection = $true

$SharePointConnection = ConnectTo-SharePointOnline -SiteUrl $SiteUrl -AuthenticationMethod $AuthenticationMethod -TenantId $TenantId -UserName $UserName -Pass $Pass -ClientId $ClientId -ClientSecret $ClientSecret -CertificateThumbprint $CertificateThumbprint -CertificatePath $CertificatePath -CertificatePass $CertificatePass -ReturnConnection $ReturnConnection -Verbose