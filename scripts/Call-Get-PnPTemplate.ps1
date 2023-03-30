<#
    .SYNOPSIS
        Module with helper functions to fetch a PnP template from a SharePoint site.

    .DESCRIPTION
        This module allows you to connect to SharePoint Online with PnP using various authentication methods, like:
            - ManagedIdentity
            - CredentialsWithoutMFA
            - CredentialsWithMFA
            - ClientSecret
            - CertificateThumbprint
            - CertificatePath
        Once authenticated, the script will fetch a PnP site template from the connected SharePoint site.
        Lastly, the script will attempt to open the saved template in VS Code if installed on the executing machine.
    
    .FUNCTIONALITY
        - Get-PnPTemplate

    .PARAMETER SiteUrl <string> [required]
        The SharePoint Online site URL to connect and fetch the site template from.
        Example: https://contoso.sharepoint.com/sites/contoso

    .PARAMETER TemplateType <string> [required]
        The file type of the PnP site template to fetch from the connected SharePoint site.
        Either pnp or xml (default)
        
    .PARAMETER Workspace <string> [optional]
        The path to the workspace folder where the PnP site template will be saved.
        Example: C:\templates

    .PARAMETER Handlers <array> [optional]
        Allows you to only process a specific type of artifact in the site. 
        Notice that this might result in a non-working template, as some of the handlers require 
        other artifacts in place if they are not part of what your extracting.

        Allowed values: None, AuditSettings, ComposedLook, CustomActions, ExtensibilityProviders, 
        Features, Fields, Files, Lists, Pages, Publishing, RegionalSettings, SearchSettings, 
        SitePolicy, SupportedUILanguages, TermGroups, Workflows, SiteSecurity, ContentTypes, 
        PropertyBagEntries, PageContents, WebSettings, Navigation, ImageRenditions, ApplicationLifecycleManagement, 
        Tenant, WebApiPermissions, SiteHeader, SiteFooter, Theme, SiteSettings, All

    .PARAMETER ListsToExtract <string> [optional]
        Specify the lists to extract, either providing their ID or their Title.
        Example: 00000000-0000-0000-0000-000000000000, List1, List2

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

    .OUTPUTS
        PnP site template file (.pnp or .xml)

    .LINK
        Inspired by:
            - https://pnp.github.io/powershell/cmdlets/Get-PnPProvisioningTemplate.html
        
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

Remove-Module -Name Get-PnPTemplate -WarningAction SilentlyContinue -ErrorAction SilentlyContinue -Verbose:$false -Force
Import-Module -Name (Join-Path (Split-Path -parent $PSCommandPath) "\Modules\Get-PnPTemplate.psm1") -Scope Local -Force -ErrorAction Stop -WarningAction SilentlyContinue -Verbose:$false | Out-Null

$SiteUrl = 'https://{SHAREPOINT_DOMAIN}.sharepoint.com/sites/{SITE_URL}'
$TemplateType = 'xml'
$Workspace = '../templates'
$Handlers = $null
$ListsToExtract = @('User Profile Sync Config', 'User Profile Sync Jobs')
$AuthenticationMethod = 'CertificateThumbprint'
$UserName = $null
$Pass = $null
$TenantId = '{TENANT_ID}'
$ClientId = '{CLIENT_ID}'
$ClientSecret = $null
$CertificateThumbprint = '{CERTIFICATE_THUMBPRINT}}'
$CertificatePath = $null

Get-PnPTemplate -SiteUrl $SiteUrl -TemplateType $TemplateType -Workspace $Workspace -Handlers $Handlers -ListsToExtract $ListsToExtract -AuthenticationMethod $AuthenticationMethod -TenantId $TenantId -UserName $UserName -Pass $Pass -ClientId $ClientId -ClientSecret $ClientSecret -CertificateThumbprint $CertificateThumbprint -CertificatePath $CertificatePath -CertificatePass $CertificatePass -Verbose