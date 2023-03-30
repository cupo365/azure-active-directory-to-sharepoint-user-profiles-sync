<#
    .SYNOPSIS
        Module with helper functions to fetch the status of a SharePoint User Profile bulk import job.

    .DESCRIPTION
        This module allows you to connect to SharePoint Online with PnP using various authentication methods, like:
            - ManagedIdentity
            - CredentialsWithoutMFA
            - CredentialsWithMFA
            - ClientSecret
            - CertificateThumbprint
            - CertificatePath
        Once authenticated, the script will fetch the status of the sync job with the specified ID.
        To fetch the job status, you have to have Sites.FullControl.All (SharePoint admin) permissions.
    
    .FUNCTIONALITY
        - Get-SharePointUserProfilesBulkImportJobStatus

    .PARAMETER Domain <string> [required]
        The SharePoint Online tenant domain to connect to.
        Example: contoso

    .PARAMETER JobId <string> [required]
        The ID of the sync job to fetch the status from.
        Example: 00000000-0000-0000-0000-000000000000

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
        Microsoft.Online.SharePoint.TenantManagement.ImportProfilePropertiesJobInfo object
        See also: https://learn.microsoft.com/en-us/previous-versions/office/sharepoint-csom/mt643008(v%3Doffice.15)

        {
            "JobId": "<string>",
            "State": "<string>",
            "NumberOfUsers": "<int | null>",
            "PropertyMapping": "<object>",
            "SourceUri": "<string | null>",
            "LogFolderUri": "<string | null>",
            "Error": "<string>",
            "ErrorMessage": "<string | null>",
            "Context": "<PnP.Framework.PnPClientContext>",
            "Tag": "<object>",
            "Path": "<Microsoft.SharePoint.Client.ObjectPathMethod>",
            "ObjectVersion": "<string | null>",
            "ServerObjectIsNull": "<bool>",
            "TypedObject": "<Microsoft.Online.SharePoint.TenantManagement.ImportProfilePropertiesJobInfo>"
        }

    .LINK
        Inspired by:
            - https://pnp.github.io/powershell/cmdlets/Get-PnPUPABulkImportStatus.html
        
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

Remove-Module -Name Get-SharePointUserProfilesBulkImportJobStatus -WarningAction SilentlyContinue -ErrorAction SilentlyContinue -Verbose:$false -Force
Import-Module -Name (Join-Path (Split-Path -parent $PSCommandPath) "\Modules\Get-SharePointUserProfilesBulkImportJobStatus.psm1") -Scope Local -Force -ErrorAction Stop -WarningAction SilentlyContinue -Verbose:$false | Out-Null

$Domain = '{SHAREPOINT_DOMAIN}'
$JobId = '{JOB_ID}'
$AuthenticationMethod = 'CertificateThumbprint'
$UserName = $null
$Pass = $null
$TenantId = '{TENANT_ID}'
$ClientId = '{CLIENT_ID}'
$ClientSecret = $null
$CertificateThumbprint = '{CERTIFICATE_THUMBPRINT}'
$CertificatePath = $null
$CertificatePass = $null

Get-SharePointUserProfilesBulkImportJobStatus -Domain $Domain -JobId $JobId -AuthenticationMethod $AuthenticationMethod -TenantId $TenantId -UserName $UserName -Pass $Pass -ClientId $ClientId -ClientSecret $ClientSecret -CertificateThumbprint $CertificateThumbprint -CertificatePath $CertificatePath -CertificatePass $CertificatePass -Verbose