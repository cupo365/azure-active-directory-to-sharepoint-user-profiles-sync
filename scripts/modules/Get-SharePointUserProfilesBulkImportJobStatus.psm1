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
$ErrorActionPreference = 'Continue'
$ProgressPreference = 'Continue'
$VerbosePreference = "Continue"
$DebugPreference = 'SilentlyContinue' # 'Continue'

<# ---------------- Global variables ---------------- #>
$GLOBAL:requiredPSModules = "PnP.PowerShell"

<# ---------------- Program execution ---------------- #>

Function Get-SharePointUserProfilesBulkImportJobStatus {
    Param(
        [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [string] $Domain,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [ValidatePattern('^[a-f\d]{4}(?:[a-f\d]{4}-){4}[a-f\d]{12}$')] [string] $JobId,
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

    $SharePointConnection = $null
    If ($true -eq [string]::IsNullOrEmpty($AuthenticationMethod)) {
        $AuthenticationMethod = "CredentialsWithMFA"
    }
    
    Try {
        Initialize

        $SharePointConnection = Connect -SiteUrl "https://$($Domain)-admin.sharepoint.com" -AuthenticationMethod $AuthenticationMethod -TenantId $TenantId -UserName $UserName -Pass $Pass -ClientId $ClientId -ClientSecret $ClientSecret -CertificateThumbprint $CertificateThumbprint -CertificatePath $CertificatePath -CertificatePass $CertificatePass

        Get-JobStatus -JobId $JobId -SharePointConnection $SharePointConnection
        
        Finish-Up -SharePointConnection $SharePointConnection
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while executing the program. Message: $($_.Exception.Message)"

        Finish-Up -SharePointConnection $SharePointConnection -TerminateWithError $true
    }
}

Export-ModuleMember -Function Get-SharePointUserProfilesBulkImportJobStatus

<# ---------------- Helper functions ---------------- #>
Function Initialize {
    Try {
        Remove-Module -Name Import-Dependency -WarningAction SilentlyContinue -ErrorAction SilentlyContinue -Verbose:$false -Force
        Import-Module -Name ".\Modules\Import-Dependency.psm1" -Scope Local -Force -ErrorAction Stop -WarningAction SilentlyContinue -Verbose:$false | Out-Null

        Write-Verbose -Message "Initializing..."

        Foreach ($Module in $GLOBAL:requiredPSModules) {
            Import-Dependency -ModuleName $Module -Verbose
        }
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while initializing. Message: $($_.Exception.Message)"

        Finish-Up -TerminateWithError $true
    }
}

Function Connect {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [ValidatePattern('(http[s]?|[s]?ftp[s]?)(:\/\/)([^\s,]+)')] [string] $SiteUrl,
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

    $SharePointConnection = $null

    Try {
        Remove-Module -Name ConnectTo-SharePointOnline -WarningAction SilentlyContinue -ErrorAction SilentlyContinue -Verbose:$false -Force
        Import-Module -Name ".\Modules\ConnectTo-SharePointOnline.psm1" -Scope Local -Force -ErrorAction Stop -WarningAction SilentlyContinue -Verbose:$false | Out-Null

        $SharePointConnection = ConnectTo-SharePointOnline -SiteUrl $SiteUrl -AuthenticationMethod $AuthenticationMethod -TenantId $TenantId -UserName $UserName -Pass $Pass -ClientId $ClientId -ClientSecret $ClientSecret -CertificateThumbprint $CertificateThumbprint -CertificatePath $CertificatePath -CertificatePass $CertificatePass -ReturnConnection $true -Verbose

        If ($AuthenticationMethod -ne "ManagedIdentity" -and $null -eq $SharePointConnection) {
            throw "Could not connect to SharePoint Online."
        }

        return $SharePointConnection
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while connecting to SharePoint Online. Message: $($_.Exception.Message)"

        Finish-Up -SharePointConnection $SharePointConnection -TerminateWithError $true
    }
}

Function Get-JobStatus {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [ValidatePattern('^[a-f\d]{4}(?:[a-f\d]{4}-){4}[a-f\d]{12}$')] [string] $JobId,
        [Parameter(Mandatory = $false)] [AllowNull()] [PnP.PowerShell.Commands.Base.PnPConnection] $SharePointConnection
    )

    Try {
        Write-Verbose -Message "Fetching sync job status..."
        Write-Debug -Message "Job ID: $($JobId)"

        $JobStatus = $null
        If ($null -ne $SharePointConnection) {
            $JobStatus = Get-PnPUPABulkImportStatus -JobId $JobId -IncludeErrorDetails -Connection $SharePointConnection -ErrorVariable GetPnPUPABulkImportStatusError -Verbose:$false
        }
        Else {
            $JobStatus = Get-PnPUPABulkImportStatus -JobId $JobId -IncludeErrorDetails -ErrorVariable GetPnPUPABulkImportStatusError -Verbose:$false
        }
        
        If ($false -eq [string]::IsNullOrEmpty($GetPnPUPABulkImportStatusError)) {
            throw "Could not fetch the sync job status. Reason: $($GetPnPUPABulkImportStatusError)"
        }

        If ($null -eq $JobStatus) {
            throw "Could not fetch the sync job status."
        }

        Write-Output $JobStatus
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while fetching the status of sync job '$($JobId)'. Message: $($_.Exception.Message)"

        Finish-Up -SharePointConnection $SharePointConnection -TerminateWithError $true
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
       
        Disconnect-PnPOnline -Verbose:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
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