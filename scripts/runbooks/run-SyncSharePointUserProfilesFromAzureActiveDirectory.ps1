<#
    .SYNOPSIS
        Script to synchronize SharePoint User Profiles from Azure Active Directory.

    .DESCRIPTION
        This script allows for the synchronization of SharePoint User Profiles from Azure Active Directory
        based on a configurable property mapping. 
        
        The first run, the configured properties will be synchronized to all users. 
        The script offers the capability to set a sync mode to either 'Full' or 'Delta'. If delta,subsequent runs will 
        only synchronize the delta since the last delta check. If full, the script will always synchronize all users.
        The (delta) job information is saved to a SharePoint list so that the script can utilize this information
        in subsequent runs. The property mapping is saved to a SharePoint document library as a JSON file.
        This means that the property mapping can be changed at any given time by an admin. The script also
        validates whether the property mapping has changed since the last delta check (if sync mode is delta). 
        If this is the case, the previous delta will be abandoned and a full sync will be performed. 
        This is to ensure that configured properties are in sync for all users at all times.

        The script can be run in dry-run mode for testing and development purposes. In this mode, the script will
        run as normal, but will not start the SharePoint User Profile synchronization job and will not create a job info SharePoint list item.

    .FUNCTIONALITY
        - Configure the script via the parameters and global variables, for example
            - Whether the script should run in dry-run mode
            - The authentication method to use
            - The credentials to use for authentication
            - The SharePoint configuration used in the script
            - The sync mode
        - Import the PnP.PowerShell module
        - Install the PnP.PowerShell module if it is not installed
        - Authenticate to SharePoint Online with:
            - Managed Identity
            - Credentials without MFA
            - Credentials with MFA
            - Client secret
            - Certificate thumbprint
            - Certificate path
        - (If sync mode is delta) Get the last delta check information from the JobInfo SharePoint list
        - (If sync mode is delta) If the last delta check was not successful or is still in progress, the script will terminate
          to avoid concurrent of in errors resulting sync jobs
        - Get the property mapping from the Config SharePoint document library
        - (If sync mode is delta) Validate whether the property mapping has changed since the last delta check
        - Based on the sync mode, either get all or just the delta users from Azure Active Directory to sync
        - Create the SharePoint User Profile synchronization job if there are users to sync
        - Create a new JobInfo SharePoint list item with the information of the new job
        - Add the sync input file as attachment to the JobInfo SharePoint list item
        - Disconnect and dispose the session
        - Error handling

    .PARAMETER SyncMode <string> [required]
        The sync mode to use. Either 'Full' or 'Delta' (default). If delta, subsequent runs will only synchronize 
        the delta since the last delta check. If full, the script will synchronize all users.

    .PARAMETER DatabaseSiteUrl <string> [required]
        The SharePoint site URL where the SharePoint User Profile database is located.
        Example: https://contoso.sharepoint.com/sites/contoso

    .PARAMETER ConfigListServerRelativeUri <string> [required]
        The server relative URI of the SharePoint document library where the property mapping JSON is stored
        and where the script can place the sync job input file.
        Example: /sites/contoso/UserProfileSyncConfig
    
    .PARAMETER WipFolderListRelativeUri <string> [required]
        The list relative URI of the SharePoint document library folder where the script can place the sync job input file.
        Example: /sites/contoso/UserProfileSyncConfig/WIP
    
    .PARAMETER PropertyMappingFileListRelativePath <string> [required]
        The list relative path to the property mapping JSON file.
        Example: /sites/contoso/UserProfileSyncConfig/PropertyMapping/property-mapping.json
    
    .PARAMETER JobListName <string> [required]
        The name of the SharePoint list where the script can store the job information.
        Example: User Profile Sync Jobs (default)
    
    .PARAMETER AuthenticationMethod <string> [required]
        The authentication method to use to connect to SharePoint Online. 
        Either ManagedIdentity (default), CredentialsWithoutMFA, CredentialsWithMFA, ClientSecret, CertificateThumbprint or CertificatePath.
    
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
    
    .PARAMETER DryRun <boolean> [optional]
        Whether the script should run in dry-run mode. In this mode, the script will run as normal, 
        but will not start the SharePoint User Profile synchronization job and will not create a 
        job info SharePoint list item. 
        Default: false
    
    .OUTPUTS
        {
            "Id": "<string | null>",
            "State": "<string>",
            "NumberOfUsers": "<int | null>",
            "PropertyMapping": "<object>",
            "SourceUri": "<string | null>",
            "LogFolderUri": "<string | null>",
            "Error": "<string>",
            "Message": "<string | null>"
        }
    
    .NOTES
        Version: 1.0.0
        Requires:
            - PnP.PowerShell v1.12 - https://www.powershellgallery.com/packages/PnP.PowerShell/1.12.0
            - PowerShell7 - https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.3#installing-the-msi-package
            - SharePoint Online User.Read.All and Sites.FullControl.All permissions 
            - Graph User.Read.All permissions
            - SharePoint list and document library to store the job information and property mapping in (see also the accomanying PnP provisioning template)

    .LINK
        Developers:
            - Lennart de Waart
                - GitHub: https://github.com/cupo365
                - LinkedIn: https://www.linkedin.com/in/lennart-dewaart/

        Inspired by:
            - https://learn.microsoft.com/en-us/sharepoint/dev/solution-guidance/bulk-user-profile-update-api-for-sharepoint-online
            - https://pnp.github.io/powershell/cmdlets/Sync-PnPSharePointUserProfilesFromAzureActiveDirectory.html
            - https://pnp.github.io/powershell/cmdlets/Get-PnPAzureADUser.html
#>

<# ------------------- Parameters ------------------- #>
Param (
    [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [ValidateSet("Full", "Delta")] [string] $SyncMode = "Delta",
    [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [ValidatePattern('(http[s]?|[s]?ftp[s]?)(:\/\/)([^\s,]+)')] [string] $DatabaseSiteUrl = 'https://cupo365.sharepoint.com/sites/aad2spup',
    [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [string] $ConfigListServerRelativeUri = "UserProfileSyncConfig",
    [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [string] $WipFolderListRelativeUri = "/WIP",
    [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [string] $PropertyMappingFileListRelativePath = "/PropertyMapping/property-mapping.json",
    [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [string] $JobListName = "User Profile Sync Jobs",
    [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [ValidateSet("ManagedIdentity", "CredentialsWithoutMFA", "CredentialsWithMFA", "ClientSecret", "CertificateThumbprint", "CertificatePath")] [string] $AuthenticationMethod = "CredentialsWithMFA",
    [Parameter(Mandatory = $false)] [ValidatePattern('^[a-f\d]{4}(?:[a-f\d]{4}-){4}[a-f\d]{12}$')] [string] $TenantId = $null,
    [Parameter(Mandatory = $false)] [mailaddress] $UserName = $null,
    [Parameter(Mandatory = $false)] [string] $Pass = $null,
    [Parameter(Mandatory = $false)] [ValidatePattern('^[a-f\d]{4}(?:[a-f\d]{4}-){4}[a-f\d]{12}$')] [string] $ClientId = $null,
    [Parameter(Mandatory = $false)] [string] $ClientSecret = $null,
    [Parameter(Mandatory = $false)] [string] $CertificateThumbprint = $null,
    [Parameter(Mandatory = $false)] [string] $CertificatePath = $null,
    [Parameter(Mandatory = $false)] [string] $CertificatePass = $null,
    [Parameter(Mandatory = $false)] [boolean] $DryRun = $false
)

<# ------------------ Requirements ------------------ #>
#Requires -Version 7
##Requires -Modules @{ModuleName='PnP.PowerShell';ModuleVersion='1.12.0'}

<# --------------- Runtime variables ---------------- #>
$ErrorActionPreference = 'Continue'
$ProgressPreference = 'Continue'
$VerbosePreference = 'Continue'
$DebugPreference = 'SilentlyContinue' # 'Continue'

<# ---------------- Global variables ---------------- #>
$GLOBAL:requiredPSModules = @("PnP.PowerShell")
$GLOBAL:dryRun = $DryRun ?? $false
$GLOBAL:deltaToken = $null
$GLOBAL:jobListName = $false -eq [string]::IsNullOrEmpty($JobListName) ? $JobListName : "User Profile Sync Jobs"
$GLOBAL:azureRunAsConnectionName = "AzureRunAsConnection"
$GLOBAL:azureRunAsCredentialName = "AzureRunAsCredential"
$GLOBAL:partitionKey = "run-SyncSharePointUserProfilesFromAzureActiveDirectory"
$GLOBAL:output = [PSCustomObject]@{
    Id              = $null;
    State           = "Error"; # Set default
    NumberOfUsers   = $null;
    PropertyMapping = $null;
    SourceUri       = $null;
    LogFolderUri    = $null;
    Error           = "InternalError"; # Set default
    Message         = $null;
}

<# ---------------- Helper functions ---------------- #>
Function Import-PSModule {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $ModuleName,
        [Parameter(Mandatory = $false)] [boolean] $AttemptedBefore
    )

    Try {
        Write-Verbose -Message "Importing $($ModuleName) module..."
        $VerbosePreference = "SilentlyContinue" # Ignore verbose import messages that significantly slow down the Runbook
        Import-Module -Name $ModuleName -Scope Local -Force -ErrorAction Stop -Verbose:$false | Out-Null
        $VerbosePreference = "Continue" # Restore verbose preference
    }
    Catch [Exception] {
        $VerbosePreference = "Continue" # Restore verbose preference
        Write-Warning -Message "$($ModuleName) was not found."

        # Prevent looping, so stop after one install attempt
        If ($true -eq $AttemptedBefore) {
            $GLOBAL:output.Message = "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to import the module '$($ModuleName)'. Message: $($_.Exception.Message)."
            Write-Error -Message  $GLOBAL:output.Message

            Write-Verbose -Message "Terminating program. Reason: could not import dependend modules."

            Finish-Up -TerminateWithError $true
        }
        Else {
            Install-PSModule -ModuleName $ModuleName
        }
    }
}

Function Install-PSModule {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $ModuleName
    )

    Try {
        # Install module
        Write-Verbose -Message "Installing $($ModuleName) module..."
        Install-Module -Name $ModuleName -Scope CurrentUser -AllowClobber -Force -ErrorAction Stop -Verbose:$false | Out-Null

        # Installing inherently imports the module, so no need to call Import-PSModule again
        #Import-PSModule -ModuleName $ModuleName -AttemptedBefore $true
    }
    Catch [Exception] {
        $GLOBAL:output.Message = "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to install the module '$($ModuleName)'. Message: $($_.Exception.Message)."
        Write-Error -Message $GLOBAL:output.Message

        Write-Verbose -Message "Terminating program. Reason: could not install dependend modules."

        Finish-Up -TerminateWithError $true
    }
}

Function ConnectTo-SharePointOnlineWithManagedIdentity {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [ValidatePattern('(http[s]?|[s]?ftp[s]?)(:\/\/)([^\s,]+)')] [string] $SiteUrl
    )
    
    Try {
        # Commented for enhanced performance
        #Write-Debug -Message "SiteUrl: $($SiteUrl)"
        Connect-PnPOnline -Url $SiteUrl -ManagedIdentity -ErrorVariable SharePointConnectionError -Verbose:$false | Out-Null

        If ($false -eq [string]::IsNullOrEmpty($SharePointConnectionError)) {
            throw "Could not connect to SharePoint Online. Reason: $($SharePointConnectionError)."
        }
        
        return $true
    }
    Catch [Exception] {
        $GLOBAL:output.Message = "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to connect to SharePoint Online site '$($SiteUrl)' with managedidentity. Message: $($_.Exception.Message)"
        Write-Error -Message $GLOBAL:output.Message

        If ($GLOBAL:dryRun -ne $true) {
            # Log error to SharePoint
            Add-JobListItem
        }

        Finish-Up -TerminateWithError $true
    }
}

Function ConnectTo-SharePointOnlineWithCredentials {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [ValidatePattern('(http[s]?|[s]?ftp[s]?)(:\/\/)([^\s,]+)')] [string] $SiteUrl,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [mailaddress] $UserName,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [securestring] $Pass,
        [Parameter(Mandatory = $true)] [boolean] $WithMFA
    )
    
    Try {
        # Commented for enhanced performance
        # Write-Debug -Message "SiteUrl: $($SiteUrl)"
        # Write-Debug -Message "WithMFA: $($WithMFA)"

        $SharePointConnection = $null

        If ($WithMFA -eq $true) {
            Write-Verbose -Message "If the authentication dialog doesn't come up, please check the background (known bug in VS Code)."
            $SharePointConnection = Connect-PnPOnline -Url $SiteUrl -Interactive -ReturnConnection -ErrorVariable SharePointConnectionError -Verbose:$false
        }
        Else {
            Write-Verbose -Message "Fetching Run As credential '$($GLOBAL:azureRunAsCredentialName)' from the environment..."
            $RunAsCredential = $null
            Try {
                $RunAsCredential = Get-AutomationPSCredential -Name $GLOBAL:azureRunAsCredentialName -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -Verbose:$false
            }
            Catch [Exception] {
                # Do nothing
            }

            If ($null -eq $RunAsCredential) {
                Write-Verbose -Message "Could not fetch a Run As credential from the environment. Attempting to create a Run As credential from parameter input..."
                $RunAsCredential = @{
                    "UserName" = $UserName
                    "Password" = $Pass
                }
            }

            If ($null -ne $RunAsCredential -and $false -eq [string]::IsNullOrEmpty($RunAsCredential.UserName) -and $false -eq [string]::IsNullOrEmpty($RunAsCredential.Password)) {
                # Commented for enhanced performance
                # Write-Debug -Message "RunAsCredential: $(@{
                #     "UserName" = $RunAsCredential.UserName
                #     "Password" = ([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($RunAsCredential.Password)))
                # } | ConvertTo-Json)"
                $creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $RunAsCredential.UserName, $RunAsCredential.Password
                $SharePointConnection = Connect-PnPOnline -Url $SiteUrl -Credential $creds -ReturnConnection -ErrorVariable SharePointConnectionError -Verbose:$false
            }
            Else {
                If ($false -eq [string]::IsNullOrEmpty($RunAsCredential.UserName) -and $true -eq [string]::IsNullOrEmpty($RunAsCredential.Password)) {
                    throw "Cannot connect to SharePoint with credentials. Reason: password is not provided."
                }
                Elseif ($true -eq [string]::IsNullOrEmpty($RunAsCredential.UserName) -and $false -eq [string]::IsNullOrEmpty($RunAsCredential.Password)) {
                    throw "Cannot connect to SharePoint with credentials. Reason: username is not provided."
                }
                Else {
                    throw "Cannot connect to SharePoint with credentials. Reason: username and password are not provided."
                }
            }
        }

        If ($false -eq [string]::IsNullOrEmpty($SharePointConnectionError)) {
            throw "Could not connect to SharePoint Online. Reason: $($SharePointConnectionError)."
        }

        return $SharePointConnection
    }
    Catch [Exception] {
        $GLOBAL:output.Message = "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to connect to SharePoint Online site '$($SiteUrl)' with credentials. Message: $($_.Exception.Message)"
        Write-Error -Message $GLOBAL:output.Message

        If ($GLOBAL:dryRun -ne $true) {
            # Log error to SharePoint
            Add-JobListItem
        }

        Finish-Up -TerminateWithError $true
    }
}

Function ConnectTo-SharePointOnlineWithClientSecret {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [ValidatePattern('(http[s]?|[s]?ftp[s]?)(:\/\/)([^\s,]+)')] [string] $SiteUrl,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $ClientId,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $ClientSecret
    )
    
    Try {
        $SharePointConnection = $null
        
        Write-Verbose -Message "Fetching Run As credential '$($GLOBAL:azureRunAsCredentialName)' from the environment..."
        $RunAsCredential = $null
        Try {
            $RunAsCredential = Get-AutomationPSCredential -Name $GLOBAL:azureRunAsCredentialName -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -Verbose:$false
        }
        Catch [Exception] {
            # Do nothing
        }

        If ($null -eq $RunAsCredential) {
            Write-Verbose -Message "Could not fetch a Run As credential from the environment. Attempting to create a Run As credential from parameter input..."
            $RunAsCredential = @{
                "UserName" = $ClientId
                "Password" = $(ConvertTo-SecureString $ClientSecret -AsPlainText -Force)
            }
        }

        If ($null -ne $RunAsCredential -and $false -eq [string]::IsNullOrEmpty($RunAsCredential.UserName) -and $false -eq [string]::IsNullOrEmpty($RunAsCredential.Password)) {
            # Commented for enhanced performance
            # Write-Debug -Message "SiteUrl: $($SiteUrl)"
            # Write-Debug -Message "RunAsCredential: $(@{
            #     "UserName" = $RunAsCredential.UserName
            #     "Password" = ([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($RunAsCredential.Password)))
            # } | ConvertTo-Json)"
            $SharePointConnection = Connect-PnPOnline -Url $SiteUrl -ClientId $RunAsCredential.UserName -ClientSecret ([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($RunAsCredential.Password))) -ReturnConnection -ErrorVariable SharePointConnectionError -WarningAction Ignore -Verbose:$false
        }
        Else {
            If ($false -eq [string]::IsNullOrEmpty($RunAsCredential.UserName) -and $true -eq [string]::IsNullOrEmpty($RunAsCredential.Password)) {
                throw "Cannot connect to SharePoint with client secret. Reason: client secret is not provided."
            }
            Elseif ($true -eq [string]::IsNullOrEmpty($RunAsCredential.UserName) -and $false -eq [string]::IsNullOrEmpty($RunAsCredential.Password)) {
                throw "Cannot connect to SharePoint with client secret. Reason: client id is not provided."
            }
            Else {
                throw "Cannot connect to SharePoint with client secret. Reason: client id and client secret are not provided."
            }
        }
        
        If ($false -eq [string]::IsNullOrEmpty($SharePointConnectionError)) {
            throw "Could not connect to SharePoint Online. Reason: $($SharePointConnectionError)."
        }

        return $SharePointConnection
    }
    Catch [Exception] {
        $GLOBAL:output.Message = "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to connect to SharePoint Online site '$($SiteUrl)' with clientsecret. Message: $($_.Exception.Message)"
        Write-Error -Message $GLOBAL:output.Message

        If ($GLOBAL:dryRun -ne $true) {
            # Log error to SharePoint
            Add-JobListItem
        }

        Finish-Up -TerminateWithError $true
    }
}

Function ConnectTo-SharePointOnlineWithCertificateThumbprint {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [ValidatePattern('(http[s]?|[s]?ftp[s]?)(:\/\/)([^\s,]+)')] [string] $SiteUrl,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $TenantId,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $ClientId,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $CertificateThumbprint
    )
    
    Try {
        $SharePointConnection = $null
        
        Write-Verbose -Message "Fetching Run As connection '$($GLOBAL:azureRunAsConnectionName)' from the environment..."
        $RunAsConnection = $null
        Try {
            $RunAsConnection = Get-AutomationConnection -Name $GLOBAL:azureRunAsConnectionName -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -Verbose:$false
        }
        Catch [Exception] {
            # Do nothing
        }

        If ($null -eq $RunAsConnection) {
            Write-Verbose -Message "Could not fetch a Run As connection from the environment. Attempting to create a Run As connection from parameter input..."
            $RunAsConnection = @{
                "SubscriptionId"        = $null # Not used for authentication
                "ApplicationId"         = $ClientId
                "CertificateThumbprint" = $CertificateThumbprint
                "TenantId"              = $TenantId
            }
        }

        If ($null -ne $RunAsConnection -and $false -eq [string]::IsNullOrEmpty($RunAsConnection.ApplicationId) -and $false -eq [string]::IsNullOrEmpty($RunAsConnection.CertificateThumbprint) -and $false -eq [string]::IsNullOrEmpty($RunAsConnection.TenantId)) {
            # Commented for enhanced performance
            # Write-Debug -Message "SiteUrl: $($SiteUrl)"
            # Write-Debug -Message "RunAsConnection: $($RunAsConnection | ConvertTo-Json)"
            $SharePointConnection = Connect-PnPOnline -Url $SiteUrl -ClientId $RunAsConnection.ApplicationId -Tenant $RunAsConnection.TenantId -Thumbprint $RunAsConnection.CertificateThumbprint -ReturnConnection -ErrorVariable SharePointConnectionError -Verbose:$false
        }
        Else {
            throw "Cannot connect to SharePoint with certificate thumbprint. Reason: one of the following parameters is not provided: ClientId, CertificateThumbprint, TenantId."
        }

        If ($false -eq [string]::IsNullOrEmpty($SharePointConnectionError)) {
            throw "Could not connect to SharePoint Online. Reason: $($SharePointConnectionError)."
        }

        return $SharePointConnection
    }
    Catch [Exception] {
        $GLOBAL:output.Message = "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to connect to SharePoint Online site '$($SiteUrl)' with certificatethumbprint. Message: $($_.Exception.Message)"
        Write-Error -Message $GLOBAL:output.Message

        If ($GLOBAL:dryRun -ne $true) {
            # Log error to SharePoint
            Add-JobListItem
        }

        Finish-Up -TerminateWithError $true
    }
}

Function ConnectTo-SharePointOnlineWithCertificatePath {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [ValidatePattern('(http[s]?|[s]?ftp[s]?)(:\/\/)([^\s,]+)')] [string] $SiteUrl,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [ValidatePattern('^[a-f\d]{4}(?:[a-f\d]{4}-){4}[a-f\d]{12}$')] [string] $TenantId,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [ValidatePattern('^[a-f\d]{4}(?:[a-f\d]{4}-){4}[a-f\d]{12}$')] [string] $ClientId,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $CertificatePath,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [securestring] $CertificatePass
    )

    Try {
        $SharePointConnection = $null
        
        If ($false -eq [string]::IsNullOrEmpty($ClientId) -and $false -eq [string]::IsNullOrEmpty($CertificatePass) -and $false -eq [string]::IsNullOrEmpty($CertificatePath) -and $false -eq [string]::IsNullOrEmpty($TenantId)) {
            # Commented for enhanced performance
            # Write-Debug -Message "SiteUrl: $($SiteUrl)"
            # Write-Debug -Message "TenantId: $($TenantId)"
            # Write-Debug -Message "ClientId: $($ClientId)"
            # Write-Debug -Message "CertificatePath: $($CertificatePath)"
            # Write-Debug -Message "CertificatePass: $(([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($CertificatePass))))"
            
            If ($CertificatePath | Test-Path) {
                $SharePointConnection = Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -CertificatePath $CertificatePath -CertificatePassword $CertificatePass -Tenant $TenantId -ReturnConnection -ErrorVariable SharePointConnectionError -Verbose:$false
            }
            Else {
                throw "Cannot connect to SharePoint with certificate path. Reason: the provided certificate path is invalid or non existent."
            }
        }
        Else {
            throw "Cannot connect to SharePoint with certificate path. Reason: one of the following parameters is not provided: ClientId, CertificatePass, CertificatePath, TenantId."
        }

        If ($false -eq [string]::IsNullOrEmpty($SharePointConnectionError)) {
            throw "Could not connect to SharePoint Online. Reason: $($SharePointConnectionError)."
        }

        return $SharePointConnection
    }
    Catch [Exception] {
        $GLOBAL:output.Message = "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to connect to SharePoint Online site '$($SiteUrl)' with certificatepath. Message: $($_.Exception.Message)"
        Write-Error -Message $GLOBAL:output.Message

        If ($GLOBAL:dryRun -ne $true) {
            # Log error to SharePoint
            Add-JobListItem
        }

        Finish-Up -TerminateWithError $true
    }
}

Function ConnectTo-SharePointOnline {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [ValidatePattern('(http[s]?|[s]?ftp[s]?)(:\/\/)([^\s,]+)')] [string] $SiteUrl,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [ValidateSet("ManagedIdentity", "CredentialsWithoutMFA", "CredentialsWithMFA", "ClientSecret", "CertificateThumbprint", "CertificatePath")] [string] $AuthenticationMethod,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $TenantId,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [mailaddress] $UserName,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [securestring] $Pass,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $ClientId,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $ClientSecret,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $CertificateThumbprint,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $CertificatePath,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [securestring] $CertificatePass
    )

    Try {
        Write-Verbose -Message "Connecting to SharePoint..."
        $SharePointConnection = $null
        Switch ($AuthenticationMethod) {
            "ManagedIdentity" {
                ConnectTo-SharePointOnlineWithManagedIdentity -SiteUrl $SiteUrl
                return $null
            }
            "CredentialsWithoutMFA" {
                $SharePointConnection = ConnectTo-SharePointOnlineWithCredentials -SiteUrl $SiteUrl -UserName $UserName -Pass $Pass -WithMFA $false
            }
            "CredentialsWithMFA" {
                $SharePointConnection = ConnectTo-SharePointOnlineWithCredentials -SiteUrl $SiteUrl -UserName $UserName -Pass $Pass -WithMFA $true
            }
            "ClientSecret" {
                $SharePointConnection = ConnectTo-SharePointOnlineWithClientSecret -SiteUrl $SiteUrl -ClientId $ClientId -ClientSecret $ClientSecret
            }
            "CertificateThumbprint" {
                $SharePointConnection = ConnectTo-SharePointOnlineWithCertificateThumbprint -SiteUrl $SiteUrl -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
            }
            "CertificatePath" {
                $SharePointConnection = ConnectTo-SharePointOnlineWithCertificatePath -SiteUrl $SiteUrl -TenantId $TenantId -ClientId $ClientId -CertificatePath $CertificatePath -CertificatePass $CertificatePass
            }
        }

        If ($AuthenticationMethod -ne "ManagedIdentity" -and $null -eq $SharePointConnection) {
            throw "Could not connect to SharePoint Online."
        }

        return $SharePointConnection
    } 
    Catch [Exception] {
        $GLOBAL:output.Message = "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to connect to SharePoint Online site '$($SiteUrl)' with $($AuthenticationMethod.ToLower()). Message: $($_.Exception.Message)"
        Write-Error -Message $GLOBAL:output.Message

        If ($GLOBAL:dryRun -ne $true) {
            # Log error to SharePoint
            Add-JobListItem -JobListName $GLOBAL:jobListName
        }

        Finish-Up -TerminateWithError $true
    }
}

Function Get-PropertyMapping {  
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $FileServerRelativePath,
        [Parameter(Mandatory = $false)] [AllowNull()] [PnP.PowerShell.Commands.Base.PnPConnection] $SharePointConnection
    )

    Try {
        Write-Verbose -Message "Fetching property mapping..."
        # Commented for enhanced performance
        #Write-Debug -Message "File server relative path: $($FileServerRelativePath)"

        $PropertyMapping = $null
        If ($null -ne $SharePointConnection) {
            $PropertyMapping = Get-PnPFile -Url $FileServerRelativePath -AsString -Connection $SharePointConnection -ErrorVariable PropertyMappingGetError -Verbose:$false
        }
        Else {
            $PropertyMapping = Get-PnPFile -Url $FileServerRelativePath -AsString -ErrorVariable PropertyMappingGetError -Verbose:$false
        }

        If ($false -eq [string]::IsNullOrEmpty($PropertyMappingGetError)) {
            throw "Could not fetch property mapping. Reason: $($PropertyMappingGetError)."
        }

        If ($null -eq $PropertyMapping) {
            throw "Could not fetch property mapping."
        }

        # Commented for enhanced performance
        #Write-Debug -Message "PropertyMapping: $($PropertyMapping | ConvertFrom-Json | Out-String)"

        return $PropertyMapping | ConvertFrom-Json
    }
    Catch [Exception] {
        $GLOBAL:output.Message = "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to fetch the property mapping. Message: $($_.Exception.Message)"
        Write-Error -Message $GLOBAL:output.Message
        
        If ($GLOBAL:dryRun -ne $true) {
            # Log error to SharePoint
            Add-JobListItem -JobListName $GLOBAL:jobListName -SharePointConnection $SharePointConnection
        }

        Finish-Up -SharePointConnection $SharePointConnection -TerminateWithError $true
    }
}

Function Get-LastDeltaCheck {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $JobListName,
        [Parameter(Mandatory = $false)] [AllowNull()] [PnP.PowerShell.Commands.Base.PnPConnection] $SharePointConnection
    )

    Try {
        Write-Verbose -Message "Fetching last delta check..."
        # Commented for enhanced performance
        #Write-Debug -Message "Job list name: $($JobListName)"

        $CamlQuery = "@
        <View Scope='RecursiveAll'>
            <ViewFields>
                <FieldRef Name='ID'/>
                <FieldRef Name='JobId'/>
                <FieldRef Name='Created'/>
                <FieldRef Name='State'/>
                <FieldRef Name='NoUsers'/>
                <FieldRef Name='PropertyMapping'/>
                <FieldRef Name='DeltaToken'/>
                <FieldRef Name='Error'/>
                <FieldRef Name='SourceUri'/>
                <FieldRef Name='LogFolderUri'/>
                <FieldRef Name='Message'/>
            </ViewFields>
            <Query>
                <Where>
                   <Eq>
                        <FieldRef Name='Title' />
                        <Value Type='Text'>$($GLOBAL:partitionKey)</Value>
                    </Eq>
                </Where>
                <OrderBy>
                    <FieldRef Name='ID' Ascending='False' />
                </OrderBy>
            </Query>
        </View>"

        $LastDeltaCheck = $null
        If ($null -ne $SharePointConnection) {
            $LastDeltaCheck = (Get-PnPListItem -List $JobListName -Query $CamlQuery -Connection $SharePointConnection -ErrorVariable GetPnPListItemError -Verbose:$false | Select-Object -First 1).FieldValues
        }
        Else {
            $LastDeltaCheck = (Get-PnPListItem -List $JobListName -Query $CamlQuery -ErrorVariable GetPnPListItemError -Verbose:$false | Select-Object -First 1).FieldValues
        }

        If ($false -eq [string]::IsNullOrEmpty($GetPnPListItemError)) {
            throw "Could not fetch last delta check. Reason: $($GetPnPListItemError)."
        }

        # Commented for enhanced performance
        #Write-Debug -Message "Last delta check: $(($LastDeltaCheck | ConvertTo-Json -Depth 10 | Out-String))"

        return $LastDeltaCheck        
    }
    Catch [Exception] {
        $GLOBAL:output.Message = "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to fetch the last delta check. Message: $($_.Exception.Message)"
        Write-Error -Message $GLOBAL:output.Message
        
        If ($GLOBAL:dryRun -ne $true) {
            # Log error to SharePoint
            Add-JobListItem -JobListName $GLOBAL:jobListName -SharePointConnection $SharePointConnection
        }

        Finish-Up -SharePointConnection $SharePointConnection -TerminateWithError $true
    }
}

Function Validate-PropertyMapping {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [object] $ReferenceObject,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [object] $DifferenceObject
    )
    
    Try {
        Write-Verbose -Message "Validating whether property mapping has changed since the last delta check..."
        $PropertyMappingHasChanged = $false

        Foreach ($PropertyMappingProperty in $ReferenceObject.PSObject.Properties) {
            $PropertyMappingPropertyComparison = Compare-Object -ReferenceObject $ReferenceObject -DifferenceObject $DifferenceObject -Property $PropertyMappingProperty.Name -IncludeEqual
            
            If ($null -eq $PropertyMappingPropertyComparison -or "==" -ne $PropertyMappingPropertyComparison[0].SideIndicator) {
                $PropertyMappingHasChanged = $true
                Break
            }
        }

        # Commented for enhanced performance
        # Write-Debug -Message "Property mapping reference object: $($ReferenceObject | Out-String)"
        # Write-Debug -Message "Property mapping difference object: $($DifferenceObject | Out-String)"

        return $PropertyMappingHasChanged
    }
    Catch [Exception] {
        $GLOBAL:output.Message = "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to validate whether the property mapping has changed since the last delta check. Message: $($_.Exception.Message)"
        Write-Error -Message $GLOBAL:output.Message
        
        If ($GLOBAL:dryRun -ne $true) {
            # Log error to SharePoint
            Add-JobListItem -JobListName $GLOBAL:jobListName
        }

        Finish-Up -TerminateWithError $true
    }
}

Function Get-DeltaToken {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $JobListName,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [pscustomobject] $PropertyMapping,
        [Parameter(Mandatory = $false)] [AllowNull()] [PnP.PowerShell.Commands.Base.PnPConnection] $SharePointConnection
    )

    Try {
        $LastDeltaCheck = Get-LastDeltaCheck -JobListName $JobListName -SharePointConnection $SharePointConnection
        If ($null -eq $LastDeltaCheck) {
            Write-Verbose -Message "No last delta check found. Performing full sync..."
            return $null
        }

        If ($null -ne $LastDeltaCheck -and ($LastDeltaCheck.State -ne "Succeeded" -or $LastDeltaCheck.Error -ne "NoError")) {
            $GLOBAL:output.State = 'Skipped'
            $GLOBAL:output.Error = 'DeltaCheckError'
            throw "Last delta check was not successful or has not finished yet. Please check the lastest delta check in the SharePoint job list. Terminating run to avoid concurrent or in error resulting sync jobs..."
        }

        If ($true -eq [string]::IsNullOrEmpty($LastDeltaCheck.DeltaToken) -or $true -eq [string]::IsNullOrEmpty($LastDeltaCheck.PropertyMapping)) {
            Write-Verbose -Message "Found last delta check, but either the delta token or the property mapping is invalid. Performing full sync..."
            return $null
        }

        $PropertyMappingChangedSinceLastDeltaCheck = Validate-PropertyMapping -ReferenceObject $PropertyMapping -DifferenceObject ($LastDeltaCheck.PropertyMapping | ConvertFrom-Json)

        If ($PropertyMappingChangedSinceLastDeltaCheck) {
            Write-Verbose -Message "Property mapping has changed since the last delta check. Abandoning delta token. Performing full sync..."
            return $null
        }
        Else {
            Write-Verbose -Message "Property mapping has not changed since the last delta check. Continuing..."
        }

        # Commented for enhanced performance
        #Write-Debug -Message "Delta token: $($LastDeltaCheck.DeltaToken)"

        return $LastDeltaCheck.DeltaToken
    } 
    Catch [Exception] {
        $GLOBAL:output.Message = "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to fetch the delta token. Message: $($_.Exception.Message)"
        Write-Error -Message $GLOBAL:output.Message
        
        If ($GLOBAL:dryRun -ne $true) {
            # Log error to SharePoint
            Add-JobListItem -JobListName $GLOBAL:jobListName -SharePointConnection $SharePointConnection
        }

        Finish-Up -SharePointConnection $SharePointConnection -TerminateWithError $true
    }
}

Function Get-AzureADUsersToSync {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [array] $SelectProperties,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $DeltaToken,
        [Parameter(Mandatory = $false)] [AllowNull()] [PnP.PowerShell.Commands.Base.PnPConnection] $SharePointConnection
    )

    Try {
        # Commented for enhanced performance
        #Write-Debug -Message "Select properties: $($SelectProperties)"

        # There is a bug in PowerShell 5.1 for the command 'Get-PnPAzureADUser'
        # The bug pertains selecting properties, so use PowerShell 7
        # Source: https://github.com/pnp/powershell/issues/1979
        $AzureADUsers = $null
        If ($null -eq $DeltaToken) {
            Write-Verbose -Message "Fetching all users from Azure Active Directory..."
            If ($null -ne $SharePointConnection) {
                $AzureADUsers = Get-PnPAzureADUser -Delta -Select $SelectProperties -Connection $SharePointConnection -ErrorVariable GetPnPAzureADUserError -Verbose:$false
            }
            Else {
                $AzureADUsers = Get-PnPAzureADUser -Delta -Select $SelectProperties -ErrorVariable GetPnPAzureADUserError -Verbose:$false
            }
        }
        Else {
            Write-Verbose -Message "Fetching delta users from Azure Active Directory..."
            If ($null -ne $SharePointConnection) {
                $AzureADUsers = Get-PnPAzureADUser -Delta -DeltaToken $DeltaToken -Select $SelectProperties -Connection $SharePointConnection -ErrorVariable GetPnPAzureADUserError -Verbose:$false
            }
            Else {
                $AzureADUsers = Get-PnPAzureADUser -Delta -DeltaToken $DeltaToken -Select $SelectProperties -ErrorVariable GetPnPAzureADUserError -Verbose:$false
            }
        }

        If ($false -eq [string]::IsNullOrEmpty($GetPnPAzureADUserError)) {
            throw "Could not get users from Azure Active Directory. Reason: $($GetPnPAzureADUserError)"
        }

        If ($null -eq $AzureADUsers -or $null -eq $AzureADUsers.Users -or $AzureADUsers.Users.Count -eq 0) {
            Write-Verbose -Message "Found 0 users to sync. Skipping sync job..."
            $GLOBAL:output.NumberOfUsers = 0

            # Overwrite default values
            $GLOBAL:output.State = 'Succeeded'
            $GLOBAL:output.Error = 'NoError'
        }
        Else {
            If ($AzureADUsers.Users.Count -eq 1) {
                Write-Verbose -Message "Found $($AzureADUsers.Users.Count) user to sync."
            }
            Else {
                Write-Verbose -Message "Found $($AzureADUsers.Users.Count) users to sync."
            }
            $GLOBAL:output.NumberOfUsers = $AzureADUsers.Users.Count
            $GLOBAL:deltaToken = $AzureADUsers.DeltaToken
        }

        # Commented for enhanced performance
        #Write-Debug -Message "AzureADUsers: $($AzureADUsers | ConvertTo-Json -Depth 10 | Out-String)"

        return $AzureADUsers
    }
    Catch [Exception] {
        $GLOBAL:output.Message = "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to fetch the Azure AD users to sync. Message: $($_.Exception.Message)"
        Write-Error -Message $GLOBAL:output.Message
        
        If ($GLOBAL:dryRun -ne $true) {
            # Log error to SharePoint
            Add-JobListItem -JobListName $GLOBAL:jobListName -SharePointConnection $SharePointConnection
        }

        Finish-Up -SharePointConnection $SharePointConnection -TerminateWithError $true
    }
}

Function New-SharePointUserProfileSyncJob {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [array] $UsersToSync,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [hashtable] $PropertyMapping,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $SyncFileServerRelativeUri,
        [Parameter(Mandatory = $false)] [AllowNull()] [PnP.PowerShell.Commands.Base.PnPConnection] $SharePointConnection
    )

    Try {
        Write-Verbose -Message "Creating new SharePoint User Profile sync job..."
        # Commented for enhanced performance
        #Write-Debug -Message "Sync file server relative uri: $($SyncFileServerRelativeUri)"

        $SyncJob = $null
        If ($GLOBAL:dryRun -eq $true) {
            If ($null -ne $SharePointConnection) {
                $SyncJob = Sync-PnPSharePointUserProfilesFromAzureActiveDirectory -UserProfilePropertyMapping $PropertyMapping -Users $UsersToSync -Folder $SyncFileServerRelativeUri -WhatIf -Connection $SharePointConnection -ErrorVariable SyncJobError
            }
            Else {
                $SyncJob = Sync-PnPSharePointUserProfilesFromAzureActiveDirectory -UserProfilePropertyMapping $PropertyMapping -Users $UsersToSync -Folder $SyncFileServerRelativeUri -WhatIf -ErrorVariable SyncJobError
            }
        }
        Else {
            If ($null -ne $SharePointConnection) {
                $SyncJob = Sync-PnPSharePointUserProfilesFromAzureActiveDirectory -UserProfilePropertyMapping $PropertyMapping -Users $UsersToSync -Folder $SyncFileServerRelativeUri -Connection $SharePointConnection -ErrorVariable SyncJobError
            }
            Else {
                $SyncJob = Sync-PnPSharePointUserProfilesFromAzureActiveDirectory -UserProfilePropertyMapping $PropertyMapping -Users $UsersToSync -Folder $SyncFileServerRelativeUri -ErrorVariable SyncJobError
            }
        }

        If ($false -eq [string]::IsNullOrEmpty($SyncJobError)) {
            throw "Could not create new SharePoint User Profile sync job. Reason: $($SyncJobError)."
        }

        If ($null -eq $SyncJob) {
            throw "Could not create new SharePoint User Profile sync job."
        }

        $GLOBAL:output.Id = $SyncJob.JobId ? $SyncJob.JobId : $null
        $GLOBAL:output.State = !$GLOBAL:dryRun ? "$([Microsoft.Online.SharePoint.TenantManagement.ImportProfilePropertiesJobState]$SyncJob.State.ToString())" : "WontStart"
        $GLOBAL:output.Error = !$GLOBAL:dryRun ? "$([Microsoft.Online.SharePoint.TenantManagement.ImportProfilePropertiesJobError]$SyncJob.Error.ToString())" : "NoError"
        $GLOBAL:output.Message = $SyncJob.ErrorMessage ? $SyncJob.ErrorMessage : $null
        $GLOBAL:output.SourceUri = $SyncJob.SourceUri ? [uri]::EscapeUriString($SyncJob.SourceUri) : $null
        $GLOBAL:output.LogFolderUri = $SyncJob.LogFolderUri ? [uri]::EscapeUriString($SyncJob.LogFolderUri) : $null

        # Commented for enhanced performance
        #Write-Debug -Message "SharePoint User Profile sync job: $($SyncJob | ConvertTo-Json | Out-String)"

        return $SyncJob
    }
    Catch [Exception] {
        $GLOBAL:output.Message = "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to create the new SharePoint User Profile sync job. Message: $($_.Exception.Message)"
        Write-Error -Message $GLOBAL:output.Message
        
        If ($GLOBAL:dryRun -ne $true) {
            # Log error to SharePoint
            Add-JobListItem -JobListName $GLOBAL:jobListName -SharePointConnection $SharePointConnection
        }

        Finish-Up -SharePointConnection $SharePointConnection -TerminateWithError $true
    }
}

Function Add-SyncFileAsJobListItemAttachment {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [ValidatePattern('(http[s]?|[s]?ftp[s]?)(:\/\/)([^\s,]+)')] [string] $SiteUrl,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $SourceUri,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $JobListName,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [int] $JobListItemId,
        [Parameter(Mandatory = $false)] [AllowNull()] [PnP.PowerShell.Commands.Base.PnPConnection] $SharePointConnection
    )

    Try {
        Write-Verbose -Message "Appending the job input file to the job info list item as attachment..."
        $SyncFileLeafRef = ($SourceUri -split "/") | Select-Object -Last 1

        # Commented for enhanced performance
        #Write-Debug -Message "SyncFileLeafRef: $($SyncFileLeafRef)"
            
        $SyncFileContent = $null
        If ($null -ne $SharePointConnection) {
            $SyncFileContent = Get-PnPFile -Url "$($SourceUri -replace $SiteUrl, '')" -AsString -Connection $SharePointConnection -ErrorVariable SyncFileContentError -ErrorAction SilentlyContinue -Verbose:$false
        }
        Else {
            $SyncFileContent = Get-PnPFile -Url "$($SourceUri -replace $SiteUrl, '')" -AsString -ErrorVariable SyncFileContentError -ErrorAction SilentlyContinue -Verbose:$false
        }

        If ($false -eq [string]::IsNullOrEmpty($SyncFileContentError)) {
            throw "Could not get content of sync file '$($SyncFileLeafRef)'. Reason: $($SyncFileContentError). Continuing..."
        }

        If ($null -eq $SyncFileContent) {
            throw "Could not get content of sync file '$($SyncFileLeafRef)'. Continuing..."
        }

        # Commented for enhanced performance
        #Write-Debug -Message "Sync file content: $($SyncFileContent)"

        $JobListItemAttachment = $null
        If ($null -ne $SharePointConnection) {
            $JobListItemAttachment = Add-PnPListItemAttachment -List $JobListName -Identity $JobListItemId -FileName $SyncFileLeafRef -Content $SyncFileContent -Connection $SharePointConnection -ErrorVariable JobListItemAttachmentError -ErrorAction SilentlyContinue -Verbose:$false
        }
        Else {
            $JobListItemAttachment = Add-PnPListItemAttachment -List $JobListName -Identity $JobListItemId -FileName $SyncFileLeafRef -Content $SyncFileContent -ErrorVariable JobListItemAttachmentError -ErrorAction SilentlyContinue -Verbose:$false
        }

        If ($false -eq [string]::IsNullOrEmpty($JobListItemAttachmentError)) {
            throw "Could not add sync file '$($SyncFileLeafRef)' to the job info list item (ID: $($JobListItemId)) as attachment. Reason: $($JobListItemAttachmentError). Continuing..."
        }

        If ($null -eq $JobListItemAttachment) {
            throw "Could not add sync file '$($SyncFileLeafRef)' to the job info list item (ID: $($JobListItemId)) as attachment. Continuing..."
        }
    }
    Catch [Exception] {
        $GLOBAL:output.Message = "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to create the job info list item. Message: $($_.Exception.Message)"
        Write-Warning -Message $GLOBAL:output.Message
    }
}

Function Add-JobListItem {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $JobListName,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [ValidatePattern('(http[s]?|[s]?ftp[s]?)(:\/\/)([^\s,]+)')] [string] $SiteUrl,
        [Parameter(Mandatory = $false)] [AllowNull()] [PnP.PowerShell.Commands.Base.PnPConnection] $SharePointConnection
    )

    Try {
        Write-Verbose -Message "Creating job list item in SharePoint list '$($JobListName)'..."
        # Commented for enhanced performance
        #Write-Debug -Message "JobListName: $($JobListName)"

        $ListItemValues = @{
            "Title"           = "$($GLOBAL:partitionKey)"; 
            "JobId"           = "$($GLOBAL:output.Id)";
            "State"           = "$($GLOBAL:output.State)";
            "NoUsers"         = $GLOBAL:output.NumberOfUsers;
            "PropertyMapping" = $GLOBAL:output.PropertyMapping ? ($GLOBAL:output.PropertyMapping | ConvertTo-Json -Depth 10) : $null; 
            "DeltaToken"      = "$($GLOBAL:deltaToken)";
            "SourceUri"       = $GLOBAL:output.SourceUri ? "$($GLOBAL:output.SourceUri), Source uri" : $null;
            "LogFolderUri"    = $GLOBAL:output.LogFolderUri ? "$($GLOBAL:output.LogFolderUri), Log folder uri" : $null;
            "Error"           = "$($GLOBAL:output.Error)";
            "Message"         = "$($GLOBAL:output.Message)";
        }

        # Commented for enhanced performance
        #Write-Debug -Message "JobListValues: $($ListItemValues | Format-List | Out-String)"

        $JobListItem = $null
        If ($null -ne $SharePointConnection) {
            $JobListItem = Add-PnPListItem -List $JobListName -Values $ListItemValues -Connection $SharePointConnection -ErrorVariable JobListItemError -Verbose:$false
        }
        Else {
            $JobListItem = Add-PnPListItem -List $JobListName -Values $ListItemValues -ErrorVariable JobListItemError -Verbose:$false
        }

        If ($false -eq [string]::IsNullOrEmpty($JobListItemError)) {
            throw throw "Could not create job info list item. Reason: $($JobListItemError)"
        }

        If ($null -eq $JobListItem) {
            throw "Could not create job info list item."
        }

        # Commented for enhanced performance
        #Write-Debug -Message "Job list item: $($JobListItem | Out-String)"

        If ($null -ne $GLOBAL:output.SourceUri -and $null -ne $SiteUrl) {
            Add-SyncFileAsJobListItemAttachment -SiteUrl $SiteUrl -SourceUri $GLOBAL:output.SourceUri -JobListName $JobListName -JobListItemId $JobListItem.ID -SharePointConnection $SharePointConnection
        }
    }
    Catch [Exception] {
        $GLOBAL:output.Message = "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to create the job info list item. Message: $($_.Exception.Message)"
        Write-Error -Message $GLOBAL:output.Message

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

        Write-Output ($GLOBAL:output | ConvertTo-Json -Depth 10)
        
        If ($TerminateWithError) {
            exit 1 # Terminate the program as failed
        }
    }
}

<# -------------------- Program --------------------- #>
Try {
    If ($GLOBAL:dryRun -eq $true) {
        Write-Verbose -Message "[DRY RUN MODE] - No actual changes will be made to the SharePoint User Profiles and the job info list will not be updated for this run."
    }

    Write-Verbose -Message "Sync mode: $($SyncMode.ToUpper())."

    Foreach ($Module in $GLOBAL:requiredPSModules) {
        Import-PSModule -ModuleName $Module
    }

    $SharePointConnection = ConnectTo-SharePointOnline -SiteUrl $DatabaseSiteUrl -AuthenticationMethod $AuthenticationMethod -TenantId $TenantId -UserName $UserName -Pass ($Pass ? (ConvertTo-SecureString -AsPlainText $Pass -Force) : $null) -ClientId $ClientId -ClientSecret $ClientSecret -CertificateThumbprint $CertificateThumbprint -CertificatePath $CertificatePath -CertificatePass ($CertificatePass ? (ConvertTo-SecureString -AsPlainText $CertificatePass -Force) : $null)
    If ($AuthenticationMethod -eq 'ManagedIdentity') {
        # For some reason, even though the ConnectTo-SharePointOnline function specifically states that
        # if the authentication method is 'ManagedIdentity', it returns $null,
        # somehow it comes out as a System.Object[]. This causes the following actions to fail,
        # so explicity set the connection to $null once again is the auth method is 'ManagedIdentity'
        $SharePointConnection = $null
    }

    $PropertyMappingFileServerRelativePath = [uri]::EscapeUriString("$($ConfigListServerRelativeUri.StartsWith('/') ? '' : '/')$($ConfigListServerRelativeUri)$($PropertyMappingFileListRelativePath.StartsWith('/') ? '' : '/')$($PropertyMappingFileListRelativePath)")
    $GLOBAL:output.PropertyMapping = Get-PropertyMapping -FileServerRelativePath $PropertyMappingFileServerRelativePath -SharePointConnection $SharePointConnection

    Switch ($SyncMode) {
        "Delta" {
            $GLOBAL:deltaToken = Get-Deltatoken -JobListName $GLOBAL:jobListName -PropertyMapping $GLOBAL:output.PropertyMapping -SharePointConnection $SharePointConnection
        }
        "Full" { 
            $GLOBAL:deltaToken = $null
        }
    }

    $AADSelectProperties = @()
    $GLOBAL:output.PropertyMapping.PSObject.Properties | ForEach-Object { $AADSelectProperties += $_.Name }
    $AzureADUsersToSync = Get-AzureADUsersToSync -SelectProperties $AADSelectProperties -DeltaToken $GLOBAL:deltaToken -SharePointConnection $SharePointConnection

    If ($null -ne $AzureADUsersToSync -and $null -ne $AzureADUsersToSync.Users -and $AzureADUsersToSync.Users.Count -gt 0) {
        # The -UserProfilePropertyMapping parameter of Sync-PnPSharePointUserProfilesFromAzureActiveDirectory
        # requires a hashtable, so convert the PSCustomObject $GLOBAL:output.PropertyMapping to a hashtable
        $PropertyMappingHashtable = @{}
        $GLOBAL:output.PropertyMapping.PSObject.Properties | ForEach-Object { $PropertyMappingHashtable[$_.Name] = $_.Value }
        $SyncFileServerRelativeUri = [uri]::UnescapeDataString("$($ConfigListServerRelativeUri.StartsWith('/') ? '' : '/')$($ConfigListServerRelativeUri)$($WipFolderListRelativeUri.StartsWith('/') ? '' : '/')$($WipFolderListRelativeUri)")

        $SyncJob = New-SharePointUserProfileSyncJob -UsersToSync $AzureADUsersToSync.Users -PropertyMapping $PropertyMappingHashtable -SyncFileServerRelativeUri $SyncFileServerRelativeUri -SharePointConnection $SharePointConnection
    }

    If ($GLOBAL:dryRun -ne $true) {
        Add-JobListItem -JobListName $GLOBAL:jobListName -SiteUrl $DatabaseSiteUrl -SharePointConnection $SharePointConnection
    }

    Finish-Up -SharePointConnection $SharePointConnection
}
Catch [Exception] {
    $GLOBAL:output.Message = "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while executing the program. Message: $($_.Exception.Message)"
    Write-Error -Message $GLOBAL:output.Message

    Finish-Up -TerminateWithError $true
}