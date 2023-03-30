<#
    .SYNOPSIS
        Script to update the job status of a SharePoint User Profile sync job.

    .DESCRIPTION
        This script allows you to connect to SharePoint Online with PnP using various authentication methods, like:
            - ManagedIdentity
            - CredentialsWithoutMFA
            - CredentialsWithMFA
            - ClientSecret
            - CertificateThumbprint
            - CertificatePath
        Once authenticated, the script will fetch the status of the sync job with the specified ID.
        The status will be stored in a SharePoint list item, so that the status can be used in other scripts

    .FUNCTIONALITY
        - Configure the script via the parameters and global variables, for example
            - Whether the script should run in dry-run mode
            - The authentication method to use
            - The credentials to use for authentication
            - The SharePoint configuration used in the script
        - Import the PnP.PowerShell module
        - Install the PnP.PowerShell module if it is not installed
        - Authenticate to SharePoint Online with:
            - Managed Identity
            - Credentials without MFA
            - Credentials with MFA
            - Client secret
            - Certificate thumbprint
            - Certificate path
        - Fetch the status of the sync job with the specified ID
        - Get the SharePoint list item ID where the script can store the updated job information
        - Update the status of the sync job in the SharePoint list

    .PARAMETER DatabaseSiteUrl <string> [required]
        The SharePoint site URL where the SharePoint User Profile database is located.
        Example: https://contoso.sharepoint.com/sites/contoso

    .PARAMETER JobId <string> [required]
        The ID of the sync job to fetch the status from.
        Example: 00000000-0000-0000-0000-000000000000
    
    .PARAMETER JobListName <string> [required]
        The name of the SharePoint list where the script can store the updated job information.
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
        Microsoft.Online.SharePoint.TenantManagement.ImportProfilePropertiesJobInfo object
        See also: https://learn.microsoft.com/en-us/previous-versions/office/sharepoint-csom/mt643008(v%3Doffice.15)

        {
            "JobId": "<string>",
            "State": "<string>",
            "SourceUri": "<string | null>",
            "LogFolderUri": "<string | null>",
            "Error": "<string>",
            "ErrorMessage": "<string | null>"
        }
    
    .NOTES
        Version: 1.0.0
        Requires:
            - PnP.PowerShell v1.12.0 - https://www.powershellgallery.com/packages/PnP.PowerShell/1.12.0
            - SharePoint Online Sites.FullControl.All permissions 
            - SharePoint list to store the updated job information

    .LINK
        Developers:
            - Lennart de Waart
                - GitHub: https://github.com/cupo365
                - LinkedIn: https://www.linkedin.com/in/lennart-dewaart/

        Inspired by:
            - https://pnp.github.io/powershell/cmdlets/Get-PnPUPABulkImportStatus.html
            - https://learn.microsoft.com/en-us/previous-versions/office/sharepoint-csom/mt643008(v%3Doffice.15)
#>

<# ------------------- Parameters ------------------- #>
Param (
    [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [ValidatePattern('(http[s]?|[s]?ftp[s]?)(:\/\/)([^\s,]+)')] [string] $DatabaseSiteUrl = 'https://cupo365.sharepoint.com/sites/aad2spup',
    [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [ValidatePattern('^[a-f\d]{4}(?:[a-f\d]{4}-){4}[a-f\d]{12}$')] [string] $JobId = 'b0228679-3415-4b77-88c8-7281f780ab01',
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
    [Parameter(Mandatory = $false)] [boolean] $DryRun = $true
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
$GLOBAL:jobListName = $false -eq [string]::IsNullOrEmpty($JobListName) ? $JobListName : "User Profile Sync Jobs"
$GLOBAL:azureRunAsConnectionName = "AzureRunAsConnection"
$GLOBAL:azureRunAsCredentialName = "AzureRunAsCredential"
$GLOBAL:output = $null

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
            Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to import the module '$($ModuleName)'. Message: $($_.Exception.Message)."

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
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to install the module '$($ModuleName)'. Message: $($_.Exception.Message)."

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
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to connect to SharePoint Online site '$($SiteUrl)' with managedidentity. Message: $($_.Exception.Message)"
        
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
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to connect to SharePoint Online site '$($SiteUrl)' with credentials. Message: $($_.Exception.Message)"

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
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to connect to SharePoint Online site '$($SiteUrl)' with clientsecret. Message: $($_.Exception.Message)"

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
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to connect to SharePoint Online site '$($SiteUrl)' with certificatethumbprint. Message: $($_.Exception.Message)"

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
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to connect to SharePoint Online site '$($SiteUrl)' with certificatepath. Message: $($_.Exception.Message)"

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
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to connect to SharePoint Online site '$($SiteUrl)' with $($AuthenticationMethod.ToLower()). Message: $($_.Exception.Message)"

        Finish-Up -TerminateWithError $true
    }
}

Function Get-SharePointUserProfilesBulkImportJobStatus {  
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [ValidatePattern('^[a-f\d]{4}(?:[a-f\d]{4}-){4}[a-f\d]{12}$')] [string] $JobId,
        [Parameter(Mandatory = $false)] [AllowNull()] [PnP.PowerShell.Commands.Base.PnPConnection] $SharePointConnection
    )

    Try {
        Write-Verbose -Message "Fetching sync job status..."
        # Commented for enhanced performance
        # Write-Debug -Message "Job ID: $($JobId)"

        $JobStatus = $null
        If ($null -ne $SharePointConnection) {
            $JobStatus = Get-PnPUPABulkImportStatus -JobId $JobId -IncludeErrorDetails -Connection $SharePointConnection -ErrorVariable GetPnPUPABulkImportStatusError -Verbose:$false
        }
        Else {
            $JobStatus = Get-PnPUPABulkImportStatus -JobId $JobId -IncludeErrorDetails -ErrorVariable GetPnPUPABulkImportStatusError -Verbose:$false
        }
        
        If ($false -eq [string]::IsNullOrEmpty($GetPnPUPABulkImportStatusError)) {
            throw "Could not fetch the updated sync job status. Reason: $($GetPnPUPABulkImportStatusError)"
        }

        If ($null -eq $JobStatus) {
            throw "Could not fetch the updated sync job status."
        }

        # Commented for enhanced performance
        # Write-Debug -Message "Job status: $($JobStatus.State)"
        # Write-Debug -Message "Job error: $($JobStatus.Error)"
        # Write-Debug -Message "Job error message: $($JobStatus.ErrorMessage)"
        # Write-Debug -Message "Job log folder uri: $($JobStatus.LogFolderUri)"

        return $JobStatus
    }
    Catch [Exception] {
        Write-Error "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to fetch the updated sync job status. Message: $($_.Exception.Message)"

        Finish-Up -SharePointConnection $SharePointConnection -TerminateWithError $true
    }
}

Function Get-JobListItemId {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $JobListName,
        [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [ValidatePattern('^[a-f\d]{4}(?:[a-f\d]{4}-){4}[a-f\d]{12}$')] [string] $JobId,
        [Parameter(Mandatory = $false)] [AllowNull()] [PnP.PowerShell.Commands.Base.PnPConnection] $SharePointConnection
    )

    Try {
        Write-Verbose -Message "Fetching job list item id from SharePoint..."
        # Commented for enhanced performance
        #Write-Debug -Message "Job list name: $($JobListName)"

        $CamlQuery = "@
        <View Scope='RecursiveAll'>
            <ViewFields>
                <FieldRef Name='ID'/>
            </ViewFields>
            <Query>
                <Where>
                   <Eq>
                        <FieldRef Name='JobId' />
                        <Value Type='Text'>$($JobId)</Value>
                    </Eq>
                </Where>
            </Query>
            <RowLimit>1</RowLimit> 
        </View>"

        $JobListItem = $null
        If ($null -ne $SharePointConnection) {
            $JobListItem = (Get-PnPListItem -List $JobListName -Query $CamlQuery -Connection $SharePointConnection -ErrorVariable GetPnPListItemError -Verbose:$false | Select-Object -First 1).FieldValues
        }
        Else {
            $JobListItem = (Get-PnPListItem -List $JobListName -Query $CamlQuery -ErrorVariable GetPnPListItemError -Verbose:$false | Select-Object -First 1).FieldValues
        }

        If ($false -eq [string]::IsNullOrEmpty($GetPnPListItemError)) {
            throw "Could not fetch last delta check. Reason: $($GetPnPListItemError)."
        }

        If ($null -eq $JobListItem -or $null -eq $JobListItem.ID) {
            throw "Could not fetch the job list item id from SharePoint."
        }

        # Commented for enhanced performance
        #Write-Debug -Message "Job list item: $(($JobListItem | ConvertTo-Json -Depth 10 | Out-String))"

        return $JobListItem.ID
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to fetch the job list item id from SharePoint. Message: $($_.Exception.Message)"

        Finish-Up -SharePointConnection $SharePointConnection -TerminateWithError $true
    }
}

Function Update-JobListItem {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $JobListName,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [int] $JobListItemId,
        [Parameter(Mandatory = $false)] [AllowNull()] [PnP.PowerShell.Commands.Base.PnPConnection] $SharePointConnection
    )

    Try {
        Write-Verbose -Message "Storing updated job status in SharePoint list '$($JobListName)'..."
        # Commented for enhanced performance
        # Write-Debug -Message "JobListItemId: $($JobListItemId)"
        # Write-Debug -Message "JobListName: $($JobListName)"

        $ListItemValues = @{ 
            "State"        = "$($GLOBAL:output.State)";
            "LogFolderUri" = $GLOBAL:output.LogFolderUri ? "$($GLOBAL:output.LogFolderUri), Log folder uri" : $null;
            "Error"        = "$($GLOBAL:output.Error)";
            "Message"      = "$($GLOBAL:output.ErrorMessage)";
        }

        # Commented for enhanced performance
        #Write-Debug -Message "JobListValues: $($ListItemValues | Out-String)"

        $JobListItem = $null
        If ($null -ne $SharePointConnection) {
            $JobListItem = Set-PnPListItem -List $JobListName -Identity $JobListItemId -Values $ListItemValues -Connection $SharePointConnection -ErrorVariable JobListItemError -Verbose:$false
        }
        Else {
            $JobListItem = Set-PnPListItem -List $JobListName -Identity $JobListItemId -Values $ListItemValues -ErrorVariable JobListItemError -Verbose:$false
        }

        If ($false -eq [string]::IsNullOrEmpty($JobListItemError)) {
            throw throw "Could not store the updated job info. Reason: $($JobListItemError)"
        }

        If ($null -eq $JobListItem) {
            throw "Could not store the updated job info."
        }

        # Commented for enhanced performance
        #Write-Debug -Message "Updated job list item: $($JobListItem | Out-String)"
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to store the updated job info. Message: $($_.Exception.Message)"

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
        # Do nothing
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
        Write-Verbose -Message "[DRY RUN MODE] - The fetched sync job status will not be updated in the job info SharePoint list."
    }

    Foreach ($Module in $GLOBAL:requiredPSModules) {
        Import-PSModule -ModuleName $Module
    }

    # Extract the tenant domain from the database site url
    $Domain = ($DatabaseSiteUrl.Split(".")[0].Split("/")[-1] -replace '/', '')
    # Commented for enhanced performance
    #Write-Debug -Message "Domain: $($Domain)"
    
    # Connect to SharePoint admin to fetch the sync job status
    $SharePointConnection = ConnectTo-SharePointOnline -SiteUrl "https://$($Domain)-admin.sharepoint.com" -AuthenticationMethod $AuthenticationMethod -TenantId $TenantId -UserName $UserName -Pass ($Pass ? (ConvertTo-SecureString -AsPlainText $Pass -Force) : $null) -ClientId $ClientId -ClientSecret $ClientSecret -CertificateThumbprint $CertificateThumbprint -CertificatePath $CertificatePath -CertificatePass ($CertificatePass ? (ConvertTo-SecureString -AsPlainText $CertificatePass -Force) : $null)
    If ($AuthenticationMethod -eq 'ManagedIdentity') {
        # For some reason, even though the ConnectTo-SharePointOnline function specifically states that
        # if the authentication method is 'ManagedIdentity', it returns $null,
        # somehow it comes out as a System.Object[]. This causes the following actions to fail,
        # so explicity set the connection to $null once again is the auth method is 'ManagedIdentity'
        $SharePointConnection = $null
    }

    $ImportJobStatus = Get-SharePointUserProfilesBulkImportJobStatus -JobId $JobId -SharePointConnection $SharePointConnection
    $GLOBAL:output = [PSCustomObject]@{
        JobId        = $ImportJobStatus.JobId;
        State        = "$($ImportJobStatus.State)";
        SourceUri    = $ImportJobStatus.SourceUri;
        LogFolderUri = $ImportJobStatus.LogFolderUri;
        Error        = "$($ImportJobStatus.Error)";
        ErrorMessage = $ImportJobStatus.ErrorMessage;
    }

    # Dispose the SharePoint admin session
    $SharePointConnection = $null

    # Connect to the database SharePoint site to update the sync job status
    $SharePointConnection = ConnectTo-SharePointOnline -SiteUrl $DatabaseSiteUrl -AuthenticationMethod $AuthenticationMethod -TenantId $TenantId -UserName $UserName -Pass ($Pass ? (ConvertTo-SecureString -AsPlainText $Pass -Force) : $null) -ClientId $ClientId -ClientSecret $ClientSecret -CertificateThumbprint $CertificateThumbprint -CertificatePath $CertificatePath -CertificatePass ($CertificatePass ? (ConvertTo-SecureString -AsPlainText $CertificatePass -Force) : $null)
    If ($AuthenticationMethod -eq 'ManagedIdentity') {
        # For some reason, even though the ConnectTo-SharePointOnline function specifically states that
        # if the authentication method is 'ManagedIdentity', it returns $null,
        # somehow it comes out as a System.Object[]. This causes the following actions to fail,
        # so explicity set the connection to $null once again is the auth method is 'ManagedIdentity'
        $SharePointConnection = $null
    }
    
    $JobListItemId = Get-JobListItemId -JobListName $GLOBAL:jobListName -JobId $JobId -SharePointConnection $SharePointConnection

    If ($GLOBAL:dryRun -ne $true -and $null -ne $JobListItemId) {
        Update-JobListItem -JobListName $GLOBAL:jobListName -JobListItemId $JobListItemId -SharePointConnection $SharePointConnection
    }

    Finish-Up -SharePointConnection $SharePointConnection
}
Catch [Exception] {
    Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while executing the program. Message: $($_.Exception.Message)"

    Finish-Up -TerminateWithError $true
}