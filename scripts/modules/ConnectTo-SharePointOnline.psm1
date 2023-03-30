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
        Either ManagedIdentity, CredentialsWithoutMFA, CredentialsWithMFA (default), ClientSecret, CertificateThumbprint or CertificatePath.
    
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
$ErrorActionPreference = 'Continue'
$ProgressPreference = 'SilentlyContinue'
$VerbosePreference = 'Continue'
$DebugPreference = 'SilentlyContinue' # 'Continue'

<# ---------------- Global variables ---------------- #>
$GLOBAL:azureRunAsConnectionName = "AzureRunAsConnection"
$GLOBAL:azureRunAsCredentialName = "AzureRunAsCredential"

<# ---------------- Program execution ---------------- #>
Function ConnectTo-SharePointOnline {
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
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $CertificatePass,
        [Parameter(Mandatory = $false)] [boolean] $ReturnConnection
    )

    If ($true -eq [string]::IsNullOrEmpty($AuthenticationMethod)) {
        $AuthenticationMethod = "CredentialsWithMFA"
    }
    If (!$ReturnConnection) {
        $ReturnConnection = $false
    }

    Try {
        Write-Verbose -Message "Connecting to SharePoint..."
        Write-Debug -Message "AuthenticationMethod: $($AuthenticationMethod)"

        $SharePointConnection = $null
        Switch ($AuthenticationMethod) {
            "ManagedIdentity" {
                ConnectTo-SharePointOnlineWithManagedIdentity -SiteUrl $SiteUrl
            }
            "CredentialsWithoutMFA" {
                $SharePointConnection = ConnectTo-SharePointOnlineWithCredentials -SiteUrl $SiteUrl -UserName $UserName -Pass $Pass -WithMFA $false -ReturnConnection $ReturnConnection
            }
            "CredentialsWithMFA" {
                $SharePointConnection = ConnectTo-SharePointOnlineWithCredentials -SiteUrl $SiteUrl -UserName $UserName -Pass $Pass -WithMFA $true -ReturnConnection $ReturnConnection
            }
            "ClientSecret" {
                $SharePointConnection = ConnectTo-SharePointOnlineWithClientSecret -SiteUrl $SiteUrl -ClientId $ClientId -ClientSecret $ClientSecret -ReturnConnection $ReturnConnection
            }
            "CertificateThumbprint" {
                $SharePointConnection = ConnectTo-SharePointOnlineWithCertificateThumbprint -SiteUrl $SiteUrl -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -ReturnConnection $ReturnConnection
            }
            "CertificatePath" {
                $SharePointConnection = ConnectTo-SharePointOnlineWithCertificatePath -SiteUrl $SiteUrl -TenantId $TenantId -ClientId $ClientId -CertificatePath $CertificatePath -CertificatePass $CertificatePass -ReturnConnection $ReturnConnection
            }
        }

        If ($AuthenticationMethod -ne "ManagedIdentity" -and $ReturnConnection -eq $true -and $null -eq $SharePointConnection) {
            throw "Could not connect to SharePoint Online."
        }

        return $SharePointConnection
    } 
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to connect to SharePoint Online site '$($SiteUrl)' with $($AuthenticationMethod.ToLower()). Message: $($_.Exception.Message)"

        Finish-Up -TerminateWithError $true
    }
}

Export-ModuleMember -Function ConnectTo-SharePointOnline

<# ---------------- Helper functions ---------------- #>

Function ConnectTo-SharePointOnlineWithManagedIdentity {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [ValidatePattern('(http[s]?|[s]?ftp[s]?)(:\/\/)([^\s,]+)')] [string] $SiteUrl
    )
    
    Try {
        Write-Debug -Message "SiteUrl: $($SiteUrl)"
        Connect-PnPOnline -Url $SiteUrl -ManagedIdentity -ErrorVariable SharePointConnectionError -Verbose:$false | Out-Null

        If ($false -eq [string]::IsNullOrEmpty($SharePointConnectionError)) {
            throw "Could not connect to SharePoint Online. Reason: $($SharePointConnectionError)."
        }
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
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $Pass,
        [Parameter(Mandatory = $true)] [boolean] $WithMFA,
        [Parameter(Mandatory = $false)] [boolean] $ReturnConnection
    )
    
    Try {
        Write-Debug -Message "SiteUrl: $($SiteUrl)"
        Write-Debug -Message "WithMFA: $($WithMFA)"

        $SharePointConnection = $null

        If ($WithMFA -eq $true) {
            Write-Verbose -Message "If the authentication dialog doesn't come up, please check the background (known bug in VS Code)."
            If ($ReturnConnection -eq $true) {
                $SharePointConnection = Connect-PnPOnline -Url $SiteUrl -Interactive -ReturnConnection -ErrorVariable SharePointConnectionError -Verbose:$false
            }
            Else {
                $SharePointConnection = Connect-PnPOnline -Url $SiteUrl -Interactive -ErrorVariable SharePointConnectionError -Verbose:$false
            }
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
                    "Password" = $(ConvertTo-SecureString $Pass -AsPlainText -Force)
                }
            }

            If ($null -ne $RunAsCredential -and $false -eq [string]::IsNullOrEmpty($RunAsCredential.UserName) -and $false -eq [string]::IsNullOrEmpty($RunAsCredential.Password)) {
                Write-Debug -Message "RunAsCredential: $(@{
                    "UserName" = $RunAsCredential.UserName
                    "Password" = ([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($RunAsCredential.Password)))
                } | ConvertTo-Json)"
                $creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $RunAsCredential.UserName, $RunAsCredential.Password
                
                If ($ReturnConnection -eq $true) {
                    $SharePointConnection = Connect-PnPOnline -Url $SiteUrl -Credential $creds -ReturnConnection -ErrorVariable SharePointConnectionError -Verbose:$false
                }
                Else {
                    $SharePointConnection = Connect-PnPOnline -Url $SiteUrl -Credential $creds -ErrorVariable SharePointConnectionError -Verbose:$false
                }
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
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $ClientSecret,
        [Parameter(Mandatory = $false)] [boolean] $ReturnConnection
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
            Write-Debug -Message "SiteUrl: $($SiteUrl)"
            Write-Debug -Message "RunAsCredential: $(@{
                "UserName" = $RunAsCredential.UserName
                "Password" = ([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($RunAsCredential.Password)))
                } | ConvertTo-Json)"
            
            If ($ReturnConnection -eq $true) {
                $SharePointConnection = Connect-PnPOnline -Url $SiteUrl -ClientId $RunAsCredential.UserName -ClientSecret ([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($RunAsCredential.Password))) -ReturnConnection -ErrorVariable SharePointConnectionError -WarningAction Ignore -Verbose:$false
            }
            Else {
                $SharePointConnection = Connect-PnPOnline -Url $SiteUrl -ClientId $RunAsCredential.UserName -ClientSecret ([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($RunAsCredential.Password))) -ErrorVariable SharePointConnectionError -WarningAction Ignore -Verbose:$false
            }
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
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $CertificateThumbprint,
        [Parameter(Mandatory = $false)] [boolean] $ReturnConnection
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
            Write-Debug -Message "SiteUrl: $($SiteUrl)"
            Write-Debug -Message "RunAsConnection: $($RunAsConnection | ConvertTo-Json)"

            If ($ReturnConnection -eq $true) {
                $SharePointConnection = Connect-PnPOnline -Url $SiteUrl -ClientId $RunAsConnection.ApplicationId -Tenant $RunAsConnection.TenantId -Thumbprint $RunAsConnection.CertificateThumbprint -ReturnConnection -ErrorVariable SharePointConnectionError -Verbose:$false
            }
            Else {
                $SharePointConnection = Connect-PnPOnline -Url $SiteUrl -ClientId $RunAsConnection.ApplicationId -Tenant $RunAsConnection.TenantId -Thumbprint $RunAsConnection.CertificateThumbprint -ErrorVariable SharePointConnectionError -Verbose:$false
            }
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
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $CertificatePass,
        [Parameter(Mandatory = $false)] [boolean] $ReturnConnection
    )

    Try {
        $SharePointConnection = $null
        
        If ($false -eq [string]::IsNullOrEmpty($ClientId) -and $false -eq [string]::IsNullOrEmpty($CertificatePass) -and $false -eq [string]::IsNullOrEmpty($CertificatePath) -and $false -eq [string]::IsNullOrEmpty($TenantId)) {
            Write-Debug -Message "SiteUrl: $($SiteUrl)"
            Write-Debug -Message "TenantId: $($TenantId)"
            Write-Debug -Message "ClientId: $($ClientId)"
            Write-Debug -Message "CertificatePath: $($CertificatePath)"
            Write-Debug -Message "CertificatePass: $($CertificatePass)"
            
            If ($CertificatePath | Test-Path) {
                If ($ReturnConnection -eq $true) {
                    $SharePointConnection = Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -CertificatePath $CertificatePath -CertificatePassword (ConvertTo-SecureString $CertificatePass -AsPlainText -Force) -Tenant $TenantId -ReturnConnection -ErrorVariable SharePointConnectionError -Verbose:$false
                }
                Else {
                    $SharePointConnection = Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -CertificatePath $CertificatePath -CertificatePassword (ConvertTo-SecureString $CertificatePass -AsPlainText -Force) -Tenant $TenantId -ErrorVariable SharePointConnectionError -Verbose:$false
                }
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
