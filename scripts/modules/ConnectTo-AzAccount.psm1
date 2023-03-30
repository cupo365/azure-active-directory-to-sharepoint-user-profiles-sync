<#
    .SYNOPSIS
        Module with helper functions to connect to Azure.

    .DESCRIPTION
        This module allows you to connect to Azure using various authentication methods, like:
            - ManagedIdentity
            - CredentialsWithoutMFA
            - CredentialsWithMFA
            - ClientSecret
            - CertificateThumbprint
            - CertificatePath
    
    .FUNCTIONALITY
        - ConnectTo-AzAccount

    .PARAMETER AuthenticationMethod <string> [optional]
        The authentication method to use to connect to Azure. 
        Either ManagedIdentity, CredentialsWithoutMFA, CredentialsWithMFA, ClientSecret, CertificateThumbprint or CertificatePath.
    
    .PARAMETER TenantId <string> [optional]
        The tenant ID of the Azure Active Directory tenant that is used to connect to Azure.
        Only required when using the authentication methods ClientSecret, CertificateThumbprint or CertificatePath.
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
        Only required when using the authentication methods ClientSecret, CertificateThumbprint or CertificatePath.
        Example: 00000000-0000-0000-0000-000000000000
    
    .PARAMETER ClientSecret <string> [optional]
        The client secret of the Azure AD app registration that is used to connect to Azure.
        Only required when using the authentication method ClientSecret.
    
    .PARAMETER CertificateThumbprint <string> [optional]
        The thumbprint of the certificate that is used to connect to Azure.
        Only required when using the authentication method CertificateThumbprint.
        Example: 0000000000000000000000000000000000000000
    
    .PARAMETER CertificatePath <string> [optional]
        The path to the certificate that is used to connect to Azure.
        Only required when using the authentication method CertificatePath.
        Example: C:\certificates\contoso.pfx
    
    .PARAMETER CertificatePass <string> [optional]
        The password of the certificate that is used to connect to Azure.
        Only required when using the authentication method CertificatePath.

    .OUTPUTS
        None

    .LINK
        Inspired by:
            - https://learn.microsoft.com/en-us/powershell/module/az.accounts/connect-azaccount
        
        Developers:
            - Lennart de Waart
                - GitHub: https://github.com/cupo365
                - LinkedIn: https://www.linkedin.com/in/lennart-dewaart/
#>

<# --------------- Runtime variables ---------------- #>
$ErrorActionPreference = 'Continue'
$ProgressPreference = 'Continue'
$VerbosePreference = 'Continue'
$DebugPreference = 'SilentlyContinue' # 'Continue'

<# ---------------- Global variables ---------------- #>
$GLOBAL:azureRunAsConnectionName = "AzureRunAsConnection"
$GLOBAL:azureRunAsCredentialName = "AzureRunAsCredential"
$GLOBAL:requiredPSVersion = "5.1"
$GLOBAL:psCommandPath = $MyInvocation.PSCommandPath

<# ---------------- Program execution ---------------- #>
Function ConnectTo-AzAccount {
    Param(
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
    
    If ($true -eq [string]::IsNullOrEmpty($AuthenticationMethod)) {
        $AuthenticationMethod = "CredentialsWithMFA"
    }

    Try {
        Initialize
        
        Write-Verbose -Message "Connecting to Azure..."
        Write-Debug -Message "AuthenticationMethod: $($AuthenticationMethod)"

        Switch ($AuthenticationMethod) {
            "ManagedIdentity" {
                ConnectTo-AzAccountWithManagedIdentity
            }
            "CredentialsWithoutMFA" {
                ConnectTo-AzAccountWithCredentials -UserName $UserName -Pass $Pass -WithMFA $false
            }
            "CredentialsWithMFA" {
                ConnectTo-AzAccountWithCredentials -UserName $UserName -Pass $Pass -WithMFA $true
            }
            "ClientSecret" {
                ConnectTo-AzAccountWithClientSecret -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
            }
            "CertificateThumbprint" {
                ConnectTo-AzAccountWithCertificateThumbprint -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
            }
            "CertificatePath" {
                ConnectTo-AzAccountWithCertificatePath -TenantId $TenantId -ClientId $ClientId -CertificatePath $CertificatePath -CertificatePass $CertificatePass
            }
        }
    } 
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to connect to Azure with $($AuthenticationMethod.ToLower()). Message: $($_.Exception.Message)"

        Finish-Up -TerminateWithError $true
    }
}

Export-ModuleMember -Function ConnectTo-AzAccount

<# ---------------- Helper functions ---------------- #>

Function Initialize {
    Try {
        Remove-Module -Name Validate-PSVersion -WarningAction SilentlyContinue -ErrorAction SilentlyContinue -Verbose:$false -Force
        Import-Module -Name ".\Modules\Validate-PSVersion.psm1" -Scope Local -Force -ErrorAction Stop -WarningAction SilentlyContinue -Verbose:$false | Out-Null

        Validate-PSVersion -RequiredPSVersion $GLOBAL:requiredPSVersion -CommandPath $GLOBAL:psCommandPath -Verbose
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while initializing. Message: $($_.Exception.Message)"

        Finish-Up -TerminateWithError $true
    }
}

Function ConnectTo-AzAccountWithManagedIdentity {
    Try {
        Connect-AzAccount -Identity -ErrorVariable AzureConnectionError -Verbose:$false | Out-Null

        If ($false -eq [string]::IsNullOrEmpty($AzureConnectionError)) {
            throw "Could not connect to Azure. Reason: $($AzureConnectionError)."
        }
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to connect to Azure with managedidentity. Message: $($_.Exception.Message)"

        Finish-Up -TerminateWithError $true
    }
}

Function ConnectTo-AzAccountWithCredentials {
    Param(
        [Parameter(Mandatory = $true)] [AllowNull()] [AllowEmptyString()] [mailaddress] $UserName,
        [Parameter(Mandatory = $true)] [AllowNull()] [AllowEmptyString()] [string] $Pass,
        [Parameter(Mandatory = $true)] [boolean] $WithMFA
    )
    
    Try {
        Write-Debug -Message "WithMFA: $($WithMFA)"

        If ($WithMFA -eq $true) {
            Write-Verbose -Message "If the authentication dialog doesn't come up, please check the background (known bug in VS Code)."
            Connect-AzAccount -ErrorVariable AzureConnectionError -Verbose:$false
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
                
                Connect-AzAccount -Credential $creds -ErrorVariable AzureConnectionError -Verbose:$false
            }
            Else {
                If ($false -eq [string]::IsNullOrEmpty($RunAsCredential.UserName) -and $true -eq [string]::IsNullOrEmpty($RunAsCredential.Password)) {
                    throw "Cannot connect to Azure with credentials. Reason: password is not provided."
                }
                Elseif ($true -eq [string]::IsNullOrEmpty($RunAsCredential.UserName) -and $false -eq [string]::IsNullOrEmpty($RunAsCredential.Password)) {
                    throw "Cannot connect to Azure with credentials. Reason: username is not provided."
                }
                Else {
                    throw "Cannot connect to Azure with credentials. Reason: username and password are not provided."
                }
            }
        }

        If ($false -eq [string]::IsNullOrEmpty($AzureConnectionError)) {
            throw "Could not connect to Azure. Reason: $($AzureConnectionError)."
        }
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to connect to Azure with credentials. Message: $($_.Exception.Message)"

        Finish-Up -TerminateWithError $true
    }
}

Function ConnectTo-AzAccountWithClientSecret {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [ValidatePattern('^[a-f\d]{4}(?:[a-f\d]{4}-){4}[a-f\d]{12}$')] [string] $TenantId,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [ValidatePattern('^[a-f\d]{4}(?:[a-f\d]{4}-){4}[a-f\d]{12}$')] [string] $ClientId,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $ClientSecret
    )
    
    Try {        
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

        If ($null -ne $RunAsCredential -and $false -eq [string]::IsNullOrEmpty($RunAsCredential.UserName) -and $false -eq [string]::IsNullOrEmpty($RunAsCredential.Password) -and $false -eq [string]::IsNullOrEmpty($TenantId)) {
            Write-Debug -Message "TenantId: $($TenantId)"
            Write-Debug -Message "RunAsCredential: $(@{
                "UserName" = $RunAsCredential.UserName
                "Password" = ([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($RunAsCredential.Password)))
            } | ConvertTo-Json)"
            
            $creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $RunAsCredential.UserName, $RunAsCredential.Password
            Connect-AzAccount -ServicePrincipal -TenantId $TenantId -Credential $creds -ErrorVariable AzureConnectionError -Verbose:$false
        }
        Else {
            throw "Cannot connect to Azure with certificate thumbprint. Reason: one of the following parameters is not provided: ClientId, ClientSecret, TenantId."
        }

        If ($false -eq [string]::IsNullOrEmpty($AzureConnectionError)) {
            throw "Could not connect to Azure. Reason: $($AzureConnectionError)."
        }
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to connect to Azure with clientsecret. Message: $($_.Exception.Message)"

        Finish-Up -TerminateWithError $true
    }
}

Function ConnectTo-AzAccountWithCertificateThumbprint {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [ValidatePattern('^[a-f\d]{4}(?:[a-f\d]{4}-){4}[a-f\d]{12}$')] [string] $TenantId,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [ValidatePattern('^[a-f\d]{4}(?:[a-f\d]{4}-){4}[a-f\d]{12}$')] [string] $ClientId,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $CertificateThumbprint
    )
    
    Try {
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
            Write-Debug -Message "RunAsConnection: $($RunAsConnection | ConvertTo-Json)"

            Connect-AzAccount -ServicePrincipal -ApplicationId $RunAsConnection.ApplicationId -Tenant $RunAsConnection.TenantId -CertificateThumbprint $RunAsConnection.CertificateThumbprint -ErrorVariable AzureConnectionError -Verbose:$false
        }
        Else {
            throw "Cannot connect to Azure with certificate thumbprint. Reason: one of the following parameters is not provided: ClientId, CertificateThumbprint, TenantId."
        }

        If ($false -eq [string]::IsNullOrEmpty($AzureConnectionError)) {
            throw "Could not connect to Azure. Reason: $($AzureConnectionError)."
        }
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to connect to Azure with certificatethumbprint. Message: $($_.Exception.Message)"

        Finish-Up -TerminateWithError $true
    }
}

Function ConnectTo-AzAccountWithCertificatePath {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [ValidatePattern('^[a-f\d]{4}(?:[a-f\d]{4}-){4}[a-f\d]{12}$')] [string] $TenantId,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [ValidatePattern('^[a-f\d]{4}(?:[a-f\d]{4}-){4}[a-f\d]{12}$')] [string] $ClientId,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $CertificatePath,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $CertificatePass
    )

    Try {
        If ($false -eq [string]::IsNullOrEmpty($ClientId) -and $false -eq [string]::IsNullOrEmpty($CertificatePass) -and $false -eq [string]::IsNullOrEmpty($CertificatePath) -and $false -eq [string]::IsNullOrEmpty($TenantId)) {
            Write-Debug -Message "SiteUrl: $($SiteUrl)"
            Write-Debug -Message "TenantId: $($TenantId)"
            Write-Debug -Message "ClientId: $($ClientId)"
            Write-Debug -Message "CertificatePath: $($CertificatePath)"
            Write-Debug -Message "CertificatePass: $($CertificatePass)"
            
            If ($CertificatePath | Test-Path) {
                Connect-AzAccount -ServicePrincipal -ApplicationId $ClientId -CertificatePath $CertificatePath -CertificatePassword (ConvertTo-SecureString $CertificatePass -AsPlainText -Force) -TenantId $TenantId -ErrorVariable AzureConnectionError -Verbose:$false
            }
            Else {
                throw "Cannot connect to Azure with certificate path. Reason: the provided certificate path is invalid or non existent."
            }
        }
        Else {
            throw "Cannot connect to Azure with certificate path. Reason: one of the following parameters is not provided: ClientId, CertificatePass, CertificatePath, TenantId."
        }

        If ($false -eq [string]::IsNullOrEmpty($AzureConnectionError)) {
            throw "Could not connect to Azure. Reason: $($AzureConnectionError)."
        }
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to connect to Azure with certificatepath. Message: $($_.Exception.Message)"

        Finish-Up -TerminateWithError $true
    }
}

Function Finish-Up {
    Param(
        [Parameter(Mandatory = $false)] [boolean] $TerminateWithError = $false
    )

    Try {
        Write-Verbose -Message "Disconnecting the session..."       
        Disconnect-AzAccount -Verbose:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
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
