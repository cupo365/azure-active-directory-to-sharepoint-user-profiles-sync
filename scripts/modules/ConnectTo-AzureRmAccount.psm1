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
Function ConnectTo-AzureRmAccount {
    Param(
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [ValidateSet("ManagedIdentity", "CredentialsWithoutMFA", "CredentialsWithMFA", "ClientSecret", "CertificateThumbprint")] [string] $AuthenticationMethod,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $TenantId,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [mailaddress] $UserName,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $Pass,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $ClientId,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $ClientSecret,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $CertificateThumbprint
    )

    If ($true -eq [string]::IsNullOrEmpty($AuthenticationMethod)) {
        $AuthenticationMethod = "CredentialsWithMFA"
    }

    Try {
        Initialize

        Write-Warning -Message "[CAUTION] - Azure Remote has been deprecated and replaced by the Az module. Beware that Azure Remote will retire as of 29 February 2024! More information here: https://learn.microsoft.com/en-us/powershell/module/azurerm.profile/connect-azurermaccount"
        # Prevent get_SerializationSettings error by enabling AzureRm prefix aliases for Az modules.
        # More information here: https://azurelessons.com/method-get_serializationsettings-error/
        Enable-AzureRmAlias -Scope CurrentUser -Verbose:$false | Out-Null

        Write-Verbose -Message "Connecting to Azure Remote..."
        Write-Debug -Message "AuthenticationMethod: $($AuthenticationMethod)"

        Switch ($AuthenticationMethod) {
            "ManagedIdentity" {
                ConnectTo-AzureRmAccountWithManagedIdentity
            }
            "CredentialsWithoutMFA" {
                ConnectTo-AzureRmAccountWithCredentials -UserName $UserName -Pass $Pass -WithMFA $false
            }
            "CredentialsWithMFA" {
                ConnectTo-AzureRmAccountWithCredentials -UserName $UserName -Pass $Pass -WithMFA $true
            }
            "ClientSecret" {
                ConnectTo-AzureRmAccountWithClientSecret -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
            }
            "CertificateThumbprint" {
                ConnectTo-AzureRmAccountWithCertificateThumbprint -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
            }
        }
    } 
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to connect to Azure with $($AuthenticationMethod.ToLower()). Message: $($_.Exception.Message)"

        Finish-Up -TerminateWithError $true
    }
}

Export-ModuleMember -Function ConnectTo-AzureRmAccount

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

Function ConnectTo-AzureRmAccountWithManagedIdentity {
    Try {
        Connect-AzureRmAccount -MSI -ErrorVariable AzureConnectionError -Verbose:$false | Out-Null

        If ($false -eq [string]::IsNullOrEmpty($AzureConnectionError)) {
            throw "Could not connect to Azure Remote. Reason: $($AzureConnectionError)."
        }
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to connect to Azure Remote with managedidentity. Message: $($_.Exception.Message)"

        Finish-Up -TerminateWithError $true
    }
}

Function ConnectTo-AzureRmAccountWithCredentials {
    Param(
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [mailaddress] $UserName,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $Pass,
        [Parameter(Mandatory = $true)] [boolean] $WithMFA
    )
    
    Try {
        Write-Debug -Message "WithMFA: $($WithMFA)"

        If ($WithMFA -eq $true) {
            Write-Verbose -Message "If the authentication dialog doesn't come up, please check the background (known bug in VS Code)."
            Connect-AzureRmAccount -ErrorVariable AzureConnectionError -Verbose:$false
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
                
                Connect-AzureRmAccount -Credential $creds -ErrorVariable AzureConnectionError -Verbose:$false
            }
            Else {
                If ($false -eq [string]::IsNullOrEmpty($RunAsCredential.UserName) -and $true -eq [string]::IsNullOrEmpty($RunAsCredential.Password)) {
                    throw "Cannot connect to Azure Remote with credentials. Reason: password is not provided."
                }
                Elseif ($true -eq [string]::IsNullOrEmpty($RunAsCredential.UserName) -and $false -eq [string]::IsNullOrEmpty($RunAsCredential.Password)) {
                    throw "Cannot connect to Azure Remote with credentials. Reason: username is not provided."
                }
                Else {
                    throw "Cannot connect to Azure Remote with credentials. Reason: username and password are not provided."
                }
            }
        }

        If ($false -eq [string]::IsNullOrEmpty($AzureConnectionError)) {
            throw "Could not connect to Azure Remote. Reason: $($AzureConnectionError)."
        }
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to connect to Azure Remote with credentials. Message: $($_.Exception.Message)"

        Finish-Up -TerminateWithError $true
    }
}

Function ConnectTo-AzureRmAccountWithClientSecret {
    Param(
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $TenantId,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $ClientId,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $ClientSecret
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
            Connect-AzureRmAccount -ServicePrincipal -Tenant $TenantId -Credential $creds -ErrorVariable AzureConnectionError -Verbose:$false
        }
        Else {
            throw "Cannot connect to Azure Remote with certificate thumbprint. Reason: one of the following parameters is not provided: ClientId, ClientSecret, TenantId."
        }

        If ($false -eq [string]::IsNullOrEmpty($AzureConnectionError)) {
            throw "Could not connect to Azure Remote. Reason: $($AzureConnectionError)."
        }
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to connect to Azure Remote with clientsecret. Message: $($_.Exception.Message)"

        Finish-Up -TerminateWithError $true
    }
}

Function ConnectTo-AzureRmAccountWithCertificateThumbprint {
    Param(
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $TenantId,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $ClientId,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $CertificateThumbprint
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

            Connect-AzureRmAccount -ServicePrincipal -ApplicationId $RunAsConnection.ApplicationId -Tenant $RunAsConnection.TenantId -CertificateThumbprint $RunAsConnection.CertificateThumbprint -ErrorVariable AzureConnectionError -Verbose:$false
        }
        Else {
            throw "Cannot connect to Azure Remote with certificate thumbprint. Reason: one of the following parameters is not provided: ClientId, CertificateThumbprint, TenantId."
        }

        If ($false -eq [string]::IsNullOrEmpty($AzureConnectionError)) {
            throw "Could not connect to Azure Remote. Reason: $($AzureConnectionError)."
        }
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to connect to Azure Remote with certificatethumbprint. Message: $($_.Exception.Message)"

        Finish-Up -TerminateWithError $true
    }
}

Function Finish-Up {
    Param(
        [Parameter(Mandatory = $false)] [boolean] $TerminateWithError = $false
    )

    Try {
        Write-Verbose -Message "Disconnecting the session..."       
        Disconnect-AzureRmAccount -Verbose:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
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
