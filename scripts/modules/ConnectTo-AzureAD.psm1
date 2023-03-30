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
$ErrorActionPreference = 'Continue'
$ProgressPreference = 'Continue'
$VerbosePreference = 'Continue'
$DebugPreference = 'SilentlyContinue' # 'Continue'

<# ---------------- Global variables ---------------- #>
$GLOBAL:azureRunAsConnectionName = "AzureRunAsConnection"
$GLOBAL:azureRunAsCredentialName = "AzureRunAsCredential"

<# ---------------- Program execution ---------------- #>
Function ConnectTo-AzureAD {
    Param(
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [ValidateSet("CredentialsWithoutMFA", "CredentialsWithMFA", "CertificateThumbprint")] [string] $AuthenticationMethod,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $TenantId,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [mailaddress] $UserName,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $Pass,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $ClientId,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $CertificateThumbprint
    )

    Try {
        If (!$AuthenticationMethod) {
            $AuthenticationMethod = "CredentialsWithMFA"   
        }

        Write-Verbose -Message "Connecting to Azure AD..."
        Switch ($AuthenticationMethod) {
            "CredentialsWithoutMFA" {
                ConnectTo-AzureADWithCredentials -UserName $UserName -Pass $Pass -WithMFA $false
            }
            "CredentialsWithMFA" {
                ConnectTo-AzureADWithCredentials -UserName $UserName -Pass $Pass -WithMFA $true
            }
            "CertificateThumbprint" {
                ConnectTo-AzureADWithCertificateThumbprint -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
            }
        }

        return $true
    } 
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to connect to Azure AD with $($AuthenticationMethod.ToLower()). Message: $($_.Exception.Message)"

        Finish-Up -TerminateWithError $true
    }
}

Export-ModuleMember -Function ConnectTo-AzureAD

<# ---------------- Helper functions ---------------- #>

Function ConnectTo-AzureADWithCredentials {
    Param(
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [mailaddress] $UserName,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $Pass,
        [Parameter(Mandatory = $true)] [boolean] $WithMFA
    )
    
    Try {
        Write-Debug -Message "WithMFA: $($WithMFA)"

        If ($WithMFA -eq $true) {
            Write-Verbose -Message "If the authentication dialog doesn't come up, please check the background (known bug in VS Code)."
            Connect-AzureAD -ErrorVariable AzureADConnectionError -Verbose:$false
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
                
                Connect-AzureAD -Credential $creds -ErrorVariable AzureADConnectionError -Verbose:$false
            }
            Else {
                If ($false -eq [string]::IsNullOrEmpty($RunAsCredential.UserName) -and $true -eq [string]::IsNullOrEmpty($RunAsCredential.Password)) {
                    throw "Cannot connect to Azure AD with credentials. Reason: password is not provided."
                }
                Elseif ($true -eq [string]::IsNullOrEmpty($RunAsCredential.UserName) -and $false -eq [string]::IsNullOrEmpty($RunAsCredential.Password)) {
                    throw "Cannot connect to Azure AD with credentials. Reason: username is not provided."
                }
                Else {
                    throw "Cannot connect to Azure AD with credentials. Reason: username and password are not provided."
                }
            }
        }

        If ($false -eq [string]::IsNullOrEmpty($AzureADConnectionError)) {
            throw "Could not connect to Azure AD. Reason: $($AzureADConnectionError)."
        }
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to connect to Azure AD with credentials. Message: $($_.Exception.Message)"

        Finish-Up -TerminateWithError $true
    }
}

Function ConnectTo-AzureADWithCertificateThumbprint {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [ValidatePattern('^[a-f\d]{4}(?:[a-f\d]{4}-){4}[a-f\d]{12}$')] [string] $TenantId,
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

            Connect-AzureAD -TenantId $RunAsConnection.TenantId -ApplicationId $RunAsConnection.ApplicationId -CertificateThumbprint $RunAsConnection.CertificateThumbprint -ErrorVariable AzureADConnectionError -Verbose:$false
        }
        Else {
            throw "Cannot connect to Azure AD with certificate thumbprint. Reason: one of the following parameters is not provided: ClientId, CertificateThumbprint, TenantId."
        }

        If ($false -eq [string]::IsNullOrEmpty($AzureADConnectionError)) {
            throw "Could not connect to Azure AD. Reason: $($AzureADConnectionError)."
        }
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to connect to Azure AD with certificatethumbprint. Message: $($_.Exception.Message)"

        Finish-Up -TerminateWithError $true
    }
}

Function Finish-Up {
    Param(
        [Parameter(Mandatory = $false)] [boolean] $TerminateWithError = $false
    )

    Try {
        Write-Verbose -Message "Disconnecting the session..."
        Disconnect-AzureAD -Verbose:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
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
