<#
    .SYNOPSIS
        Module with helper functions to verify the authorization of an Azure connection.

    .DESCRIPTION
        This module allows you to verify the authorization of an Azure connection. If the connection
        is not authorized, the module will try to authorize the connection by presenting the user with
        an authorization dialog.
    
    .FUNCTIONALITY
        - Verify-ConnectionAuthorization

    .PARAMETER ResourceGroupName <string> [required]
        The resource group name in which the connection to verify resides.
    
    .PARAMETER ConnectionName <string> [required]
        The name of the connection to verify.

    .OUTPUTS
        None

    .LINK
        Inspired by:
            - https://github.com/logicappsio/LogicAppConnectionAuth/blob/master/LogicAppConnectionAuth.ps1
        
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

Function Verify-ConnectionAuthorization {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $ResourceGroupName,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $ConnectionName
    )

    Try {
        Initialize

        # Prevent get_SerializationSettings error by enabling AzureRm prefix aliases for Az modules.
        # More information here: https://azurelessons.com/method-get_serializationsettings-error/
        Enable-AzureRmAlias -Scope CurrentUser
       
        $Connection = Get-ConnectionStatus -ResourceGroupName $ResourceGroupName -ConnectionName $ConnectionName

        If ($Connection.Properties.Statuses[0].status -ne "Connected") {
            Authorize-Connection -Connection $Connection

            $Connection = Get-ConnectionStatus -ResourceGroupName $ResourceGroupName -ConnectionName $ConnectionName

            If ($Connection.Properties.Statuses[0].status -ne "Connected") {
                throw "Could not verify connection. Status: $($Connection.Properties.Statuses[0].status)."
            }

            Write-Verbose -Message "Connection is authorized!"
        }
        Else {
            Write-Verbose -Message "Connection is already authorized."
        }
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while verifying the connection '$($ConnectionName)'. Message: $($_.Exception.Message)"

        Write-Verbose -Message "Terminating program. Reason: could not verify the connection."
        Finish-Up -TerminateWithError $true
    }
}

Export-ModuleMember -Function Verify-ConnectionAuthorization

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

Function Get-ConnectionStatus {
    Param(
        [parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $ResourceGroupName,
        [parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $ConnectionName
    )

    Try {
        Write-Verbose -Message "Fetching connection '$($ConnectionName)'..."
        $Connection = Get-AzureRmResource -ResourceType "Microsoft.Web/connections" -ResourceGroupName $ResourceGroupName -ResourceName $ConnectionName -ErrorAction SilentlyContinue -Verbose:$false

        If ($null -eq $Connection) {
            throw "Could not fetch connection '$($ConnectionName)'."
        }

        Write-Debug -Message "Connection: $($Connection)"

        return $Connection
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while fetching the connection status of connection '$($ConnectionName)'. Message: $($_.Exception.Message)"

        Write-Verbose -Message "Terminating program. Reason: could not fetch the connection status."
        Finish-Up -TerminateWithError $true
    }
}

Function Show-OAuthWindow {
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $AuthorizationUrl
    )

    Try {
        Write-Verbose -Message "Assembling authorization dialog..."
        Add-Type -AssemblyName System.Windows.Forms
 
        $Form = New-Object -TypeName System.Windows.Forms.Form -Property @{Width = 600; Height = 800 }
        $Web = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{Width = 580; Height = 780; Url = ($AuthorizationUrl -f ($Scope -join "%20")) }
        $DocComp = {
            $Global:uri = $Web.Url.AbsoluteUri
            if ($Global:Uri -match "error=[^&]*|code=[^&]*") { 
                $Form.Close() 
            }
        }
        $Web.ScriptErrorsSuppressed = $true
        $Web.Add_DocumentCompleted($DocComp)
        $Form.Controls.Add($Web)
        $Form.Add_Shown({ $Form.Activate() })

        Write-Verbose -Message "Assembled! Pending user authorization input..."
        $Form.ShowDialog() | Out-Null
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while assembling the authorization dialog. Message: $($_.Exception.Message)"

        Write-Verbose -Message "Terminating program. Reason: could not assemble the authorization dialog."
        Finish-Up -TerminateWithError $true
    }
}

Function Authorize-Connection {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [object] $Connection
    )

    Try {
        Write-Verbose -Message "Connection is not authorized yet. Current status: $($Connection.Properties.Statuses[0].status). Attempting connection authorization..."
        
        Write-Verbose -Message "Fetching connection authorization link..."

        $Parameters = @{
            "parameters" = , @{
                "parameterName" = "token";
                "redirectUrl"   = "https://ema1.exp.azure.com/ema/default/authredirect"
            }
        }

        $Response = Invoke-AzureRmResourceAction -Action "listConsentLinks" -ResourceId $Connection.ResourceId -Parameters $Parameters -Force -ErrorAction SilentlyContinue -Verbose:$false
        
        If ($null -eq $Response -or $null -eq $Response.Value.Link) {
            throw "Could not fetch connection authorization link."
        }

        Write-Debug -Message "Authorization action: $($Response)"
        
        $AuthorizationUrl = $Response.Value.Link
        Show-OAuthWindow -AuthorizationUrl $AuthorizationUrl

        $Code = ($uri | Select-string -pattern '(code=)(.*)$').Matches[0].Groups[2].Value
        If ($true -eq [string]::IsNullOrEmpty($Code)) {
            throw "Could not fetch access code from user authorization."
        }

        $Parameters = @{ }
        $Parameters.Add("code", $Code)

        Write-Debug -Message "Access code: $($Code)"
        
        Write-Verbose "Authorizing connection with access code..."     
        Invoke-AzureRmResourceAction -Action "confirmConsentCode" -ResourceId $Connection.ResourceId -Parameters $Parameters -Force -ErrorAction SilentlyContinue -ErrorVariable InvokeAzureRmResourceActionError -Verbose:$false | Out-Null

        # There is a known bug in PowerShell which throws a meaningless error about an argument 'obj' being null and not being able to process it.
        # This error does not mean the action failed. Therefore, it should not be declared as a breaking error
        # More information here: https://social.technet.microsoft.com/Forums/windowsserver/en-US/f0edd01c-7711-46cd-b47b-45785d44f062/why-is-the-value-of-argument-quotobjquot-null
        If ('Cannot process argument because the value of argument "obj" is null. Change the value of argument "obj" to a non-null value.' -ne $InvokeAzureRmResourceActionError) {
            throw "Could not authorize connection."
        }
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while authorizing the connection. Message: $($_.Exception.Message)"

        Write-Verbose -Message "Terminating program. Reason: could not authorize the connection."
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