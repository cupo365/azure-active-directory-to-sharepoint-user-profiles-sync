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
$ErrorActionPreference = 'Continue'
$ProgressPreference = 'Continue'
$VerbosePreference = "Continue"
$DebugPreference = 'Continue' # 'Continue'

<# ---------------- Global variables ---------------- #>
$GLOBAL:requiredPSModules = "PnP.PowerShell"

<# ---------------- Program execution ---------------- #>

Function Get-PnPTemplate {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [ValidatePattern('(http[s]?|[s]?ftp[s]?)(:\/\/)([^\s,]+)')] [string] $SiteUrl,
        [Parameter(Mandatory = $true)] [ValidateSet("pnp", "xml")] [string] $TemplateType,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $Workspace,
        [Parameter(Mandatory = $false)] [AllowNull()] [array] $Handlers,
        [Parameter(Mandatory = $false)] [AllowNull()] [array] $ListsToExtract,
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
    If ($true -eq [string]::IsNullOrEmpty($TemplateType)) {
        $TemplateType = "xml"
    }
    
    Try {
        Initialize

        $SharePointConnection = ConnectTo-SharePointOnline -SiteUrl $SiteUrl -AuthenticationMethod $AuthenticationMethod -TenantId $TenantId -UserName $UserName -Pass $Pass -ClientId $ClientId -ClientSecret $ClientSecret -CertificateThumbprint $CertificateThumbprint -CertificatePath $CertificatePath -CertificatePass $CertificatePass

        Get-Template -SiteUrl $SiteUrl -TemplateType $TemplateType -Workspace $Workspace -Handlers $Handlers -ListsToExtract $ListsToExtract -SharePointConnection $SharePointConnection
        
        Finish-Up -SharePointConnection $SharePointConnection
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while executing the program. Message: $($_.Exception.Message)"

        Finish-Up -SharePointConnection $SharePointConnection -TerminateWithError $true
    }
}

Export-ModuleMember -Function Get-PnPTemplate

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
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while initialize. Message: $($_.Exception.Message)"

        Finish-Up -TerminateWithError $true
    }
}

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
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $CertificatePass
    )

    $SharePointConnection = $null

    Try {
        Remove-Module -Name ConnectTo-SharePointOnline -WarningAction SilentlyContinue -ErrorAction SilentlyContinue -Verbose:$false -Force
        Import-Module -Name ".\Modules\ConnectTo-SharePointOnline.psm1" -Scope Local -Force -ErrorAction Stop -WarningAction SilentlyContinue -Verbose:$false | Out-Null

        $SharePointConnection = ConnectTo-SharePointOnline\ConnectTo-SharePointOnline -SiteUrl $SiteUrl -AuthenticationMethod $AuthenticationMethod -TenantId $TenantId -UserName $UserName -Pass $Pass -ClientId $ClientId -ClientSecret $ClientSecret -CertificateThumbprint $CertificateThumbprint -CertificatePath $CertificatePath -CertificatePass $CertificatePass -ReturnConnection $true -Verbose

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

Function Get-Template {
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $SiteUrl,
        [Parameter(Mandatory = $true)] [ValidateSet("pnp", "xml")] [string] $TemplateType,
        [Parameter(Mandatory = $false)] [AllowNull()] [AllowEmptyString()] [string] $Workspace,
        [Parameter(Mandatory = $false)] [AllowNull()] [array] $Handlers,
        [Parameter(Mandatory = $false)] [AllowNull()] [array] $ListsToExtract,
        [Parameter(Mandatory = $false)] [AllowNull()] [PnP.PowerShell.Commands.Base.PnPConnection] $SharePointConnection
    )

    Try {
        Write-Verbose -Message "Fetching provisioning template..."

        $TemplateFileDestinationPath = $null
        If ($null -ne $Workspace) {
            If ($Workspace | Test-Path) {
                $TemplateFileDestinationPath = Join-Path $Workspace "pnp-site-template.$($TemplateType)"
            }
            Else {
                throw "The specified workspace does not exist or is invalid."
            }
        }
        Else {
            $TemplateFileDestinationPath = "./pnp-site-template.$($TemplateType)"
        }

        If ($null -ne $Handlers -and $Handlers.Count -gt 0) {
            If ($null -ne $ListsToExtract -and $ListsToExtract.Count -gt 0) {
                If ($Handlers.Contains("Lists") -eq $false) {
                    $Handlers += "Lists"
                }

                If ($null -ne $SharePointConnection) {
                    Get-PnPSiteTemplate -Out $TemplateFileDestinationPath -Handlers $Handlers -ListsToExtract $ListsToExtract -Connection $SharePointConnection -ErrorVariable GetPnPSiteTemplateError -Verbose:$false
                }
                Else {
                    Get-PnPSiteTemplate -Out $TemplateFileDestinationPath -Handlers $Handlers -ListsToExtract $ListsToExtract -ErrorVariable GetPnPSiteTemplateError -Verbose:$false
                }
            }
            Else {
                If ($null -ne $SharePointConnection) {
                    Get-PnPSiteTemplate -Out $TemplateFileDestinationPath -Handlers $Handlers -Connection $SharePointConnection -ErrorVariable GetPnPSiteTemplateError -Verbose:$false
                }
                Else {
                    Get-PnPSiteTemplate -Out $TemplateFileDestinationPath -Handlers $Handlers -ErrorVariable GetPnPSiteTemplateError -Verbose:$false
                }
            }
        }
        Else {
            If ($null -ne $ListsToExtract -and $ListsToExtract.Count -gt 0) {
                If ($null -ne $SharePointConnection) {
                    Get-PnPSiteTemplate -Out $TemplateFileDestinationPath -Handlers Lists -ListsToExtract $ListsToExtract -Connection $SharePointConnection -ErrorVariable GetPnPSiteTemplateError -Verbose:$false
                } 
                Else {
                    Get-PnPSiteTemplate -Out $TemplateFileDestinationPath -Handlers Lists -ListsToExtract $ListsToExtract -ErrorVariable GetPnPSiteTemplateError -Verbose:$false
                }
            }
            Else {
                If ($null -ne $SharePointConnection) {
                    Get-PnPSiteTemplate -Out $TemplateFileDestinationPath -Connection $SharePointConnection -ErrorVariable GetPnPSiteTemplateError -Verbose:$false
                }
                Else {
                    Get-PnPSiteTemplate -Out $TemplateFileDestinationPath -ErrorVariable GetPnPSiteTemplateError -Verbose:$false
                }
            }
        }

        If ($false -eq [string]::IsNullOrEmpty($GetPnPSiteTemplateError)) {
            throw "Could not fetch PnP site template. Reason: $($GetPnPSiteTemplateError)."
        }

        If ($TemplateType -eq "xml") {
            Write-Verbose -Message "Attempting to open template file in VS Code..."
            Try {
                code $TemplateFileDestinationPath
            }
            Catch [Exception] {
                # Do nothing
            }
        }
    }
    Catch [Exception] {
        Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to fetch and save the $($TemplateType.ToUpper()) provisioning template from site '$($SiteUrl)' to workspace '$($Workspace)'. Message: $($_.Exception.Message)"

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