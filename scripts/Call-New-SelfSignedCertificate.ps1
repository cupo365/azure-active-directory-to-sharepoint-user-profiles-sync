<#
    .SYNOPSIS
        Script to create a self-signed certificate.

    .DESCRIPTION
        This script allows you to create a self-signed certificate and its private key information file.
        The certificate can be used to authenticate identities, like Azure app registrations.
        The script also automatically removes the created certificate from the user's keystore.

    .PARAMETER CertificateName <string> [required]
        The name of the certificate.

    .PARAMETER CertificatePwd <string> [required]
        The password used for certificate encryption.

    .PARAMETER MonthsValid <int> [required]
        The amount of months before the certificate becomes invalid.

    .PARAMETER FolderPath <string> [required]
        The (relative) folder path the certificate should be created in.

    .PARAMETER RemoveFromKeyStore <boolean> [optional]
        Whether or not the created certificate should be removed from the user's keystore.
      
    .OUTPUTS
        A .cer and .pfx file in the specified folder path of type
        System.Security.Cryptography.X509Certificates.X509Certificate2.

    .LINK
        Developers:
            - Lennart de Waart
                - GitHub: https://github.com/cupo365
                - LinkedIn: https://www.linkedin.com/in/lennart-dewaart/
#>

<# --------------- Runtime variables ---------------- #>
$ErrorActionPreference = 'Continue'
$ProgressPreference = 'SilentlyContinue'
$VerbosePreference = "Continue"
$DebugPreference = 'SilentlyContinue' # 'Continue'

Remove-Module -Name New-SelfSignedCertificate -WarningAction SilentlyContinue -ErrorAction SilentlyContinue -Verbose:$false -Force
Import-Module -Name (Join-Path (Split-Path -parent $PSCommandPath) "\Modules\New-SelfSignedCertificate.psm1") -Scope Local -Force -ErrorAction Stop -WarningAction SilentlyContinue -Verbose:$false | Out-Null

$CertificateName = 'cer-aad2spup'
$CertificatePwd = ''
$MonthsValid = 36
$FolderPath = '../certificates'
$RemoveFromKeyStore = $false

New-SelfSignedCertificate -CertificateName $CertificateName -CertificatePwd $CertificatePwd -MonthsValid $MonthsValid -FolderPath $FolderPath -RemoveFromKeyStore $RemoveFromKeyStore -Verbose
