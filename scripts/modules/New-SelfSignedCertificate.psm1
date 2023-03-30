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
$VerbosePreference = 'Continue'
$DebugPreference = 'SilentlyContinue' # 'Continue'

<# ---------------- Program execution ---------------- #>

Function New-SelfSignedCertificate {
  Param(
    [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $CertificateName,
    [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $CertificatePwd,
    [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [int] $MonthsValid,
    [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $FolderPath,
    [Parameter(Mandatory = $false)] [boolean] $RemoveFromKeyStore
  )

  Try {
    If (($FolderPath | Test-Path) -eq $false) {
      throw "The specified folder path does not exist."
    }

    If (!$RemoveFromKeyStore) {
      $RemoveFromKeyStore = $false
    }
    
    If ($FolderPath.EndsWith('/') -eq $false) {
      $FolderPath = "$($FolderPath)/"
    }
    
    $CertificatePath = $FolderPath + $CertificateName
    
    $Certificate = New-Certificate -CertificateName $CertificateName -CertificatePath $CertificatePath -MonthsValid $MonthsValid

    $PrivateKey = New-PrivateKey -Certificate $Certificate -CertificatePath $CertificatePath -CertificatePwd $CertificatePwd
    
    If ($RemoveFromKeyStore -eq $true) {
      Remove-CertificateFromPersonalKeyStore -CertificateName $CertificateName
    }

    $CertificateBase64 = Get-Base64 -FilePath "$($CertificatePath).cer"
    $PrivateKeyBase64 = Get-Base64 -FilePath "$($CertificatePath).pfx"
    
    Write-Verbose -Message "Output folder: $($CertificatePath)"
    Write-Verbose -Message "Certificate thumbnail: $($Certificate.Thumbprint)"
    Write-Verbose -Message "Certificate password: $($CertificatePwd)"
    Write-Verbose -Message "Certificate base64 value: $($CertificateBase64)"
    Write-Verbose -Message "Private key base64 value: $($PrivateKeyBase64)"

    Finish-Up
  }
  Catch [Exception] {
    Write-Verbose -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while executing the program. Message: $($_.Exception.Message)"

    Finish-Up -TerminateWithError $true
  }
}

Export-ModuleMember -Function New-SelfSignedCertificate

<# ---------------- Helper functions ---------------- #>

Function New-Certificate {
  Param(
    [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $CertificateName,
    [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $CertificatePath,
    [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [int] $MonthsValid
  )

  Try {
    Write-Verbose -Message "Creating certificate..."
    Write-Debug -Message "CertificateName: $($CertificateName)"
    Write-Debug -Message "CertificatePath: $($CertificatePath)"
    Write-Debug -Message "MonthsValid: $($MonthsValid)"

    $Certificate = PKI\New-SelfSignedCertificate -Subject "CN=$($CertificateName)" -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable -KeySpec Signature -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256 -NotAfter (Get-Date).AddMonths($MonthsValid) -ErrorVariable NewSelfSignedCertificatedError -Verbose:$false
    
    If ($false -eq [string]::IsNullOrEmpty($NewSelfSignedCertificatedError)) {
      throw "Could not create the self signed certificate. Reason: $($NewSelfSignedCertificatedError)"
    }
    Export-Certificate -Cert $Certificate -FilePath "$($CertificatePath).cer" -Verbose:$false | Out-Null

    Write-Debug -Message "Certificate: $($Certificate)"

    return $Certificate
  }
  Catch [Exception] {
    Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to create the certificate. Message: $($_.Exception.Message)"

    Finish-Up -TerminateWithError $true
  }
}

Function New-PrivateKey {
  Param(
    [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [System.Security.Cryptography.X509Certificates.X509Certificate2] $Certificate,
    [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $CertificatePath,
    [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $CertificatePwd
  )

  Try {
    Write-Verbose -Message "Creating private key..."
    Write-Debug -Message "CertificatePath: $($CertificatePath)"
    Write-Debug -Message "CertificatePwd: $($CertificatePwd)"

    $PrivateKey = Export-PfxCertificate -Cert $Certificate -FilePath "$($CertificatePath).pfx" -Password (ConvertTo-SecureString $CertificatePwd -AsPlainText -Force) -Verbose:$false

    If (!$PrivateKey) {
      throw "Could not create the private key for the certificate."
    }

    Write-Debug -Message "PrivateKey: $($PrivateKey)"

    return $PrivateKey
  }
  Catch [Exception] {
    Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to create the private key for the certificate. Message: $($_.Exception.Message)"

    Finish-Up -TerminateWithError $true
  }
}

Function Remove-CertificateFromPersonalKeyStore {
  Param(
    [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $CertificateName
  )

  Try {
    Write-Verbose -Message "Removing certificate from the user's personal keystore..."

    $KeyStoreCertThumbPrint = Get-ChildItem -Path "Cert:\CurrentUser\My" -Verbose:$false | Where-Object { $_.Subject -Match $CertificateName } | Select-Object Thumbprint
    Remove-Item -Path "Cert:\CurrentUser\My\$($KeyStoreCertThumbPrint.Thumbprint)" -DeleteKey -ErrorAction SilentlyContinue -Verbose:$false | Out-Null
  }
  Catch [Exception] {
    Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to remove the certificate from the user's personal keystore. Message: $($_.Exception.Message)"

    Finish-Up -TerminateWithError $true
  }
}

Function Get-Base64 {
  Param(
    [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string] $FilePath
  )

  Try {
    If (($FilePath | Test-Path) -eq $false) {
      throw "Could not get base 64 value of file. Reason: File does not exist."
    }

    # Get private key base64 value
    $FileBytes = $null
    # PowerShell 7 does not support Byte for encoding. 
    # The equivalent is AsByteStream (see also: https://github.com/PowerShell/PowerShell/issues/14537)
    if ($true -eq $PSVersionTable.PSVersion.ToString().StartsWith("7")) {
      $FileBytes = Get-Content $FilePath -AsByteStream -Verbose:$false
    }
    else {
      $FileBytes = Get-Content $FilePath -Encoding Byte -Verbose:$false
    }

    $FileBase64 = [System.Convert]::ToBase64String($FileBytes)

    Write-Debug -Message "FileBase64: $($FileBase64)"

    return $FileBase64
  }
  Catch [Exception] {
    Write-Error -Message "An error occurred on line $($_.InvocationInfo.ScriptLineNumber) while attempting to get the base64 value of file '$($FilePath)'. Message: $($_.Exception.Message)"

    Finish-Up -TerminateWithError $true
  }
}

Function Finish-Up {
  Param(
    [Parameter(Mandatory = $false)] [boolean] $TerminateWithError = $false
  )

  Try {
    # No session to close, do nothing
  }
  Catch [Exception] {
    # Do nothing
  }
  Finally {
    Write-Verbose -Message "Finished."
        
    If ($TerminateWithError) {
      exit 1 # Terminate the program as failed
    }
  }
}