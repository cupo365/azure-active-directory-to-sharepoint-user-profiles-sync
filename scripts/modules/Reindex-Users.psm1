<#
    .SYNOPSIS
        Script to reindex SharePoint users.

    .DESCRIPTION
        This script allows you to reindex SharePoint users so that the SharePoint Search crawl will pick up the changes in the User Profiles faster.
        The script uses a self-proclaimed 'clever' exploit to trigger the SharePoint Search crawl by changing the value of a user property,
        and placing back the original value, therefore updating the user profile property "modified", based on which the crawl decides whether or not
        to scan the user profile for new property values.

    .PARAMETER tenantName <string> [required]
        The name of the tenant.    

    .LINK
        Source of insipration: https://github.com/wobba/SPO-Trigger-Reindex
#>

<# ---------------- Global variables ---------------- #>

$global:fileName = "upa-batch-trigger"
$global:tempPath = $env:USERPROFILE

<# ---------------- Program execution ---------------- #>

Function Reindex-Users (
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$tenantName) {
    Try {
        Initialize $tenantName

        $userProfiles = Get-UserProfiles

        Remove-ExistingTrigger

        Trigger-ReindexUsers $userProfiles
    }
    Catch [Exception] {
        Write-Host "`nAn error occurred: $($_.Exception.Message)" -ForegroundColor Red
    }
    Finally {
        Finish-Up
    }
}

Export-ModuleMember -Function Reindex-Users

<# ---------------- Helper functions ---------------- #>

Function Initialize (
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$tenantName) {
    Try {
        Write-Host "Initializing..."

        Import-Module -Name PnP.PowerShell -Force -ErrorAction SilentlyContinue
        $pnpModule = Get-Module -Name PnP.PowerShell -ErrorAction SilentlyContinue
        if (!$pnpModule) {
            Write-Host "Installing required PowerShell module PnP.PowerShell..."
            Install-Module -Name PnP.PowerShell -Scope CurrentUser -AllowClobber -Force
            Write-Host "Installed module PnP.PowerShell!" -ForegroundColor Green
        }

        Write-Host "Connecting to PnP..."
        Connect-PnPOnline -Url "https://$($tenantName).sharepoint.com" -UseWebLogin | Out-Null
        Write-Host "Connected!" -ForegroundColor Green

        Write-Host "Initialized!`n" -ForegroundColor Green
    }
    Catch [Exception] {
        Write-Host "`nAn error occurred: $($_.Exception.Message)" -ForegroundColor Red
        Finish-Up
        throw $_.Exception
    }
}

Function Get-UserProfiles {
    Try {
        Write-Host "Submitting PnP search query for user profiles..."
        $profiles = Submit-PnPSearchQuery -Query '-AccountName:spofrm -AccountName:spoapp -AccountName:app@sharepoint -AccountName:spocrawler -AccountName:spocrwl -PreferredName:"Foreign Principal"' `
            -SourceId "b09a7990-05ea-4af9-81ef-edfab16c4e31" -SelectProperties "aadobjectid", "department", "write" ` -All -TrimDuplicates:$false -RelevantResults
        
        If ($null -eq $profiles) {
            throw "Search query returned invalid results."
        }
        Write-Host "Successfully executed the search query!" -ForegroundColor Green


        $fragmentTemplate = "{{""IdName"": ""{0}"",""Department"": ""{1}""}}";
        $accountFragments = @();

        $profiles | ForEach-Object {
            $aadId = $_.aadobjectid + ""
            $dept = $_.department + ""

            If (-not [string]::IsNullOrWhiteSpace($aadId) -and $aadId -ne "00000000-0000-0000-0000-000000000000") {
                $accountFragments += [string]::Format($fragmentTemplate, $aadId, $dept)
            }

        }

        If ($accountFragments.Count -eq 0) {
            throw "Query returned 0 user profiles."
        }
            
        Write-Host "Query returned $($accountFragments.Count) user profiles!" -ForegroundColor Green

        return $accountFragments
    }
    Catch [Exception] {
        Write-Host "`nAn error occurred: $($_.Exception.Message)" -ForegroundColor Red
        Finish-Up
        throw $_.Exception
    }
}

Function Remove-ExistingTrigger {
    Try {
        Write-Host "Cleaning up potential existing user reindex triggers."

        $web = Get-PnPWeb
        $folder = $web.GetFolderByServerRelativeUrl("/");
       
        $files = $folder.Files
        $folders = $folder.Folders
        Get-PnPProperty -ClientObject $folder -Property Files, Folders

        $files | ForEach-Object {
            if ($_.Name -like "*$global:fileName*") {
                Write-Host "Found an old import file on the site. Removing file..." -ForegroundColor DarkYellow
                $_.DeleteObject()
            }
        }

        $folders | ForEach-Object {
            if ($_.Name -like "*$global:fileName*") {
                Write-Host "Found an old import status folder on the site. Removing folder..." -ForegroundColor DarkYellow
                $_.DeleteObject()
            }
        }

        Invoke-PnPQuery
        Write-Host "Successfully deleted the old trigger file(s) and folder(s)!" -ForegroundColor Green
    }
    Catch [Exception] {
        Write-Host "`nAn error occurred: $($_.Exception.Message)" -ForegroundColor Red
        Finish-Up
        throw $_.Exception
    }
}

Function Trigger-ReindexUsers (
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][object]$userProfiles) {
    Try {
        Write-Host "Creating reindex trigger file locally..."
        $json = "{""value"":[" + ($userProfiles -join ',') + "]}"
        $json > "$global:tempPath/$global:fileName.txt"
        Write-Host "Successfully created the trigger file locally!" -ForegroundColor Green

        Write-Host "Creating PnP bulk import job..."
        $propertyMap = @{}
        $propertyMap.Add("Department", "Department")
        
        $job = New-PnPUPABulkImportJob -UserProfilePropertyMapping $propertyMap -IdType CloudId -IdProperty "IdName" -Folder "/" -Path "$global:tempPath/$global:fileName.txt"
        
        If ($null -eq $job) {
            throw "Could not create PnP bulk import job..."
        }
        
        If ($job.Error -ne "NoError") {
            throw "An error occured while creating the PnP bulk import job: $($job.Error)."
        }

        Write-Host "Successfully created the bulk import job!" -ForegroundColor Green

        Write-Host "`nThe UPA batch trigger file can be found here: https://$($tenantName).sharepoint.com/$($global:fileName).txt"

        Write-Host "`nPlease be patient and allow 4-24h to pass before profiles are updates in search.`nDo NOT re-run because you are impatient!" -ForegroundColor DarkYellow
        
        Write-Host "`nSuccessfully created the trigger users reindex job!" -ForegroundColor Green
        Write-Host "Current job status: $($job.State)"
        $job
        Write-Host "`nYou can re-check the status of your job with command: Get-PnPUPABulkImportStatus -JobId $($job.JobId)"
        Write-Host "Please allow 1-2h for the job to execute.`n"
    }
    Catch [Exception] {
        Write-Host "`nAn error occurred: $($_.Exception.Message)" -ForegroundColor Red
        Finish-Up
        throw $_.Exception
    }
}

Function Finish-Up {
    Try {
        Write-Host "`n"

        $localFileExists = Test-Path "$global:tempPath/$global:fileName.txt"
        If ($localFileExists -eq $true) {
            Write-Host "Removing local trigger file..."
            Remove-Item -Path "$global:tempPath/$global:fileName.txt"
            Write-Host "Successfully removed the local trigger file!" -ForegroundColor Green
        } 
        
        Write-Host "`nDisconnecting the PnP session..."
        Disconnect-PnPOnline | Out-Null
        Write-Host "Successfully disconnected the session!" -ForegroundColor Green
       
    }
    Catch [Exception] {
        Write-Host "`nAn error occurred: $($_.Exception.Message)" -ForegroundColor Red
        throw $_.Exception
    }
}