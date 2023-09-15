<# 
.SYNOPSIS
    Find active Skype for Business services on the entire tenant using Microsoft Graph.

.DESCRIPTION 
    Find active Skype for Business services on the entire tenant using Microsoft Graph. Generates an array containing the detected user and license and displays it on the screen.
    It also saves the same data in a CSV file within the folder from which you are launching the script (Current Directory in PowerShell) that can be used for later manipulation of users and licenses (and active Skype for Business plans).
    You can execute this script without parameters and wait for results.
 
.NOTES 
    Filename: findSkypeActivePlans.ps1
    Version: 0.1, 2023
    Author: GSolone
    Blog: gioxx.org
    Twitter: @gioxx

    Changes:
        13/9/23- Change: I take out currentSkuId and UserPrincipalName which are needed later to "feed" the Remove script.
        8/9/23- First version of the script.

.COMPONENT 
    -

.LINK 
    https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.identity.directorymanagement/get-mgsubscribedsku?view=graph-powershell-1.0
    https://learn.microsoft.com/it-it/microsoft-365/enterprise/view-licenses-and-services-with-microsoft-365-powershell?view=o365-worldwide
    https://learn.microsoft.com/it-it/azure/active-directory/enterprise-users/licensing-service-plan-reference#service-skype-for-business
    https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.users/get-mguserlicensedetail?view=graph-powershell-1.0
    https://learn.microsoft.com/en-us/microsoft-365/enterprise/disable-access-to-services-with-microsoft-365-powershell?view=o365-worldwide
#>

function priv_SaveFileWithProgressiveNumber($path) {
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($path)
    $extension = [System.IO.Path]::GetExtension($path)
    $directory = [System.IO.Path]::GetDirectoryName($path)
    $count = 1
    while (Test-Path $path)
    {
        $fileName = $baseName + "_$count" + $extension
        $path = Join-Path -Path $directory -ChildPath $fileName
        $count++
    }
    return $path
}

Set-Variable ProgressPreference Continue
$Result = @()
$ProcessedCount = 0

# Check https://learn.microsoft.com/it-it/azure/active-directory/enterprise-users/licensing-service-plan-reference#service-skype-for-business
$skypePlans = @(
    "afc06cb0-b4f4-4473-8286-d644f70d8faf",
    "b2669e95-76ef-4e7e-a367-002f60a39f3e",
    "0feaeb32-d00e-4d66-bd5a-43b5b83db82c",
    "70710b6b-3ab4-4a38-9f6d-9f169461650a"
)
$skypePlansActive = 0
$skypePlansActiveUsers = @()

$Users = Get-MgUser -Filter 'assignedLicenses/$count ne 0' -ConsistencyLevel eventual -CountVariable totalUsers -All
$Users | ForEach {
    $ProcessedCount++
    $PercentComplete = (($ProcessedCount / $totalUsers) * 100)
    $User = $_
    Write-Progress -Activity "Processing $($User.DisplayName)" -Status "$ProcessedCount out of $totalUsers ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete
    $GraphLicense = Get-MgUserLicenseDetail -UserId $User.Id
    if ( $GraphLicense -ne $null ) {
    ForEach ( $License in $($GraphLicense.SkuPartNumber) ) {
        $currentSkuId = $GraphLicense | ? { $_.SkuPartNumber -eq $license } | select -ExpandProperty SkuId
        $servicePlans = $GraphLicense | ? { $_.SkuPartNumber -eq $license } | select -ExpandProperty ServicePlans
        $mcoStatus = $servicePlans | ? { $_.ServicePlanId -in $skypePlans }
        if ( $mcoStatus -ne $null -and $mcoStatus.ProvisioningStatus -eq "Success" ) { 
            $skypePlansActive++
            $skypePlansActiveUsers += New-Object -TypeName PSObject -Property $([ordered]@{
                DisplayName = $User.DisplayName
                UserPrincipalName = $User.UserPrincipalName
                SMTPAddress = $User.Mail
                SkuId = $currentSkuId
                SkypePlan = $mcoStatus.ProvisioningStatus
                SkypePlanName = $mcoStatus.ServicePlanName
                SkypePlanId = $mcoStatus.ServicePlanId
            })
            break
        }
    }
    }
}


if ( $skypePlansActive -gt 0 ) {
    $CSV = priv_SaveFileWithProgressiveNumber("$($PWD)\$((Get-Date -format "yyyyMMdd").ToString())_M365-Skype-Active-Users.csv")
    Write-Host "`nSkype Active Plans found: $skypePlansActive`nAlso saved as $CSV" -f "Yellow"
    $skypePlansActiveUsers | Select DisplayName, SMTPAddress, SkypePlanName | Out-Host
    $skypePlansActiveUsers | Export-CSV $CSV -NoTypeInformation -Encoding UTF8 -Delimiter ";"
}