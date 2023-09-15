<# 
.SYNOPSIS
    Remove active Skype for Business services on the entire tenant using Microsoft Graph and a CSV file.

.DESCRIPTION 
    Remove active Skype for Business services on the entire tenant using Microsoft Graph (require findSkypeActivePlans.ps1 script).
    You can run this script by passing the parameter for the CSV file to be used to remove the Skype plan found active.
 
.NOTES 
    Filename: removeSkypeActivePlans.ps1
    Version: 0.1, 2023
    Author: GSolone
    Blog: gioxx.org
    Twitter: @gioxx

    Changes:
        13/9/23- First version of the script.

.COMPONENT 
    -

.LINK 
    https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.identity.directorymanagement/get-mgsubscribedsku?view=graph-powershell-1.0
    https://learn.microsoft.com/it-it/microsoft-365/enterprise/view-licenses-and-services-with-microsoft-365-powershell?view=o365-worldwide
    https://learn.microsoft.com/it-it/azure/active-directory/enterprise-users/licensing-service-plan-reference#service-skype-for-business
    https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.users/get-mguserlicensedetail?view=graph-powershell-1.0
    https://learn.microsoft.com/en-us/microsoft-365/enterprise/disable-access-to-services-with-microsoft-365-powershell?view=o365-worldwide
#>

Param(
    [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true, HelpMessage="CSV file to use (e.g. C:\Temp\20230908_M365-Skype-Active-Users.csv)")]
    [string] $CSV
)

Set-Variable ProgressPreference Continue
$ProcessedCount = 0
$totalUsers = Import-Csv $CSV -Delimiter ";" | Measure-Object | Select-Object -ExpandProperty count
Write-Host "Rows: $($totalUsers)" -f "Yellow"

Import-Csv $CSV -Delimiter ";" | ForEach {
    $ProcessedCount++
    $PercentComplete = (($ProcessedCount / $totalUsers) * 100)
    $User = $_        
    Write-Progress -Activity "Processing $($User.DisplayName) ($($User.SkypePlanName))" -Status "$ProcessedCount out of $totalUsers ($($PercentComplete.ToString('0.00'))%)" -PercentComplete $PercentComplete
    
    $skypePlanToRemove = @(
        @{
            SkuId = $($User.SkuId)
            DisabledPlans = $($User.SkypePlanId)
        }
    )
    
    Set-MgUserLicense -UserId $User.UserPrincipalName -RemoveLicenses @() -AddLicenses $skypePlanToRemove | Out-Null

}