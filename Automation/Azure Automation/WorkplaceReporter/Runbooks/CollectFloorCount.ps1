<#
.SYNOPSIS
   Script for collecting and analyzing user sign-ins from Microsoft Graph Audit Logs,
   calculating floor counts, and sending a summarized report via email.

.DESCRIPTION
   This script connects to Microsoft Graph and SharePoint Online, retrieves user sign-ins,
   filters data, excludes specified users, calculates floor counts, updates a SharePoint list,
   and sends an email with the summarized floor count statistics.

.NOTES
   File Name      : FloorCountScript.ps1
   Author         : Ashley Forde
   Prerequisite   : PowerShell, Microsoft Graph PowerShell Module, SharePoint PnP PowerShell Module
#>


#Connect to Microsoft Graph
Connect-MgGraph -Identity -nowelcome 

# Connect to SharePoint Online
$siteUrl = "https://mhud.sharepoint.com/sites/infomgmt"
$pnpConnection = Connect-PnPOnline -Url $siteUrl -ManagedIdentity
$pnpConnection

# Set the timezone to NZ Standard Time
Set-TimeZone -Id "New Zealand Standard Time"

# Assuming New Zealand Time (NZT) is UTC+12 or UTC+13 during daylight saving time
$nztOffset = if ([TimeZoneInfo]::Local.IsDaylightSavingTime((Get-Date))) { 13 } else { 12 }
$currentNztTime = (Get-Date).Date

# Calculate the start and end times in NZT
# For a 24-hour window starting from the current day at 00:00:00 NZT
$startTimeNzt = $currentNztTime.Date.AddDays(-1)
$endTimeNzt = $startTimeNzt.AddDays(1)

# Convert these NZT times back to UTC
$startTimeUtc = $startTimeNzt.AddHours(-$nztOffset)
$endTimeUtc = $endTimeNzt.AddHours(-$nztOffset)

# Format the start and end times for the filter query
$startTimeStr = $startTimeUtc.ToString("yyyy-MM-ddTHH:mm:ssZ")
$endTimeStr = $endTimeUtc.ToString("yyyy-MM-ddTHH:mm:ssZ")

# Format the date period
$datePeriod = "$($startTimeNzt.ToString('dd/MM/yyyy'))"

Write-Output "Date being collected: $datePeriod"

# Define the IP addresses to filter by
$ipAddress1 = "203.167.143.72"
$ipAddress2 = "203.167.143.80"
$ipAddress3 = "203.97.18.66"
$ipAddress4 = "203.97.25.26"
$ipAddress5 = "203.167.151.106"
$ipAddress6 = "203.97.3.186"

# Update the filter query
$filter = "createdDateTime ge " + $startTimeStr + " and createdDateTime lt " + $endTimeStr + " and (ipAddress eq '" + $ipAddress1 + "' or ipAddress eq '" + $ipAddress2 + "' or ipAddress eq '" + $ipAddress3 + "' or ipAddress eq '" + $ipAddress4 + "' or ipAddress eq '" + $ipAddress5 + "' or ipAddress eq '" + $ipAddress6 + "')"

# Get the audit log sign-ins with the filter
$auditLogs = Get-MgAuditLogSignIn -All -Filter $filter

# Initialize an array to store the combined results
$combinedResults = @()

 # Add the UserPrincipalNames of users to be excluded
$excludedUsers = @(
    "DigitalSupport@hud.govt.nz",
    "AppAdmin@mhud.onmicrosoft.com",
    "AppMinDevTest@hud.govt.nz",
    "AwareGroup_Test@hud.govt.nz",
    "AwareGroup_Test2@hud.govt.nz",
    "Dynamics365.ServiceAccount@hud.govt.nz",
    "EPTest@hud.govt.nz",
    "External.Manager@hud.govt.nz",
    "N6ICJpVNdhFSmDr@mhud.onmicrosoft.com",
    "admin@mhud.onmicrosoft.com",
    "HUD_Data_Lab@hud.govt.nz",
    "Hud_Data_Lab_Test@hud.govt.nz",
    "HUDDigital1@hud.govt.nz",
    "HUDMainNumber@hud.govt.nz",
    "ILAdmin@mhud.onmicrosoft.com",
    "ILService@mhud.onmicrosoft.com",
    "IL_Kate_admin@mhud.onmicrosoft.com",
    "MSAdminInt@hud.govt.nz",
    "mwaas_wdgsoc@hud.govt.nz",
    "mwaas_soc_ro@hud.govt.nz",
    "MSTest@hud.govt.nz",
    "svc-oms-alerts@hud.govt.nz",
    "Poly.RealConnect@hud.govt.nz",
    "svc-ppm@hud.govt.nz",
    "PPMTest.Account@hud.govt.nz",
    "PPMTest.Account2@hud.govt.nz",
    "Scanning@hud.govt.nz",
    "ServiceDesk@hud.govt.nz",
    "SP_Upload@hud.govt.nz",
    "SurfaceHub1@hud.govt.nz",
    "SurfaceHub2@hud.govt.nz",
    "svc_ministernal_sp@hud.govt.nz",
    "svc_notification@mhud.onmicrosoft.com",
    "SVC_Teams_DDI_Manager@hud.govt.nz",
    "svc_valo_admin@hud.govt.nz",
    "svc_brd_noreply@hud.govt.nz",
    "svc-Intune@hud.govt.nz",
    "svc-ReportingLink@hud.govt.nz",
    "svc-TFSNowSync@hud.govt.nz",
    "TeleconferenceBridgeOne@hud.govt.nz",
    "TeleconferenceBridgeThree@hud.govt.nz",
    "TeleconferenceBridgeTwo@hud.govt.nz",
    "test.user@hud.govt.nz",
    "TestUser.Procurement@hud.govt.nz"
)

# Iterate through the audit logs
foreach ($log in $auditLogs) {
   
    # Check if the user should be excluded
    if ($excludedUsers -notcontains $log.UserPrincipalName -and $log.UserPrincipalName -like "*@hud.govt.nz") {
        # Get user details from Microsoft Graph
        $userDetails = Get-MgUser -UserId $log.UserPrincipalName

        # Convert UTC to NZT for each log entry
        $log.CreatedDateTime = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($log.CreatedDateTime, [System.TimeZoneInfo]::Utc.Id, "New Zealand Standard Time")

        # Combine the information
        $combinedObject = [PSCustomObject]@{
            AuditLogId = $log.Id
            UserDisplayName = $log.UserDisplayName
            UserPrincipalName = $log.UserPrincipalName
            ClientAppUsed = $log.ClientAppUsed
            IPAddress = $log.IpAddress
            SignInDateTime = $log.CreatedDateTime
            UserLocation = $userDetails.OfficeLocation # Assuming location is stored in OfficeLocation
        }

        # Add to the results array
        $combinedResults += $combinedObject
    }
}

# Obtain 1 result per each user detected in Audit Log
$uniqueUsers = $combinedResults | Sort-Object SignInDateTime | Group-Object UserPrincipalName | ForEach-Object {
    $_.Group | Select-Object -First 1
}

# Obtain result of that user if they are at ipAddress3 and add to a separate Array
$ipAddress3Users = $combinedResults | Where-Object { $_.IPAddress -eq "ipAddress3" }

# Collect locations of each unique user
$locationCounts = $uniqueUsers | Group-Object -Property UserLocation | Select-Object @{Name='OfficeLocation'; Expression={$_.Name}}, @{Name='Count'; Expression={$_.Count}}

# Transform data and place loations into object array
$officeLocationCounts = @{}
$locationCounts | ForEach-Object {
    $officeLocationCounts[$_.OfficeLocation] = $_.Count
}

# Obtaining counts for each location
$AKL6 = $officeLocationCounts["Auckland - APO - Level 6"] + $officeLocationCounts["APO Auckland"]
$AKL7 = $officeLocationCounts["Auckland - APO - Level 7"] + $officeLocationCounts["Auckland Level 7"]
$NewAPO = $ipAddress3Users.Count
$WLN6 = $officeLocationCounts["Wellington - 7WQ - Level 6"]
$WLN7 = $officeLocationCounts["Wellington - 7WQ - Level 7"]
$WLN8 = $officeLocationCounts["Wellington - 7WQ - Level 8"]
$WLN9 = $officeLocationCounts["Wellington - 7WQ - Level 9"]
$PARLIAMENT = $officeLocationCounts["Wellington - Parliament"]

# Totals
$AKLTOTAL = ($locationCounts | Where-Object { $_.OfficeLocation -like "*Auckland*"   } | Measure-Object -Property Count -Sum).Sum + $NewAPO
$WLNTOTAL = ($locationCounts | Where-Object { $_.OfficeLocation -like "*Wellington*" } | Measure-Object -Property Count -Sum).Sum
$TOTAL = ($locationCounts | Measure-Object -Property Count -Sum).Sum + $NewAPO

# Updating sharepoint list at $siteurl location
$listName = "Test_Floor_Count_List"
$newItem = @{ 
    "Title" = "0"
    "DATE" = $datePeriod
    "WLN7WQ_x002d_L6" = $WLN6
    "WLN7WQ_x002d_L7" = $WLN7
    "WLN7WQ_x002d_L8" = $WLN8
    "WLN7WQ_x002d_L9" = $WLN9
    "AKLAPO_x002d_L6" = $AKL6
    "AKLAPO_x002d_L7" = $AKL7
    "NEWAPO" = $NewAPO
    "WLNTOTAL" = $WLNTOTAL
    "AKLTOTAL" = $AKLTOTAL
    "OTHER" = $PARLIAMENT
    "TOTAL" = $TOTAL

}
$ListOutput = Add-PnPListItem -List $listName -Values $newItem -ContentType Number
$ListOutput

# Introduction line
$introduction = "Good morning,<br><br>Please find below the floor count break down for date: $datePeriod"

# Convert $locationCounts to HTML table with improved styling and introduction
$htmlTable = $locationCounts | ConvertTo-Html -Fragment -Property OfficeLocation, Count -PreContent "<style>table { border-collapse: collapse; width: 100%; } th, td { border: 1px solid #dddddd; text-align: left; padding: 8px; } th { background-color: #f2f2f2; }</style><p>$introduction</p>"

# Add separate total counts for Wellington and Auckland to the HTML table
$htmlTable += "<p>Total Wellington Count: $WLNTOTAL</p>"
$htmlTable += "<p>Total Auckland Count: $AKLTOTAL</p>"
$htmlTable += "<p>Total Other Locations Count: $PARLIAMENT</p>"
$htmlTable += "<p>Total Overall Count: $TOTAL</p>"

# Email Results
$params = @{
    Message = @{
        Subject = "Floor Count $datePeriod"
        Body = @{
            ContentType = "html"
            Content = "$htmlTable"
        }
        ToRecipients = @(
            @{
                EmailAddress = @{
                    Address = "Ashley.Forde@hud.govt.nz"
                }
            }
        )
    }
    SaveToSentItems = "false"
}

# Send email
Send-MgUserMail -UserId "DigitalSupport@hud.govt.nz" -BodyParameter $params

Write-Output "Auckland Stats `n--------------"
Write-Output "`nAPO Auckland: $AKL6"
Write-Output "Auckland Level 7: $AKL7"
Write-Output "Auckland Total: $AKLTOTAL"
Write-Output "`nWellington Stats `n----------------"
Write-Output "`nWLN Level 6: $WLN6"
Write-Output "WLN Level 7: $WLN7"
Write-Output "WLN Level 8: $WLN8"
Write-Output "WLN Level 9: $WLN9"
Write-Output "WLN PS User: $OTHER"
Write-Output "Wellington Total: $WLNTOTAL"
Write-Output "`n----------------------"
Write-Output "Overall Total: $TOTAL"
Write-Output "`n----------------------"
Write-Output "Script Completed"