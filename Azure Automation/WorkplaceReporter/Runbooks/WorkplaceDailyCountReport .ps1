#< RUN MANUALLY
Clear-Host

# Connect to Microsoft Graph
Connect-MgGraph -NoWelcome | Out-Null
#>

#Connect to Microsoft Graph
.\GraphLogin.ps1 # block out to run locally.

# Set the timezone to NZ Standard Time
Set-TimeZone -Id "New Zealand Standard Time"

# Assuming New Zealand Time (NZT) is UTC+12 or UTC+13 during daylight saving time
$nztOffset = if ([TimeZoneInfo]::Local.IsDaylightSavingTime((Get-Date))) { 13 } else { 12 }
$currentNztTime = (Get-Date).Date.AddDays(1)

# Calculate the start and end times in NZT
# For a 24-hour window starting from the current day at 00:00:00 NZT
$startTimeNzt = $currentNztTime.Date.AddDays(-2)
$endTimeNzt = $startTimeNzt.AddDays(1)

# Convert these NZT times back to UTC
$startTimeUtc = $startTimeNzt.AddHours(-$nztOffset)
$endTimeUtc = $endTimeNzt.AddHours(-$nztOffset)

# Format the start and end times for the filter query
$startTimeStr = $startTimeUtc.ToString("yyyy-MM-ddTHH:mm:ssZ")
$endTimeStr = $endTimeUtc.ToString("yyyy-MM-ddTHH:mm:ssZ")

# Define the two IP addresses to filter by
$ipAddress1 = "203.167.143.72"
$ipAddress2 = "203.167.143.80"

# Update the filter query
$filter = "createdDateTime ge " + $startTimeStr + " and createdDateTime lt " + $endTimeStr + " and (ipAddress eq '" + $ipAddress1 + "' or ipAddress eq '" + $ipAddress2 + "')" 

# Get the audit log sign-ins with the filter
$auditLogs = Get-MgAuditLogSignIn -All -Filter $filter

# Initialize an array to store the combined results
$combinedResults = @()

 # Add the UserPrincipalNames of users to be excluded
$excludedUsers = @(
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
    if ($excludedUsers -notcontains $log.UserPrincipalName) {
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

# Unique result (avoids duplicate log entries for multiple sign-ins...)
$uniqueUsers = $combinedResults | Sort-Object SignInDateTime | Group-Object UserPrincipalName | ForEach-Object {
    $_.Group | Select-Object -First 1
}

# Group the results by office location and count the occurrences
$locationCounts = $uniqueUsers | Group-Object -Property UserLocation | Select-Object @{Name='OfficeLocation'; Expression={$_.Name}}, @{Name='Count'; Expression={$_.Count}}

# Format the date period
$datePeriod = "$($startTimeNzt.ToString('dd/MM/yyyy')) to $($endTimeNzt.ToString('dd/MM/yyyy'))"

# Introduction line
$introduction = "Good morning,<br><br>Please find the floor count stats for the date period: $datePeriod"

# Calculate separate totals for Wellington and Auckland
$totalWellington = ($locationCounts | Where-Object { $_.OfficeLocation -like "Wellington*" } | Measure-Object -Property Count -Sum).Sum
$totalAuckland = ($locationCounts | Where-Object { $_.OfficeLocation -like "Auckland*" } | Measure-Object -Property Count -Sum).Sum

# Calculate count for other locations
$totalOtherLocations = ($locationCounts | Where-Object { $_.OfficeLocation -notlike "Wellington*" -and $_.OfficeLocation -notlike "Auckland*" } | Measure-Object -Property Count -Sum).Sum

# Calculate total overall sum
$totalOverall = ($locationCounts | Measure-Object -Property Count -Sum).Sum

# Convert $locationCounts to HTML table with improved styling and introduction
$htmlTable = $locationCounts | ConvertTo-Html -Fragment -Property OfficeLocation, Count -PreContent "<style>table { border-collapse: collapse; width: 80%; } th, td { border: 1px solid #dddddd; text-align: left; padding: 8px; } th { background-color: #f2f2f2; }</style><p>$introduction</p>"

# Add separate total counts for Wellington and Auckland to the HTML table
$htmlTable += "<p>Total Wellington Count: $totalWellington</p>"
$htmlTable += "<p>Total Auckland Count: $totalAuckland</p>"

# Add count for other locations to the HTML table
$htmlTable += "<p>Total Other Locations Count: $totalOtherLocations</p>"

# Add total overall sum to the HTML table
$htmlTable += "<p>Total Overall Count: $totalOverall</p>"

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