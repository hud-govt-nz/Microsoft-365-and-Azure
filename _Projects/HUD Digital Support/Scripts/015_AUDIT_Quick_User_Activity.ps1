Clear-Host
Write-Host "## User Audit Log Search ##" -ForegroundColor Yellow

# Requirements
#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Beta

# Connect to Graph
try {
    Connect-MgGraph `
        -ClientId $env:DigitalSupportAppID `
        -TenantId $env:DigitalSupportTenantID `
        -CertificateThumbprint $env:DigitalSupportCertificateThumbprint `
        -NoWelcome
    Write-Host "Connected" -ForegroundColor Green
} catch {
    Write-Host "Error connecting to Microsoft Graph. Please check your credentials and network connection." -ForegroundColor Red
    exit 1
}

# User input
$FormatUPN      = Read-Host "Enter username (e.g., first.last@hud.govt.nz)"
$StartDateInput = Read-Host "Enter start date (format: dd/MM/yyyy)"
$EndDateInput   = Read-Host "Enter end date (format: dd/MM/yyyy)"

# Set timezone and date format
$nzTimeZone = [TimeZoneInfo]::FindSystemTimeZoneById("New Zealand Standard Time")
$dateFormat = "dd/MM/yyyy HH:mm:ss"

# Append "00:00:00" to the user-provided start and end dates
$StartDateInput += " 00:00:00"
$EndDateInput   += " 23:59:59"

# Parse date time input to an object 
$StartDate = [datetime]::ParseExact($StartDateInput, $dateFormat, [System.Globalization.CultureInfo]::InvariantCulture)
$EndDate   = [datetime]::ParseExact($EndDateInput, $dateFormat, [System.Globalization.CultureInfo]::InvariantCulture)
#[System.DateTime]::Parse("$StartDateInput")
# Convert the start and end dates to UTC
$StartDateUTC = [TimeZoneInfo]::ConvertTimeToUtc($StartDate, $nzTimeZone)
$EndDateUTC   = [TimeZoneInfo]::ConvertTimeToUtc($EndDate, $nzTimeZone)

# Construct parameters for New-MgBetaSecurityAuditLogQuery
$params = @{
    "@odata.type"            = "#microsoft.graph.security.auditLogQuery"
    displayName              = "Audit Log Query $($FormatUPN)"
    filterStartDateTime      = $StartDateUTC
    filterEndDateTime        = $EndDateUTC
    userPrincipalNameFilters = @($FormatUPN)
    # Add other filters as needed
}

# Collecting results using New-MgBetaSecurityAuditLogQuery
$Search   = New-MgBetaSecurityAuditLogQuery -BodyParameter $params -ErrorAction Stop
$SearchID = $Search.id

# Initialize variables for polling
$maxRetries   = 2000
$retryCount   = 0
$searchStatus = ""
$startTime    = Get-Date

# Function to update the console with the current status
function Update-ConsoleStatus {
    param (
        [string]$status,
        [TimeSpan]$elapsedTime
    )
    [Console]::SetCursorPosition(0, [Console]::CursorTop)
    Write-Host "Current search status: $status, Time elapsed: $($elapsedTime.ToString("hh\:mm\:ss"))"
}

# Poll the search status until it is "succeeded" or max retries are reached
while ($retryCount -lt $maxRetries -and $searchStatus -ne "succeeded") {
    Start-Sleep -Seconds 5
    $searchResult = Get-MgBetaSecurityAuditLogQuery -AuditLogQueryId $SearchID -ErrorAction Stop
    $searchStatus = $searchResult.status
    $elapsedTime  = (Get-Date) - $startTime

    Write-Progress -Activity "Polling Search Status" -Status "Current status: $searchStatus" -PercentComplete (($retryCount / $maxRetries) * 100) -SecondsRemaining (($maxRetries - $retryCount) * 5)
    
    if ($retryCount % 3 -eq 0) {
        Update-ConsoleStatus -status $searchStatus -elapsedTime $elapsedTime
    }

    $retryCount++
}

# Calculate total elapsed time
$totalElapsedTime = (Get-Date) - $startTime

# Handle the search result
if ($searchStatus -eq "succeeded") {
    Write-Host "`nSearch completed successfully." -ForegroundColor Green
    # Process the search results
    $results = Get-MgBetaSecurityAuditLogQuery -AuditLogQueryId $SearchID
    $results | Format-Table
    $resultCount = $results.Count
    Write-Host "Total time taken: $($totalElapsedTime.ToString("hh\:mm\:ss"))" -ForegroundColor Green
    Write-Host "Total results found: $resultCount" -ForegroundColor Green
} else {
    Write-Host "`nSearch did not complete successfully within the retry limit." -ForegroundColor Red
    Write-Host "Total time taken: $($totalElapsedTime.ToString("hh\:mm\:ss"))" -ForegroundColor Red
}