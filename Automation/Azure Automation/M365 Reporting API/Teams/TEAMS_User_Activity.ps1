#______________________________________________________________________________________________________________________________________________________
# Requires -Modules Microsoft.Graph.Authentication
#______________________________________________________________________________________________________________________________________________________

# Function to connect to Microsoft Graph
try {
    Connect-MgGraph `
        -ClientId $env:DigitalSupportAppID `
        -TenantId $env:DigitalSupportTenantID `
        -CertificateThumbprint $env:DigitalSupportCertificateThumbprint `
        -NoWelcome
    Write-Host "Connected to Graph" -ForegroundColor Green

    $CollectToken = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/users" -ContentType "txt" -OutputType HttpResponseMessage
    $Token        = $CollectToken.RequestMessage.Headers.Authorization.Parameter
    $Token | Out-Null
    } catch {
        Write-Host "Error connecting to Microsoft Graph. Please check your credentials and network connection." -ForegroundColor Red
        exit 1
}

#getOffice365ActiveUserDetail
$URI      = "https://graph.microsoft.com/v1.0/reports/getOffice365ActiveUserDetail(period='D180')"
$response = (Invoke-RestMethod -Method GET -Uri $URI -Headers @{Authorization = "Bearer $token "}) -Replace "ï»¿", ""
$Report    = $response | ConvertFrom-Csv

#getOffice365ActiveUserCounts
$URI      = "https://graph.microsoft.com/v1.0/reports/getOffice365ActiveUserCounts(period='D180')"
$response = (Invoke-RestMethod -Method GET -Uri $URI -Headers @{Authorization = "Bearer $token "}) -Replace "ï»¿", ""
$Report    = $response | ConvertFrom-Csv

#getOffice365ServicesUserCounts
$URI      = "https://graph.microsoft.com/v1.0/reports/getOffice365ServicesUserCounts(period='D180')"
$response = (Invoke-RestMethod -Method GET -Uri $URI -Headers @{Authorization = "Bearer $token "}) -Replace "ï»¿", ""
$Report    = $response | ConvertFrom-Csv