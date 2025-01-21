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

#getTeamsUserActivityUserDetail
$URI      = "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserDetail(period='D180')"
$response = (Invoke-RestMethod -Method GET -Uri $URI -Headers @{Authorization = "Bearer $token "}) -Replace "ï»¿", ""
$Report    = $response | ConvertFrom-Csv

return $Report | Select-Object -First 1

#getTeamsUserActivityCounts
$URI      = "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityCounts(period='D180')"
$response = (Invoke-RestMethod -Method GET -Uri $URI -Headers @{Authorization = "Bearer $token "}) -Replace "ï»¿", ""
$Report    = $response | ConvertFrom-Csv

return $Report | Select-Object -First 1

#getTeamsUserActivityUserCounts
$URI      = "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserCounts(period='D180')"
$response = (Invoke-RestMethod -Method GET -Uri $URI -Headers @{Authorization = "Bearer $token "}) -Replace "ï»¿", ""
$Report    = $response | ConvertFrom-Csv

return $Report | Select-Object -First 1