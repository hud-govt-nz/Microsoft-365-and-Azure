#______________________________________________________________________________________________________________________________________________________
# Requires -Modules ExchangeOnlineManagement
# Requires -Modules Microsoft.Graph.Authentication
# Requires -Modules PNP.Powershell
# Requires -Modules MicrosoftTeams
#______________________________________________________________________________________________________________________________________________________

# Function to connect to Exchange Online
try{
    Connect-ExchangeOnline `
        -AppId $env:DigitalSupportAppID `
        -Organization "mhud.onmicrosoft.com" `
        -CertificateThumbprint $env:DigitalSupportCertificateThumbprint `
        -ShowBanner:$false
    Write-Host "Connected to Exchange Online." -ForegroundColor Green
    } catch {
        Write-Host "Error connecting to Exchange Online. Please check your credentials and network connection." -ForegroundColor Red
        exit 1
}
#______________________________________________________________________________________________________________________________________________________

# Function to connect to Security and Compliance PowerShell
try{
    Connect-IPPSSession `
    -AppId $env:DigitalSupportAppID `
    -Organization "mhud.onmicrosoft.com" `
    -CertificateThumbprint "2A5AB205BA76E77499949DCC06919FA367A0CB58" `
    -ShowBanner:$false
    Write-Host "Connected to Security and Compliance PowerShell." -ForegroundColor Green
    } catch {
        Write-Host "Error connecting to Security and Compliance PowerShell. Please check your credentials and network connection." -ForegroundColor Red
        exit 1
}


# Function to purge emails
$Search = New-ComplianceSearch -Name "SPAM - TE ARAHANGA O NGA IWI LIMITED PROPOSAL 2" -ExchangeLocation All -ContentMatchQuery '(Subject: "TE ARAHANGA O NGA IWI LIMITED PROPOSAL")'
$Search2 = New-ComplianceSearch -Name "SPAM - TE ARAHANGA O NGA IWI LIMITED PROPOSAL 3" -ExchangeLocation Emily.Gibson@hud.govt.nz,chra@hud.govt.nz,Rebecca.Maplesden@hud.govt.nz,Malo.Ah-You@hud.govt.nz,Filani.McLean@hud.govt.nz,Michaela.Reilly@hud.govt.nz,MEETWLG805MOBILE@hud.govt.nz,Dan.Shenton@hud.govt.nz,Christina.Chase@hud.govt.nz,Grace.Gentiles-Devery@hud.govt.nz,Fiona.Fitzgerald@hud.govt.nz,Paul.Kos@hud.govt.nz -ContentMatchQuery '(Subject: "TE ARAHANGA O NGA IWI LIMITED PROPOSAL")'

# Start the search
Start-ComplianceSearch -Identity $Search.Identity
Start-ComplianceSearch -Identity $Search2.Identity

# https://learn.microsoft.com/en-us/powershell/module/exchange/new-compliancesearchaction?view=exchange-ps
# Purge the emails
New-ComplianceSearchAction –SearchName “SPAM - TE ARAHANGA O NGA IWI LIMITED PROPOSAL 2” -Purge –PurgeType SoftDelete 

