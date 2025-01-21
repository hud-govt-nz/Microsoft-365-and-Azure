#______________________________________________________________________________________________________________________________________________________
# Requires -Modules ExchangeOnlineManagement
# Requires -Modules Microsoft.Graph.Authentication
# Requires -Modules PNP.Powershell
# Requires -Modules MicrosoftTeams
#______________________________________________________________________________________________________________________________________________________

# Function to connect to Security and Compliance PowerShell
try{
    Connect-IPPSSession `
    -AppId $env:DigitalSupportAppID `
    -Organization "mhud.onmicrosoft.com" `
    -CertificateThumbprint "2A5AB205BA76E77499949DCC06919FA367A0CB58" `
    -ShowBanner:$false
    Clear-Host
    Write-Host "`nConnected to Security and Compliance PowerShell." -ForegroundColor Green
    } catch {
        Write-Host "Error connecting to Security and Compliance PowerShell. Please check your credentials and network connection." -ForegroundColor Red
}

# Retention Policies 

#Exchange Online
#______________________________________________________________________________________________________________________________________________________________________________________
$exchname = "Default Exchange Retention Policy - 7 Years"
$exchnamerule = "Default Exchange Retention Policy Rule"
$exchdescription = "Default policy applying to Exchange Objects"

    # Policy
    $exchpolicy = New-RetentionCompliancePolicy -Name $exchname -Comment $exchdescription -ExchangeLocation All -PublicFolderLocation All -Enabled $false

    # Policy Rule
    New-RetentionComplianceRule -Name $exchnamerule -Policy $exchname -RetentionDuration 2555 -RetentionComplianceAction KeepAndDelete -ExpirationDateOption ModificationAgeInDays




#SharePoint Online
#______________________________________________________________________________________________________________________________________________________________________________________
$sponame = "Default SharePoint Retention Policy - 7 Years"
$sponamerule = "Default SharePoint Retention Policy Rule"
$spodescription = "Default policy applying to SharePoint Objects"

    # Policy
    $spopolicy = New-RetentionCompliancePolicy -Name $sponame -Comment $spodescription -SharePointLocation All -ModernGroupLocation All -OneDriveLocation All -Enabled $false

    # Policy Rule
    New-RetentionComplianceRule -Name $sponamerule -Policy $sponame -RetentionDuration 2555 -RetentionComplianceAction Keep -ExpirationDateOption ModificationAgeInDays

#Teams
#______________________________________________________________________________________________________________________________________________________________________________________
$teamsname = "Default Teams Retention Policy - 7 Years"
$teamsnamerule = "Default Teams Retention Policy Rule"
$teamsdescription = "Default policy applying to Teams Objects"
  
    # Policy
    $teamspolicy = New-RetentionCompliancePolicy -Name $teamsname -Comment $teamsdescription -TeamsChannelLocation All -TeamsChatLocation All -Enabled $false
 
    # Policy Rule
    New-RetentionComplianceRule -Name $teamsnamerule -Policy $teamsname -RetentionDuration 2555 -RetentionComplianceAction KeepAndDelete



#______________________________________________________________________________________________________________________________________________________________________________________

Get-RetentionCompliancePolicy -Identity $exchname -DistributionDetail