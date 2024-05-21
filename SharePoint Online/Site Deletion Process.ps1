# Connect to Microsoft Graph and Complicance PowerSHell
Connect-MgGraph -NoWelcome
Connect-IPPSSession -ShowBanner:$false

# Step 1: Site To be removed from Relevant DMS Register


# Define the site you want to interact with as a parameter
$SiteURL = "https://mhud.sharepoint.com/sites/enterpriseprog"

# Define the name of the retention policy you want to interact with
$PolicyName = "Item Retention - 7 years default"


# Get the relevant retention policy
$Policy = Get-RetentionCompliancePolicy -Identity $PolicyName

return $Policy | Format-List *


#EXAMPLE
Set-RetentionCompliancePolicy -Identity $PolicyName -AddModernGroupLocationException "enterpriseprog@hud.govt.nz"


