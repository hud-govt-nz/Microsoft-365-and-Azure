$DL = "DL - Proactive Release Weekly Reports"
$CSVFile = "C:\Users\Ashley.Forde\OneDrive - Ministry of Housing and Urban Development\Desktop\Proactive Release Weekly Reports.xlsx"
 
# Connect to Exchange Online
Connect-ExchangeOnline -ShowBanner:$False
 
# Get Existing Members of the Distribution List
 
$DLMembers = Get-DistributionGroupMember -Identity $DL -ResultSize Unlimited | Select -Expand PrimarySmtpAddress

# Import Distribution List Members from CSV
Import-Excel $CSVFile | ForEach-Object {
 
    # Check if the Distribution List contains the particular user
    If ($DLMembers -contains $_.UPN) {
        Write-host -f Yellow "User is already member of the Distribution List:"$_.UPN
        } Else {      
            Add-DistributionGroupMember –Identity $DL -Member $_.UPN
            Write-host -f Green "Added User to Distribution List:"$_.UPN
            }
    }