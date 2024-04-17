# Connect to Exchange Online
Connect-ExchangeOnline

# Collect all user mailboxes and set as value
$Users = (Get-Mailbox -ResultSize Unlimited -Filter {RecipientType -eq 'UserMailbox'}).UserPrincipalName

# Set Weather Locations
$Locations = @("Wellington, NZ", "Auckland, NZ")

# Get the number of mailboxes
$Users.Count

# Loop through and update the mailboxconfiguration to enable the Weather Icon, Locations and Unit to Celcius
foreach ($User in $Users) {
    Set-MailboxCalendarConfiguration `
    -Identity $User `
    -WeatherEnabled enabled `
    -WeatherLocations $Locations `
    -WeatherUnit Celsius

    Write-Host "Updated settings for $($User)"
   
}
