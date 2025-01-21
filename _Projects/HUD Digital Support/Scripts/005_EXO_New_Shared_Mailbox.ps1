Clear-Host
Write-Host '## Exchange Online: Create New Shared Mailbox ##' -ForegroundColor Yellow

# Requirements
#Requires -Modules Microsoft.Graph.Authentication
#Requires -Modules ExchangeOnlineManagement

# Connect to Graph and Exchange
try {
    Connect-MgGraph `
        -ClientId $env:DigitalSupportAppID `
        -TenantId $env:DigitalSupportTenantID `
        -CertificateThumbprint $env:DigitalSupportCertificateThumbprint `
        -NoWelcome

	Connect-ExchangeOnline `
	    -AppId $env:DigitalSupportAppID `
	    -Organization "mhud.onmicrosoft.com" `
	    -CertificateThumbprint $env:DigitalSupportCertificateThumbprint `
	    -ShowBanner:$false
    Write-Host "Connected" -ForegroundColor Green
        
    } catch {
	    Write-Host "Error connecting. Please check your credentials and network connection." -ForegroundColor Red
	    exit 1
        }

# Create Mailbox
$SharedMBName = Read-Host "Please enter the name of the mailbox"
$Alias = $SharedMBName -replace ' ',""
$EmailAddress = "$Alias" + "@hud.govt.nz"
$Mailbox = New-mailbox -Shared -Name $SharedMBName -Alias $Alias -PrimarySmtpAddress $EmailAddress

# Display Mailbox Info
Get-Mailbox -Identity $SharedMBName | Select-Object Name, alias, PrimarySMTPAddress,isShared | Format-List

#Prompt to add users to mailbox
$title = "Add Users to Mailbox"
$sendasmsg = "Would you like to add users to this mailbox?"
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes"
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No"
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
do {
    $result = $host.ui.PromptForChoice($title, $sendasmsg, $options, 0) 
    switch ($result) {
        '0' {
            $UserMB = Read-Host "Please enter users email address"
            $FullAccess = Add-MailboxPermission -Identity ($Mailbox).Id -User $UserMB -AccessRights FullAccess -InheritanceType All -AutoMapping $true
            $Sendas = Add-RecipientPermission ($Mailbox).Id -AccessRights SendAs -Trustee $UserMB -Confirm:$false
            Write-Host "$($UserMB) has been granted 'Full' and 'Send As' permissions to mailbox $(($Mailbox).Id)" -foregroundcolor cyan
        } 
        '1' {
            break # Exit the loop if user selects "No"
        }
    }
} while ($result -ne '1')

# Disconnect Exchange Online Session
Disconnect-ExchangeOnline -Confirm:$false | Out-Null