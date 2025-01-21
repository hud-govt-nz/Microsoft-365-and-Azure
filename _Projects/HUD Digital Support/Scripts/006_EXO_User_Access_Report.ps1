Clear-Host
Write-Host '## Exchange Online: User Mailbox Access Report ##' -ForegroundColor Yellow

# Requirements
#Requires -Modules ExchangeOnlineManagement

# Connect to Exchange
try {
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

        $Email = Read-Host "Please enter the user's email address"

        # Guard clause for empty email
        if ([string]::IsNullOrWhiteSpace($Email)) {
            Write-Host "Email address cannot be empty." -ForegroundColor Red
            exit
        }
        
        Write-Host "Searching through Exchange Online will take some time, please be patient..." -ForegroundColor Cyan
        
        # Function to get recipient permissions
        function Get-RecipientPermissions($mailbox, $email) {
            try {
                return $mailbox | Get-RecipientPermission | Where-Object {
                    ($_.Trustee -like "*$email*") -and ($_.Trustee -notmatch "S-1-5-21") -and ($_.Trustee -notmatch "NT AUTHORITY\SELF")
                } -ErrorAction Stop
            } catch {
                Write-Host "Error retrieving recipient permissions for $($mailbox.Identity). Error message: $($_.Exception.Message)" -ForegroundColor Red
                return @()
            }
        }
        
        # Function to get mailbox permissions
        function Get-MailboxPermissions($mailbox, $email) {
            try {
                return $mailbox | Get-MailboxPermission | Where-Object {
                    ($_.User -like "*$email*") -and ($_.User -notmatch "S-1-5-21") -and ($_.User -notlike "NT AUTHORITY\SELF")
                } -ErrorAction Stop
            } catch {
                Write-Host "Error retrieving mailbox permissions for $($mailbox.Identity). Error message: $($_.Exception.Message)" -ForegroundColor Red
                return @()
            }
        }
        
        # Gather Output Array for results
        $Mailboxes = Get-Mailbox -ResultSize unlimited
        $Output = @()
        $counter = 0
        
        # Get the total number of mailboxes to process
        $totalMailboxes = $Mailboxes.Count
        Write-Host "Total number of mailboxes to process: $totalMailboxes"
        
        $Mailboxes | ForEach-Object {
            # Increment the counter for each mailbox processed
            $counter++
        
            # Calculate the percentage of mailboxes processed
            $percentComplete = ($counter / $totalMailboxes) * 100
        
            # Get the mailbox permissions and filter for the specified user
            $recipientPermissions = Get-RecipientPermissions $_ $Email
            foreach ($permission in $recipientPermissions) {
                $Output += [PSCustomObject]@{
                    Identity = $permission.Identity
                    User = $permission.Trustee
                    AccessRights = $permission.AccessRights -join ', '
                }
            }
        
            $mailboxPermissions = Get-MailboxPermissions $_ $Email
            foreach ($permission in $mailboxPermissions) {
                $Output += [PSCustomObject]@{
                    Identity = $permission.Identity
                    User = $permission.User
                    AccessRights = $permission.AccessRights -join ', '
                }
            }
        
            # Create the progress bar
            Write-Progress -Activity "Processing:" -PercentComplete $percentComplete -Status "Mailbox $counter of $totalMailboxes"
            
            # Add a line break after each mailbox processed
            Write-Host ""
        }
        
        
# Clear the progress bar
Write-Progress -Activity "Processing:" -Completed

# Display the list of mailboxes the user has access to
if ($Output){
    Clear-Host
    Write-Host "The user $($Email) has access to the following mailboxes:" -ForegroundColor Green
    $Output | Format-Table -AutoSize -Wrap
} else {
    Write-Host "The user $($Email) does not have access to any mailboxes." -ForegroundColor Green
}

 # Disconnect Exchange Online Session
 Disconnect-ExchangeOnline -Confirm:$false | Out-Null