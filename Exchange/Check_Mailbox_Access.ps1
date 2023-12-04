Clear-Host

# Connect to Exchange Online
$UPN = Whoami /upn
Connect-ExchangeOnline -UserPrincipalName $UPN -ShowBanner:$false

# Obtain User ID
Write-Host '## Check users mailbox access in Exchange Online ##' -ForegroundColor Yellow
$Email = Read-Host "Please enter the user's email address"

# Gather Information
if (![string]::IsNullOrWhiteSpace($Email)) {
    Write-Host "Searching through Exchange Online will take some time, please be patient..." -ForegroundColor Cyan
    
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
    try {
        $recipientPermissions = $_ | Get-RecipientPermission | Where-Object {($_.Trustee -like "*$Email*") -and ($_.Trustee -notmatch "S-1-5-21") -and ($_.Trustee -notmatch "NT AUTHORITY\SELF")} -ErrorAction Stop
        
        foreach ($permission in $recipientPermissions) {
            $Output += [PSCustomObject]@{
                Identity = $permission.Identity
                User = $permission.Trustee
                AccessRights = $permission.AccessRights -join ', '    
                }
            }

        try {
            $mailboxPermissions = $_ | Get-MailboxPermission | Where-Object {($_.User -like "*$Email*") -and ($_.User -notmatch "S-1-5-21") -and ($_.User -notlike "NT AUTHORITY\SELF")} -ErrorAction Stop
            
            foreach ($permission in $mailboxPermissions) {
                $Output += [PSCustomObject]@{
                    Identity = $permission.Identity
                    User = $permission.User
                    AccessRights = $permission.AccessRights -join ', '
                    }
                }
        
            }
            catch [System.Exception] {
                Write-host "Error retrieving mailbox permissions. Errormessage: $($_.Exception.Message)"
                }
        
    # Create the progress bar
    Write-Progress -Activity "Processing:" -PercentComplete $percentComplete -Status "Mailbox $counter of $totalMailboxes"    

        } catch [System.Exception] {
            Write-host "Error retrieving recipient permissions. Errormessage: $($_.Exception.Message)"
            }
       
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
}}

 # Disconnect Exchange Online Session
 Disconnect-ExchangeOnline -Confirm:$false | Out-Null