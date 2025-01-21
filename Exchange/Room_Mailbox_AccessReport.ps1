# Collect all Room Mailboxes
$Mailboxes = Get-Mailbox -Filter '(RecipientTypeDetails -eq "RoomMailbox")'

# Create lists to hold results
$mailboxPermissionsResults = @()
$calendarDelegateResults = @()

foreach ($mailbox in $Mailboxes) {
    $Email = $mailbox.PrimarySmtpAddress

    # Collect Mailbox Permissions
    $mailboxPermissions = Get-MailboxPermission -Identity $Email | Where-Object {($_.User -notlike "S-1-5-21") -and ($_.User -notlike "NT AUTHORITY\SELF")}
    # Process Mailbox Permissions
    foreach ($permission in $mailboxPermissions) {
        $mailboxPermissionsResults += [PSCustomObject]@{
            Mailbox      = $mailbox.DisplayName
            User         = $permission.User
            FullAccess   = $permission.AccessRights -contains 'FullAccess'
            SendAs       = $false
        }
    }

    # Collect Recipient Permissions
    $recipientPermissions = Get-RecipientPermission -Identity $Email | Where-Object {($_.Trustee -notlike "S-1-5-21") -and ($_.Trustee -notlike "NT AUTHORITY\SELF")}
    # Process Recipient Permissions
    foreach ($permission in $recipientPermissions) {
        # Update if existing, add new if not
        $existing = $mailboxPermissionsResults | Where-Object { $_.User -eq $permission.Trustee -and $_.Mailbox -eq $mailbox.DisplayName }
        if ($existing) {
            $existing.SendAs = $true
        } else {
            $mailboxPermissionsResults += [PSCustomObject]@{
                Mailbox    = $mailbox.DisplayName
                User       = $permission.Trustee
                FullAccess = $false
                SendAs     = $true
            }
        }
    }

    # Collect Calendar Delegate Permissions
    $calendarPermissions = Get-MailboxFolderPermission -Identity "${Email}:\Calendar" | Where-Object {($_.User -notlike "Default") -and ($_.User -notlike "Anonymous")}
    # Process Calendar Delegate Permissions
    foreach ($delegate in $calendarPermissions) {
        $calendarDelegateResults += [PSCustomObject]@{
            Identity               = $delegate.Identity
            FolderName             = $delegate.FolderName
            User                   = $delegate.User
            AccessRights           = $delegate.AccessRights
            SharingPermissionFlags = $delegate.SharingPermissionFlags
            IsValid                = $delegate.IsValid
            ObjectState            = $delegate.ObjectState
        }
    }
}

# Display Mailbox Permissions Results
$mailboxPermissionsResults | Format-Table Mailbox, User, FullAccess, SendAs

# Display Calendar Delegate Results
$calendarDelegateResults | Format-Table Identity, FolderName, User, AccessRights, SharingPermissionFlags, IsValid, ObjectState
