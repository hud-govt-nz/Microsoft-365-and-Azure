Clear-Host

# Snippet1

# Import CSV file containing display names and email addresses
$mailboxes = Import-Csv -Path "CSV FILE - Headers DisplayName, EmailAddress"

# Connect to MgGraph
$Scopes = @("Directory.Read.All", "Directory.ReadWrite.All", "User.Read.All", "User.ReadWrite.All")
Connect-MgGraph -Scopes $Scopes -NoWelcome | Out-Null
$UPN = (Get-MgContext).Account

# Connect to Exchange Online
Connect-ExchangeOnline -UserPrincipalName $UPN -ShowBanner:$false

foreach ($mailbox in $mailboxes) {
    # Employee Status
    Write-Host ''
    Write-Host $mailbox.DisplayName -ForegroundColor Green
    Write-Host ''
    $User = $mailbox.EmailAddress

    $permissions = Get-MailboxFolderPermission -Identity "${User}:\Calendar"
    $customObjects = $permissions | ForEach-Object {
        [PSCustomObject]@{
            Identity                = $_.Identity
            FolderName              = $_.FolderName
            User                    = $_.User
            AccessRights            = $_.AccessRights
            SharingPermissionFlags  = $_.SharingPermissionFlags
            IsValid                 = $_.IsValid
            ObjectState             = $_.ObjectState
        }
    }

    $customObjects | Format-Table -AutoSize
}


# Snippet2 

# Code to add Editor access for multiple users across multiple rooms
foreach ($mailbox in $mailboxes) {
    $User = $mailbox.EmailAddress
    $AddPermissions = Import-Csv -Path "CSV FILE - Headers EmailAddress, PermissionLevel"  # CSV file containing users and their permissions

    foreach ($addPermission in $AddPermissions) {
        $Delegate = $addPermission.EmailAddress
        $PermissionLevel = $addPermission.PermissionLevel

        # Add the specified permission for the delegate
        Add-MailboxFolderPermission -Identity "${User}:\Calendar" -User $Delegate -AccessRights $PermissionLevel
        Write-Host "Editor access added for $Delegate on $User's calendar."
    }
}