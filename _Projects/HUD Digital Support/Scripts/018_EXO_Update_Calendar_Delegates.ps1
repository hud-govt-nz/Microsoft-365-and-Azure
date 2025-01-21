Clear-Host
Write-Host '## Exchange Online: Calendar Delegates Access ##' -ForegroundColor Yellow

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
        -ShowBanner: $false
    Write-Host "Connected" -ForegroundColor Green
        
    } catch {
        Write-Host "Error connecting. Please check your credentials and network connection." -ForegroundColor Red
        exit 1
        }

do {
    # Employee Status
    Write-Host ''
    $User = Read-Host "Please enter the email address of the mailbox or press 'q' to quit"


    if ($User -eq 'q') {
    break
    }

    $permissions   = Get-MailboxFolderPermission -Identity "${User}:\Calendar"
    $customObjects = $permissions | ForEach-Object {
    [PSCustomObject]@{
        Identity               = $_.Identity
        FolderName             = $_.FolderName
        User                   = $_.User
        AccessRights           = $_.AccessRights
        SharingPermissionFlags = $_.SharingPermissionFlags
        IsValid                = $_.IsValid
        ObjectState            = $_.ObjectState
        }
    }

    $customObjects | Format-Table -AutoSize

    # Ask the user for their choice
    $choice = Read-Host "Do you want to add or remove a delegate? (Type 'Add', 'Remove', or 'Delegate')"

    # Implement the switch statement
    switch ($choice) {
        "Add" {
            # Code to add a delegate
            $AddPermission = Read-Host "Enter the email address of the delegate to add"
            $permission    = Read-Host "Enter the permission level (e.g., Reviewer, Editor, etc.)"
            Add-MailboxFolderPermission -Identity "${User}:\Calendar" -User $AddPermission -AccessRights $permission
            Write-Host "Delegate added."
        }
        "Remove" {
            # Code to remove a delegate
            $RemovePermission = Read-Host "Enter the email address of the delegate to remove"
        
            try {
                Remove-MailboxFolderPermission -Identity "${User}:\Calendar" -User $RemovePermission -Confirm: $false -ErrorAction Stop
            } catch {
                Write-Warning -Message "The user was not found or has been disabled. To remove the entry please type their display name"
                $RemovePermission = Read-Host "Enter the display name of the delegate to remove"
                Remove-MailboxFolderPermission -Identity "${User}:\Calendar" -User $RemovePermission -Confirm: $false
            }
            Write-Host "Delegate removed."
        }
        "Delegate" {
        # Code to remove a delegate
        $Delegate     = Read-Host "Enter the email address of the delegate to add"
        $PrivateItems = Read-Host "Do they need to view Private Items? (Yes/No)"

        Switch ($PrivateItems) {
            "yes" {
            Add-MailboxFolderPermission -Identity "${User}:\Calendar" -User $Delegate -AccessRights Editor -SharingPermissionFlags Delegate, CanViewPrivateItems -SendNotificationToUser $false -Confirm: $false
            }
            "No" {
            Add-MailboxFolderPermission -Identity "${User}:\Calendar" -User $Delegate -AccessRights Editor -SharingPermissionFlags Delegate -SendNotificationToUser $false -Confirm: $false
            }
        }
            }
            default {
            Write-Host "Invalid choice. Please type 'Add' or 'Remove' or 'Delegate'."
            }
    }
    $permissions = Get-MailboxFolderPermission -Identity "${User}:\Calendar"\
    $customObjects = $permissions | ForEach-Object {
        [PSCustomObject]@{
        Identity               = $_.Identity
        FolderName             = $_.FolderName
        User                   = $_.User
        AccessRights           = $_.AccessRights
        SharingPermissionFlags = $_.SharingPermissionFlags
        IsValid                = $_.IsValid
        ObjectState            = $_.ObjectState
        }
    }
    $customObjects | Format-Table -AutoSize
    } while ($true)

    # Disconnect Exchange Online Session
    Disconnect-ExchangeOnline -Confirm: $false | Out-Null
    Disconnect-MgGraph | Out-Null