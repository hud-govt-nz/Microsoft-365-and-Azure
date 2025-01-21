<#
.SYNOPSIS
This script manages the list of approved senders for a specified Distribution Group or Dynamic Distribution Group in Exchange Online.

.DESCRIPTION
The script provides three actions: adding, removing, or reviewing the list of approved senders. The user is prompted to enter the name of the Distribution Group or Dynamic Distribution Group, and then to specify the action they want to perform. If adding or removing, the user will be prompted to enter the email addresses to be added or removed, respectively.

.PARAMETER DL
The name of the Distribution Group or Dynamic Distribution Group to manage.

.PARAMETER action
The action to perform: 'add', 'remove', or 'review'.

.EXAMPLE
.\ManageDL.ps1 -DL "DL-AllUsers" -action "review"

This example reviews the list of approved senders for the specified Distribution Group or Dynamic Distribution Group.

.NOTES
- Ensure you have the necessary permissions in Exchange Online to manage Distribution Groups and Dynamic Distribution Groups.
- It's advisable to test the script in a safe environment before using it in a production environment.
#>

# Region Parameters
[CmdletBinding()]
Param (
    [Parameter(Mandatory = $true, HelpMessage="Please enter the Distribution List name")]
    #[ValidateNotNullOrEmpty()]
    [string]$DL,
    
    [Parameter(Mandatory = $true, HelpMessage="Please choose an action: add, remove, or review")]
    [ValidateSet("add", "remove", "review")]
    [string]$action = "review"  # Default value
)

Clear-Host
Write-Host '## Exchange Online: Update DL Approved Senders ##' -ForegroundColor Yellow

# Requirements
#Requires -Modules ExchangeOnlineManagement

# Connect to Graph and Exchange
try {
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

# Region Functions
function Get-GroupType {
    param (
        [string]$DLName
    )
    $DG = Get-DistributionGroup -Identity $DLName -ErrorAction SilentlyContinue
    $DDG = Get-DynamicDistributionGroup -Identity $DLName -ErrorAction SilentlyContinue
    if ($DG) { return "DistributionGroup" }
    elseif ($DDG) { return "DynamicDistributionGroup" }
    else { return $null }
}

function Get-ExistingAddresses {
    param (
        [string]$DLName,
        [string]$GroupType
    )
    $cmdlet = "Get-$GroupType"
    & $cmdlet -Identity $DLName |
    Select-Object -ExpandProperty AcceptMessagesOnlyFromSendersOrMembers |
    Get-Recipient |
    Select-Object -ExpandProperty PrimarySmtpAddress
}

function Update-DL {
    param (
        [string]$DLName,
        [array]$UpdatedAddresses,
        [string]$GroupType
    )
    $cmdlet = "Set-$GroupType"
    & $cmdlet -Identity $DLName -AcceptMessagesOnlyFromSendersOrMembers $UpdatedAddresses
}

# End Region Functions

# Determine group type
$GroupType = Get-GroupType -DLName $DL
if (-not $GroupType) {
    Write-Host "$DL is not found or it's neither a Distribution Group nor a Dynamic Distribution Group."
    return
}

# Perform action based on user input
switch ($action) {
    'add' {
        $addressesToAdd = Read-Host "Enter addresses to add (comma separated)"
        $existingAddresses = Get-ExistingAddresses -DLName $DL -GroupType $GroupType
        $updatedAddresses = $existingAddresses + ($addressesToAdd -split ',')
        Update-DL -DLName $DL -UpdatedAddresses $updatedAddresses -GroupType $GroupType

        Write-Host ""
        Write-Host "The following people have been granted permission to send from $DL " -ForegroundColor Green
        Write-Host "Email: $addressesToAdd"
        Write-Host ""
        
    }
    'remove' {
        $addressesToRemove = Read-Host "Enter addresses to remove (comma separated)"
        $existingAddresses = Get-ExistingAddresses -DLName $DL -GroupType $GroupType
        $updatedAddresses = $existingAddresses | Where-Object { ($addressesToRemove -split ',') -notcontains $_ }
        Update-DL -DLName $DL -UpdatedAddresses $updatedAddresses -GroupType $GroupType

        Write-Host ""
        Write-Host "The following people not longer have permissions to send from $DL " -ForegroundColor Green
        Write-Host "Email: $addressesToRemove"
        Write-Host ""
    }
    'review' {
        Write-Host ""
        Write-Host "The following people have permission to send from $DL " -ForegroundColor Green
        Write-Host ""
        Get-ExistingAddresses -DLName $DL -GroupType $GroupType | Sort-Object | Format-Table
        Write-Host ""
    }
    default {
        Write-Host "Invalid action. Please enter add, remove, or review."
    }
}

Disconnect-ExchangeOnline -Confirm:$false | Out-Null