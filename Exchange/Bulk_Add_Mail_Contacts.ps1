# Script to bulk create mail contacts in Exchange Online from CSV file
#Requires -Modules ExchangeOnlineManagement

# Connect to Exchange Online
try {
    # Check if there's a certificate-based auth environment variable
    if ($env:DigitalSupportAppID) {
        Connect-ExchangeOnline `
            -AppId $env:DigitalSupportAppID `
            -Organization "mhud.onmicrosoft.com" `
            -CertificateThumbprint $env:DigitalSupportCertificateThumbprint `
            -ShowBanner:$false
    }
    else {
        # Fall back to interactive login
        Connect-ExchangeOnline -ShowBanner:$false
    }
    Write-Host "Connected to Exchange Online." -ForegroundColor Green
}
catch {
    Write-Host "Error connecting to Exchange Online. Please check your credentials and network connection." -ForegroundColor Red
    exit 1
}

# Function to validate email address format
function Test-EmailAddress {
    param([string]$EmailAddress)
    return $EmailAddress -match "^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$"
}

# Import CSV file
$csvPath = Read-Host "Please enter the path to your CSV file (should contain columns: DisplayName,EmailAddress,FirstName,LastName)"
if (!(Test-Path $csvPath)) {
    Write-Host "CSV file not found at path: $csvPath" -ForegroundColor Red
    exit 1
}

try {
    $contacts = Import-Csv $csvPath
    $successCount = 0
    $errorCount = 0
    
    foreach ($contact in $contacts) {
        # Validate required fields
        if ([string]::IsNullOrWhiteSpace($contact.DisplayName) -or [string]::IsNullOrWhiteSpace($contact.EmailAddress)) {
            Write-Host "Skipping contact with missing required fields: $($contact.EmailAddress)" -ForegroundColor Yellow
            $errorCount++
            continue
        }

        # Validate email format
        if (!(Test-EmailAddress $contact.EmailAddress)) {
            Write-Host "Invalid email format for contact: $($contact.EmailAddress)" -ForegroundColor Yellow
            $errorCount++
            continue
        }

        try {
            # Check if contact already exists
            $existingContact = Get-MailContact -Identity $contact.EmailAddress -ErrorAction SilentlyContinue
            if ($existingContact) {
                Write-Host "Contact already exists: $($contact.EmailAddress)" -ForegroundColor Yellow
                $errorCount++
                continue
            }

            # Create new mail contact
            New-MailContact -Name $contact.DisplayName `
                          -ExternalEmailAddress $contact.EmailAddress `
                          -FirstName $contact.FirstName `
                          -LastName $contact.LastName
            
            Write-Host "Successfully created contact: $($contact.EmailAddress)" -ForegroundColor Green
            $successCount++
        }
        catch {
            Write-Host "Error creating contact $($contact.EmailAddress): $_" -ForegroundColor Red
            $errorCount++
        }
    }

    # Display summary
    Write-Host "`nImport Summary:" -ForegroundColor Cyan
    Write-Host "Successfully created: $successCount contacts" -ForegroundColor Green
    Write-Host "Errors/Skipped: $errorCount contacts" -ForegroundColor Yellow
}
catch {
    Write-Host "Error processing CSV file: $_" -ForegroundColor Red
}
finally {
    # Disconnect from Exchange Online
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Host "`nDisconnected from Exchange Online." -ForegroundColor Gray
}