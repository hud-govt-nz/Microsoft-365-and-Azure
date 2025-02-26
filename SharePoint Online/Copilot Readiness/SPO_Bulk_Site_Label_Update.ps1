# Start transcription
$logDirectory = "C:\HUD\06_Reporting\SPO\Logs"
if (-not (Test-Path $logDirectory)) {
    New-Item -ItemType Directory -Path $logDirectory | Out-Null
}
$transcriptPath = Join-Path $logDirectory "SPO_Bulk_Site_Label_Update_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
Start-Transcript -Path $transcriptPath

# Check PowerShell version and disable PnP update check
if ($PSVersionTable.PSVersion -le [Version]"7.0") {
    Write-Host "PowerShell version 7.0 or higher is required for this script." -ForegroundColor Red
    Stop-Transcript
    exit 1
}

# Disable PnP PowerShell update check
$env:PNPPOWERSHELL_UPDATECHECK = "Off"

# Add Windows Forms assembly for file picker
Add-Type -AssemblyName System.Windows.Forms

# Create and configure file picker dialog
$filePicker = New-Object System.Windows.Forms.OpenFileDialog
$filePicker.Title = "Select CSV file containing site label updates"
$filePicker.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
$filePicker.FilterIndex = 1

# Show file picker dialog
if ($filePicker.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $csvPath = $filePicker.FileName
} else {
    Write-Host "No file selected. Exiting script." -ForegroundColor Yellow
    Stop-Transcript
    exit
}

# Initialize counters and arrays
$totalItems = 0
$successful = 0
$failed = 0
$skipped = 0
$errors = @()
$libraryUpdates = @{}
$libraryUpdateSuccessful = 0
$libraryUpdateFailed = 0

# Validate CSV file exists
if (-not (Test-Path $csvPath)) {
    Write-Error "CSV file not found at path: $csvPath"
    Stop-Transcript
    exit 1
}

# Import and validate CSV
try {
    $CSV = Import-Csv -Path $csvPath
    
    # Validate CSV is not empty
    if ($null -eq $CSV -or $CSV.Count -eq 0) {
        Write-Error "CSV file is empty"
        Stop-Transcript
        exit 1
    }

    $totalItems = $CSV.Count

    # Validate required columns exist
    $requiredColumns = @('SiteUrl', 'FolderPath', 'List', 'ItemID', 'FileName', 'RetentionLabel', 'NewRetentionLabel')
    $missingColumns = $requiredColumns | Where-Object { $_ -notin $CSV[0].PSObject.Properties.Name }
    
    if ($missingColumns) {
        Write-Error "CSV is missing required columns: $($missingColumns -join ', ')"
        Stop-Transcript
        exit 1
    }

    # Get unique site URLs and group items by site
    $uniqueSites = $CSV | Select-Object -ExpandProperty SiteUrl -Unique
    $siteGroups = $CSV | Group-Object -Property SiteUrl

    # Success message with item count and site information
    Write-Host "Successfully imported CSV with $($CSV.Count) items to process across $($uniqueSites.Count) sites" -ForegroundColor Green
    Write-Host "`nSites to be processed:" -ForegroundColor Cyan
    $siteGroups | ForEach-Object {
        Write-Host "- $($_.Name) ($($_.Count) items)" -ForegroundColor Gray
    }

    # Prompt for confirmation
    $confirmation = Read-Host "Do you want to proceed with the updates? (Y/N)"
    if ($confirmation -ne 'Y') {
        Write-Host "Operation cancelled by user" -ForegroundColor Yellow
        Stop-Transcript
        exit 0
    }

    # Process each site group
    foreach ($siteGroup in $siteGroups) {
        $currentSite = $siteGroup.Name
        $siteItems = $siteGroup.Group
        
        Write-Host "`n=== Processing Site: $currentSite ===" -ForegroundColor Cyan
        Write-Host "Items to process: $($siteItems.Count)" -ForegroundColor Gray
        
        try {
            # Connect to the current site
            Write-Host "Connecting to site..." -ForegroundColor Yellow
            Connect-PnPOnline `
                -Url $currentSite `
                -ClientId $env:DigitalSupportAppID `
                -Tenant 'mhud.onmicrosoft.com' `
                -Thumbprint $env:DigitalSupportCertificateThumbprint
            Write-Host "Connected successfully to $currentSite" -ForegroundColor Green

            # Process each item in the current site
            foreach ($item in $siteItems) {
                # Display progress
                Write-Progress -Activity "Updating Retention Labels" -Status "Processing item $($siteItems.IndexOf($item) + 1) of $($siteItems.Count)" -PercentComplete ((($siteItems.IndexOf($item) + 1) / $siteItems.Count) * 100)
                
                # Skip items with "HUD Record" label
                if ($item.RetentionLabel -eq "HUD Record") {
                    Write-Host "`nSkipping file: $($item.FileName) - Has 'HUD Record' label" -ForegroundColor Yellow
                    $script:skipped++
                    continue
                }
                
                try {
                    Write-Host "`nProcessing file: $($item.FileName)" -ForegroundColor Cyan
                    Write-Host "List: $($item.List), ItemID: $($item.ItemID)" -ForegroundColor Gray
                    Write-Host "Current Label: $($item.RetentionLabel)" -ForegroundColor Gray
                    Write-Host "New Label: $($item.NewRetentionLabel)" -ForegroundColor Gray
                    
                    # Update the retention label
                    Set-PnPListItem -List $item.List -Identity $item.ItemID -Label $item.NewRetentionLabel -ErrorAction Stop
                    
                    Write-Host "✓ Successfully updated retention label" -ForegroundColor Green
                    $script:successful++

                    # Track library for update if not already processed
                    $libraryKey = "$($currentSite)|$($item.List)"
                    if (-not $libraryUpdates.ContainsKey($libraryKey)) {
                        $libraryUpdates[$libraryKey] = @{
                            SiteUrl = $currentSite
                            ListName = $item.List
                            NewLabel = $item.NewRetentionLabel
                        }
                    }
                }
                catch {
                    Write-Host "✕ Failed to update retention label: $_" -ForegroundColor Red
                    $script:failed++
                    $script:errors += [PSCustomObject]@{
                        FileName = $item.FileName
                        Path = "$($item.SiteUrl)$($item.ServerRelativePath)"
                        Error = $_.Exception.Message
                    }
                }
            }

            # Update library retention labels for this site
            Write-Host "`n=== Updating Library Retention Labels ===" -ForegroundColor Cyan
            $siteLibraries = $libraryUpdates.Values | Where-Object { $_.SiteUrl -eq $currentSite }
            
            foreach ($library in $siteLibraries) {
                try {
                    Write-Host "Updating library '$($library.ListName)' with retention label '$($library.NewLabel)'..." -ForegroundColor Yellow
                    
                    # Get the list
                    $list = Get-PnPList -Identity $library.ListName -ErrorAction Stop
                    
                    # Update the list retention label using Set-PnPRetentionLabel
                    Set-PnPRetentionLabel -List $list -Label $library.NewLabel -SyncToItems $false -ErrorAction Stop
                    
                    Write-Host "✓ Successfully updated library retention label" -ForegroundColor Green
                    $script:libraryUpdateSuccessful++
                }
                catch {
                    Write-Host "✕ Failed to update library retention label: $_" -ForegroundColor Red
                    $script:libraryUpdateFailed++
                    $script:errors += [PSCustomObject]@{
                        FileName = "Library: $($library.ListName)"
                        Path = $library.SiteUrl
                        Error = "Failed to update library retention label: $_"
                    }
                }
            }
        }
        catch {
            Write-Host "Failed to connect to site $currentSite. Skipping all items for this site." -ForegroundColor Red
            Write-Host "Error: $_" -ForegroundColor Red
            $script:failed += $siteItems.Count
            $script:errors += $siteItems | ForEach-Object {
                [PSCustomObject]@{
                    FileName = $_.FileName
                    Path = "$($_.SiteUrl)$($_.ServerRelativePath)"
                    Error = "Failed to connect to site: $_"
                }
            }
        }
    }

    # Display summary
    Write-Host "`n=== Summary ===" -ForegroundColor Yellow
    Write-Host "Total items processed: $totalItems" -ForegroundColor White
    Write-Host "Successful updates: $successful" -ForegroundColor Green
    Write-Host "Failed updates: $failed" -ForegroundColor Red
    Write-Host "Skipped items (HUD Record): $skipped" -ForegroundColor Yellow
    Write-Host "`nLibrary Updates:" -ForegroundColor Yellow
    Write-Host "Successful library updates: $libraryUpdateSuccessful" -ForegroundColor Green
    Write-Host "Failed library updates: $libraryUpdateFailed" -ForegroundColor Red

    # Display errors if any
    if ($errors.Count -gt 0) {
        Write-Host "`n=== Failed Items ===" -ForegroundColor Red
        $errors | Format-Table -AutoSize
    }

    # Remove progress bar
    Write-Progress -Activity "Updating Retention Labels" -Completed

    # Stop transcript
    Write-Host "`nLog file saved to: $transcriptPath" -ForegroundColor Cyan
    Stop-Transcript
}
catch {
    Write-Error "An error occurred during script execution: $_"
    Stop-Transcript
    exit 1
}
