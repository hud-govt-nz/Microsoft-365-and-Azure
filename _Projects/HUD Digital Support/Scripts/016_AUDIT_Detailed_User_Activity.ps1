Clear-Host
Write-host "## User Audit Log Search ##" -ForegroundColor Yellow

# Connect to Exchange Online
Connect-MgGraph -NoWelcome | Out-Null
$Login = (Get-MgContext).Account
Connect-ExchangeOnline -UserPrincipalName $Login -ShowBanner:$false
Connect-IPPSSession -UserPrincipalName $Login -ShowBanner:$false

# User input
$FormatUPN = Read-Host "Enter username (e.g., first.last@hud.govt.nz)"
$StartDateInput = Read-Host "Enter start date (format: dd/MM/yyyy)"
$EndDateInput = Read-Host "Enter end date (format: dd/MM/yyyy)"

# Set timezone and date format
$nzTimeZone = [TimeZoneInfo]::FindSystemTimeZoneById("New Zealand Standard Time")
$dateFormat = "dd/MM/yyyy HH:mm:ss"

# Append "00:00:00" to the user-provided start and end dates
$StartDateInput += " 00:00:00"
$EndDateInput += " 23:59:59"

# Parse date time input to an object 
$StartDate = [datetime]::ParseExact($StartDateInput, $dateFormat, [System.Globalization.CultureInfo]::InvariantCulture)
$EndDate = [datetime]::ParseExact($EndDateInput, $dateFormat, [System.Globalization.CultureInfo]::InvariantCulture)

# Set interval count to 1 hour
$intervalMinutes = 60

# Determine total number of intervals for progress calculation
$totalIntervals = [math]::Ceiling(($EndDate - $StartDate).TotalMinutes / $intervalMinutes)
$currentInterval = 0
$FormatStartDate = $StartDate
$FormatEndDate = $StartDate.AddMinutes($intervalMinutes)

# Arrays
$ProcessedEntra =@()
$ProcessedIntune =@()
$ProcessedTeams =@()
$ProcessedSharePoint =@()
$ProcessedOneDrive =@()
$ProcessedExchange =@()
$ProcessedOther =@()

# Function to process audit data and correct date time, and filter duplicates by Id and Workload
function Convert-ArrayToString {
    param (
        [Parameter(Mandatory=$true)]
        [Object[]]$Array
    )
    
    $stringArray = foreach ($obj in $Array) {
        $properties = $obj | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
        $propertyStrings = foreach ($property in $properties) {
            $value = $obj.$property
            if ($null -ne $value) {
                "${property}: $value"
            }
        }
        $propertyStrings -join ", "
    }
    return $stringArray -join "; "
}
function Convert-AuditData {
    param (
        [Parameter(Mandatory=$true)]
        [Object[]]$AuditData,
        [Parameter(Mandatory=$true)]
        [TimeZoneInfo]$TimeZone
    )

    $uniqueEntries = @{}
    $Data= @()

    foreach ($jsonEntry in $AuditData) {
        $entries = @($jsonEntry | ConvertFrom-Json)
        
        foreach ($entry in $entries) {
            $utcDateTime = [DateTime]::Parse($entry.CreationTime, [System.Globalization.CultureInfo]::InvariantCulture)
            $localDateTime = [TimeZoneInfo]::ConvertTimeFromUtc($utcDateTime, $TimeZone)
            $entry.CreationTime = $localDateTime

            foreach ($property in $entry.PSObject.Properties) {
                if ($property.Value -is [Array] -and $property.Value.Count -gt 0) {
                    $property.Value = Convert-ArrayToString -Array $property.Value
                }
            }

            $uniqueKey = "$($entry.Id)_$($entry.Workload)"
            if (!$uniqueEntries.ContainsKey($uniqueKey)) {
                $uniqueEntries[$uniqueKey] = $true
                $Data += $entry
            }
        }
    }
    return $Data
}

# Define available workload options  
$workloadOptions = @("EntraID (AAD)", "Intune", "Teams", "SharePoint", "OneDrive", "Exchange", "Other")  

# Display the Out-GridView selection pane to let the user choose M365 capabilities for the report  
$selectedWorkloads = $workloadOptions | Out-GridView -Title 'Select M365 capabilities to include in the report' -OutputMode Multiple  

# Check if any options were selected, exit the script if none were chosen  
if ($selectedWorkloads.Count -eq 0) {  
    Write-Host "No options selected. Exiting." -ForegroundColor Yellow  
    exit  
} 

# Main Loop  
while ($FormatStartDate -lt $EndDate) {  
    $currentInterval++  
    Write-Progress -Activity "Processing Audit Logs" -Status "Interval $currentInterval of $totalIntervals" -PercentComplete (($currentInterval / $totalIntervals) * 100)  
    Write-Host "Processing interval $currentInterval of $totalIntervals..."  

    # Generate individual GUID per loop  
    $sessionId = [Guid]::NewGuid().ToString()  

    # Collecting results  
    $Search = Search-UnifiedAuditLog -UserIds $FormatUPN -StartDate $FormatStartDate -EndDate $FormatEndDate -SessionId $sessionId -SessionCommand ReturnLargeSet -ResultSize 5000 -Formatted  

    # Collect and process Audit Data  
    if ($Search -and $Search.AuditData) {  
        $SearchResults = $Search.AuditData  

        # Process and filter results from interval  
        $IntervalResults = Convert-AuditData -AuditData $SearchResults -TimeZone $nzTimeZone  

        # Filter processed results and add to respective final arrays  
        $ProcessedEntra += $IntervalResults | Where-Object { $_.Workload -eq "azureactivedirectory" }  
        $ProcessedIntune += $IntervalResults | Where-Object { $_.Workload -eq "endpoint" }  
        $ProcessedTeams += $IntervalResults | Where-Object { $_.Workload -eq "microsoftteams" }  
        $ProcessedSharePoint += $IntervalResults | Where-Object { $_.Workload -eq "sharepoint" }  
        $ProcessedOneDrive += $IntervalResults | Where-Object { $_.Workload -eq "onedrive" }  
        $ProcessedExchange += $IntervalResults | Where-Object { $_.Workload -eq "exchange" }  
        $ProcessedOther += $IntervalResults | Where-Object { $_.Workload -ne "azureactivedirectory" -and $_.Workload -ne "endpoint" -and $_.Workload -ne "microsoftteams" -and $_.Workload -ne "sharepoint" -and $_.Workload -ne "onedrive" -and $_.Workload -ne "exchange" }  
    }  

    # Update Start and End Dates for the next interval  
    $FormatStartDate = $FormatEndDate  
    $FormatEndDate = $FormatEndDate.AddMinutes($intervalMinutes)  
}  

# Sort and remove duplicates from the final arrays  
$ProcessedEntra = $ProcessedEntra | Sort-Object Id -Unique  
$ProcessedIntune = $ProcessedIntune | Sort-Object Id -Unique  
$ProcessedTeams = $ProcessedTeams | Sort-Object Id -Unique  
$ProcessedSharePoint = $ProcessedSharePoint | Sort-Object Id -Unique  
$ProcessedOneDrive = $ProcessedOneDrive | Sort-Object Id -Unique  
$ProcessedExchange = $ProcessedExchange | Sort-Object Id -Unique  
$ProcessedOther = $ProcessedOther | Sort-Object Id -Unique

Write-Progress -Activity "Processing SharePoint Audit Logs" -Status "Interval $currentInterval of $totalIntervals" -PercentComplete (($currentInterval / $totalIntervals) * 100) -Completed

# Export Process
$FileName = "User Audit Search"
[System.Windows.Forms.Application]::EnableVisualStyles()
$SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$SaveFileDialog.Filter = "XLSX Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
$SaveFileDialog.Title = "Save as"
$SaveFileDialog.FileName = $FileName

$SaveFileResult = $SaveFileDialog.ShowDialog()
if ($SaveFileResult -eq [System.Windows.Forms.DialogResult]::OK) {
    $SelectedFilePath = $SaveFileDialog.FileName

    # Export SharePoint, SharePointFileOperation, and SharePointSharingOperation results to separate worksheets
    if ($selectedWorkloads -contains "EntraID (AAD)") {  
        $ProcessedEntra | Export-Excel -Path $SelectedFilePath -WorksheetName "EntraID (AAD)" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow  
    } 
    if ($selectedWorkloads -contains "Intune") {  
        $ProcessedIntune | Export-Excel -Path $SelectedFilePath -WorksheetName "Intune" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow  
    } 
    if ($selectedWorkloads -contains "Teams") {  
        $ProcessedTeams | Export-Excel -Path $SelectedFilePath -WorksheetName "Teams" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow  
    } 
    if ($selectedWorkloads -contains "SharePoint") {  
        $ProcessedSharePoint | Export-Excel -Path $SelectedFilePath -WorksheetName "SharePoint" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow  
    } 
    if ($selectedWorkloads -contains "OneDrive") {  
        $ProcessedOneDrive | Export-Excel -Path $SelectedFilePath -WorksheetName "OneDrive" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow  
    }
    if ($selectedWorkloads -contains "Exchange") {  
        $ProcessedExchange | Export-Excel -Path $SelectedFilePath -WorksheetName "Exchange" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow  
    } 
    if ($selectedWorkloads -contains "Other") {  
        $ProcessedOther | Export-Excel -Path $SelectedFilePath -WorksheetName "Other" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow  
    } 
    
    $excelPackage = Open-ExcelPackage -Path $SelectedFilePath

    # Loop through each worksheet in the workbook
    foreach ($worksheet in $excelPackage.Workbook.Worksheets) {

        # Assuming headers are in row 1 and you start from row 2
        $startRow = 2
        $endRow = $worksheet.Dimension.End.Row
        $startColumn = 1
        $endColumn = $worksheet.Dimension.End.Column

        # Check if the worksheet has any data
        if ($endRow -gt 0 -and $endColumn -gt 0) {
            # Set horizontal alignment to left and autosize columns
            for ($col = $startColumn; $col -le $endColumn; $col++) {
                for ($row = $startRow; $row -le $endRow; $row++) {
                    $worksheet.Cells[$row, $col].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
                    }
                $worksheet.Column($col).AutoFit() # Autosize this column
                }
            }
        }

    # Save and close the Excel package
    $excelPackage.Save()
    Close-ExcelPackage $excelPackage -Show

    Write-Host "The report $FileName has been saved in $($SelectedFilePath)" -ForegroundColor Green

} else {
    Write-Host "Save operation canceled." -ForegroundColor Yellow
}