Clear-Host
Write-host "## Audit Log Search ##" -ForegroundColor Yellow

# Connect to Exchange Online
Connect-MgGraph -NoWelcome | Out-Null
$Login = (Get-MgContext).Account
Connect-ExchangeOnline -UserPrincipalName $Login -ShowBanner:$false
Connect-IPPSSession -UserPrincipalName $Login -ShowBanner:$false

# User input
$FormatUPN = Read-Host "Enter File or location"
Write-Host "Wildcard entries can be used, for example: Https://test.sharepoint.com/sites* or 'test file name here*'" -ForegroundColor Yellow

$StartDateInput = Read-Host "Enter start date NZT (format: dd/MM/yyyy)"
$EndDateInput = Read-Host "Enter end date NZT (format: dd/MM/yyyy)"

# Set timezone and date format
$nzTimeZone = [TimeZoneInfo]::FindSystemTimeZoneById("New Zealand Standard Time")
$dateFormat = "dd/MM/yyyy HH:mm:ss"

# Append "00:00:00" to the user-provided start and end dates
$StartDateInput += " 00:00:00"
$EndDateInput += " 00:00:00"

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
$ProcessedSharePoint =@()

# Function to process audit data and correct date time, and filter duplicates by Id and Workload
function Process-AuditData {
    param (
        [Parameter(Mandatory=$true)]
        [Object[]]$AuditData,
        [Parameter(Mandatory=$true)]
        [TimeZoneInfo]$TimeZone
    )

    $uniqueEntries = @{}
    $Data= @()

# Function to convert an array of objects to a string
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

    foreach ($jsonEntry in $AuditData) {
        # Check if the JSON entry is an array or a single object
        $entries = @($jsonEntry | ConvertFrom-Json)
        
        foreach ($entry in $entries) {
            # Transform to local date time
            $utcDateTime = [DateTime]::Parse($entry.CreationTime, [System.Globalization.CultureInfo]::InvariantCulture)
            $localDateTime = [TimeZoneInfo]::ConvertTimeFromUtc($utcDateTime, $TimeZone)
            $entry.CreationTime = $localDateTime

            # Loop through all properties of the entry
            foreach ($property in $entry.PSObject.Properties) {
                if ($property.Value -is [Array] -and $property.Value.Count -gt 0) {
                    # Convert array properties to a string
                    $property.Value = Convert-ArrayToString -Array $property.Value
                }
            }

            # Create a unique key based on Id and Workload
            $uniqueKey = "$($entry.Id)_$($entry.Workload)"

            # Guard clause: Check if an entry with the same key already exists
            if (!$uniqueEntries.ContainsKey($uniqueKey))  {
                $uniqueEntries[$uniqueKey] = $true
                
                $Data += $entry
            }
        }
    }
    
    return $Data
}
  
# Display the Out-GridView selection pane to let the user choose M365 capabilities for the report  
$selectedWorkloads = "SharePoint"
  
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
    $Search = Search-UnifiedAuditLog -objectIDs $FormatUPN -StartDate $FormatStartDate -EndDate $FormatEndDate -SessionId $sessionId -SessionCommand ReturnLargeSet -ResultSize 5000 -Formatted  
      
    # Collect and process Audit Data  
    if ($Search -and $Search.AuditData) {  
        $SearchResults = $Search.AuditData  
  
        # Process and filter results from interval  
        $IntervalResults = Process-AuditData -AuditData $SearchResults -TimeZone $nzTimeZone  
  
        # Filter processed results and add to respective final arrays  
        $ProcessedSharePoint += $IntervalResults | Where-Object { $_.Workload -eq "sharepoint"}  
    }  
  
    # Update Start and End Dates for the next interval  
    $FormatStartDate = $FormatEndDate  
    $FormatEndDate = $FormatEndDate.AddMinutes($intervalMinutes)  
}  
  
# Sort and remove duplicates from the final arrays  
$ProcessedSharePoint = $ProcessedSharePoint | Sort-Object Id -Unique  

Write-Progress -Activity "Processing SharePoint Audit Logs" -Status "Interval $currentInterval of $totalIntervals" -PercentComplete (($currentInterval / $totalIntervals) * 100) -Completed

# Export Process
$FileName = "SPO Activity Report"
[System.Windows.Forms.Application]::EnableVisualStyles()
$SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$SaveFileDialog.Filter = "XLSX Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
$SaveFileDialog.Title = "Save as"
$SaveFileDialog.FileName = $FileName

$SaveFileResult = $SaveFileDialog.ShowDialog()
if ($SaveFileResult -eq [System.Windows.Forms.DialogResult]::OK) {
    $SelectedFilePath = $SaveFileDialog.FileName

    # Export SharePoint, SharePointFileOperation, and SharePointSharingOperation results to separate worksheets
    if ($selectedWorkloads -contains "SharePoint") {  
        $tmp = $ProcessedSharePoint | Select-Object CreationTime, UserID, Workload, Operation, RecordType, SiteURL, SourceRelativeUrl, SourceFileName, SourceFileExtension, Platform, ApplicationDisplayName, ClientIP, Id |  Sort-Object creationtime 
        $tmp | Export-Excel -Path $SelectedFilePath -WorksheetName "SharePoint" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow  
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