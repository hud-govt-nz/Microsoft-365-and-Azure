Clear-Host
Write-host "## SharePoint Online Reporting: User Audit Log Search ##" -ForegroundColor Yellow

# Connect to Exchange Online
Connect-MgGraph -NoWelcome | Out-Null
$Login = (Get-MgContext).Account
Connect-ExchangeOnline -UserPrincipalName $Login -ShowBanner:$false
Connect-IPPSSession -UserPrincipalName $Login -ShowBanner:$false

# Configuration
$FormatUPN = Read-Host "Enter username (e.g., first.last@hud.govt.nz)"
$nzTimeZone = [TimeZoneInfo]::FindSystemTimeZoneById("New Zealand Standard Time")
$dateFormat = "dd/MM/yyyy HH:mm:ss"
$StartDateInput = Read-Host "Enter start date (format: dd/MM/yyyy)"

# Append "00:00:00" to the user-provided start date
$StartDateInput += " 00:00:00"

$EndDateInput = Read-Host "Enter end date (format: dd/MM/yyyy)"

# Append "00:00:00" to the user-provided end date
$EndDateInput += " 00:00:00"

$StartDate = [datetime]::ParseExact($StartDateInput, $dateFormat, [System.Globalization.CultureInfo]::InvariantCulture)
$EndDate = [datetime]::ParseExact($EndDateInput, $dateFormat, [System.Globalization.CultureInfo]::InvariantCulture)
$intervalMinutes = 60

# Initialize separate arrays for each record type
$SharePointResults = @()
$SharePointFileOpResults = @()
$SharePointSharingOpResults = @()

# Determine total number of intervals for progress calculation
$totalIntervals = [math]::Ceiling(($EndDate - $StartDate).TotalMinutes / $intervalMinutes)
$currentInterval = 0

# Function to process audit data and filter duplicates by Id and Workload
function Process-AuditData {
    param (
        [Parameter(Mandatory=$true)]
        [Object[]]$AuditData,
        [Parameter(Mandatory=$true)]
        [TimeZoneInfo]$TimeZone
    )

    $uniqueEntries = @{}
    $filteredData = @()

    foreach ($entry in $AuditData | ConvertFrom-Json) {
        # Guard clause: Check if the Workload is equal to SharePoint
        if ($entry.Workload -eq "SharePoint") {
            $utcDateTime = [DateTime]::Parse($entry.CreationTime, [System.Globalization.CultureInfo]::InvariantCulture)
            $localDateTime = [TimeZoneInfo]::ConvertTimeFromUtc($utcDateTime, $TimeZone)
            $entry.CreationTime = $localDateTime

            # Guard clause: Check if an entry with the same Id already exists
            if (!$uniqueEntries.ContainsKey($entry.Id)) {
                $uniqueEntries[$entry.Id] = $true

                # Select properties based on RecordType
                if ($entry.RecordType -eq "SharePoint") {
                    $filteredData += $entry | Select-Object CreationTime, UserId, Operation, ObjectId, ClientIP, BrowserName, BrowserVersion, AuthenticationType, EventSource, IsManagedDevice, ItemType, Platform, RecordType, Workload, EventData, SearchQueryText
                } elseif ($entry.RecordType -eq "SharePointFileOperation") {
                    $filteredData += $entry | Select-Object CreationTime, UserId, Operation, SiteUrl, SourceRelativeUrl, SourceFileName, SourceFileExtension, ClientIP, AuthenticationType, EventSource, IsManagedDevice, ItemType, Platform, RecordType, Workload, ObjectId
                } elseif ($entry.RecordType -eq "SharePointSharingOperation") {
                    $filteredData += $entry | Select-Object CreationTime, UserId, Operation, TargetUserOrGroupType, TargetUserOrGroupName, SiteUrl, ObjectId, ClientIP, BrowserName, BrowserVersion, AuthenticationType, EventSource, IsManagedDevice, ItemType, Platform, RecordType, Workload
                }
            }
        }
    }

    return $filteredData
}


# Main Loop for SharePoint Operations
$FormatStartDate = $StartDate
$FormatEndDate = $StartDate.AddMinutes($intervalMinutes)
while ($FormatStartDate -lt $EndDate) {
    $currentInterval++
    Write-Progress -Activity "Processing SharePoint Audit Logs" -Status "Interval $currentInterval of $totalIntervals" -PercentComplete (($currentInterval / $totalIntervals) * 100)
    Write-Host "Processing interval $currentInterval of $totalIntervals..."

    $sessionId = [Guid]::NewGuid().ToString()

    # Collecting results for SharePoint
    $ShareOps = Search-UnifiedAuditLog -UserIds $FormatUPN -StartDate $FormatStartDate -EndDate $FormatEndDate -SessionId $sessionId -RecordType SharePoint -SessionCommand ReturnLargeSet -ResultSize 5000 -Formatted 
    if ($ShareOps -and $ShareOps.AuditData) {
        $ShareOpsResults = $ShareOps.AuditData
        $SharePointResults += Process-AuditData -AuditData $ShareOpsResults -TimeZone $nzTimeZone
    }

    # Collecting results for SharePointFileOperation
    $ShareFileOps = Search-UnifiedAuditLog -UserIds $FormatUPN -StartDate $FormatStartDate -EndDate $FormatEndDate -SessionId $sessionId -RecordType SharePointFileOperation -SessionCommand ReturnLargeSet -ResultSize 5000 -Formatted 
    if ($ShareFileOps -and $ShareFileOps.AuditData) {
        $ShareFileOpsResults = $ShareFileOps.AuditData
        $SharePointFileOpResults += Process-AuditData -AuditData $ShareFileOpsResults -TimeZone $nzTimeZone
    }

    # Collecting results for SharePointSharingOperation
    $ShareSharingOps = Search-UnifiedAuditLog -UserIds $FormatUPN -StartDate $FormatStartDate -EndDate $FormatEndDate -SessionId $sessionId -RecordType SharePointSharingOperation -SessionCommand ReturnLargeSet -ResultSize 5000 -Formatted 
    if ($ShareSharingOps -and $ShareSharingOps.AuditData) {
        $ShareSharingOpsResults = $ShareSharingOps.AuditData
        $SharePointSharingOpResults += Process-AuditData -AuditData $ShareSharingOpsResults -TimeZone $nzTimeZone
    }

    # Update Start and End Dates for the next interval
    $FormatStartDate = $FormatEndDate
    $FormatEndDate = $FormatEndDate.AddMinutes($intervalMinutes)

}


Write-Progress -Activity "Processing SharePoint Audit Logs" -Status "Interval $currentInterval of $totalIntervals" -PercentComplete (($currentInterval / $totalIntervals) * 100) -Completed

# Export Process
$FileName = "SharePoint Search"
[System.Windows.Forms.Application]::EnableVisualStyles()
$SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$SaveFileDialog.Filter = "XLSX Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
$SaveFileDialog.Title = "Save as"
$SaveFileDialog.FileName = $FileName

$SaveFileResult = $SaveFileDialog.ShowDialog()
if ($SaveFileResult -eq [System.Windows.Forms.DialogResult]::OK) {
    $SelectedFilePath = $SaveFileDialog.FileName

    # Export SharePoint, SharePointFileOperation, and SharePointSharingOperation results to separate worksheets
    $SharePointResults | Export-Excel -Path $SelectedFilePath -WorksheetName "SharePoint" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    $SharePointFileOpResults | Export-Excel -Path $SelectedFilePath -WorksheetName "SharePointFileOperation" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    $SharePointSharingOpResults | Export-Excel -Path $SelectedFilePath -WorksheetName "SharePointSharingOperation" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow

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

[System.GC]::Collect()
