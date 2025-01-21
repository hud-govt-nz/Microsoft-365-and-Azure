#Connect-ExchangeOnline -ShowBanner: $false

[array]$ReviewTags = Get-ComplianceTag | Where-Object {$_.id -eq "DA-5.3.4 Complaints and issues"} | Sort-Object Name
If (!($ReviewTags)) { Write-Host "No retention tags with manual disposition found - exiting"; break }

Write-Host ("Looking for Review Items for {0} retention tags: {1}" -f $ReviewTags.count, ($ReviewTags.Name -join ", "))

$Report = [System.Collections.Generic.List[Object]]::new()

# Initialize progress tracking variables
$totalItemsCount  = $ReviewTags.Count
$currentItemCount = 0

[array]$ItemsForReport = $Null
ForEach ($ReviewTag in $ReviewTags) {
    Write-Host ("Processing disposition items for the {0} label" -f $ReviewTag.Name)
    [array]$ItemDetails = $Null; [array]$ItemDetailsExport = $Null

    try {
        # Set the error action preference to silently continue
        $ErrorActionPreference = 'SilentlyContinue'

        # Check if there are review items for the tag
        $ReviewItems = Get-ReviewItems -TargetLabelId $ReviewTag.ImmutableId -IncludeHeaders $True -Disposed $false
        
        $ItemDetails += $ReviewItems.ExportItems
        # If more pages of data are available, fetch them and add to the Item details array
        While (![string]::IsNullOrEmpty($ReviewItems.PaginationCookie))
        {
            $ReviewItems  = Get-ReviewItems -TargetLabelId $ReviewTag.ImmutableId -IncludeHeaders $True -PagingCookie $ReviewItems.PaginationCookie
            $ItemDetails += $ReviewItems.ExportItems
        }
        # Convert data from CSV
        If ($ItemDetails) {
            [array]$ItemDetailsExport = $ItemDetails | ConvertFrom-Csv -Header $ReviewItems.Headers
            ForEach ($Item in $ItemDetailsExport) {
            # Sometimes the data doesn't include the label name, so we add the label name to be sure
            $Item | Add-Member -NotePropertyName Label -NotePropertyValue $ReviewTag.Name }
            $ItemsForReport += $ItemDetailsExport
        }

        # Reset the error action preference to default
        $ErrorActionPreference = 'Continue'
    } catch {
        # This catch block will now only execute for other unexpected errors
        Write-Host "An unexpected error occurred processing tag $($ReviewTag.Name): $_"
    }

    # Update the current item count and the progress bar
    $currentItemCount++
    $progress = @{
        Activity        = "Processing Review Tags"
        Status          = "Processing tag $($ReviewTag.Name)"
        PercentComplete = ($currentItemCount / $totalItemsCount) * 100
    }
    Write-Progress @progress
}

ForEach ($Record in $ItemsForReport) {
    $RecordCreationDate     = if ($Record.ItemCreationTime) { Get-Date($Record.ItemCreationTime) -format g } else { "Unknown" }
    $RecordLastModifiedDate = if ($Record.ItemLastModifiedTime) { Get-Date($Record.ItemLastModifiedTime) -format g } else { "Unknown" }
    $RecordDeletedDate      = if ($Record.DeletedDate) { Get-Date($Record.DeletedDate) -format g } else { "Unknown" }

    $DataLine  = [PSCustomObject] @{
        TimeStamp             = $RecordCreationDate
        Subject               = $Record.Subject
        LabelName             = $Record.LabelName
        LabelAppliedBy        = $Record.LabelAppliedBy
        RecordType            = $Record.RecordType
        ItemLastModifiedTime  = $RecordLastModifiedDate
        LastModifiedBy        = $Record.LastModifiedBy
        ReviewAction          = $Record.ReviewAction
        Comment               = $Record.Comment
        DeletedDate           = $RecordDeletedDate
        DeletedBy             = $Record.DeletedBy
        Author                = $Record.Author
        'Internet Message ID' = $Record.InternetMessageId
        Location              = $Record.Location
        LabelAppliedDate      = if ($Record.LabelAppliedDate) { Get-Date($Record.LabelAppliedDate) -format g } else { "Unknown" }
        ExpiryDate            = if ($Record.ExpiryDate) { Get-Date($Record.ExpiryDate) -format g } else { "Unknown" }
    }
    $Report.Add($DataLine)
}

# Optional: Export Report to a CSV file
$Report | Export-Excel -Path "C:\HUD\DispositionReport_Pending.xlsx" -WorksheetName "Report" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
