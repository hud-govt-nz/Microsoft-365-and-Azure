#Parameters
$SiteURL = "https://mhud.sharepoint.com/sites/infomgmt"
$ReportOutput = "C:\HUD\06_Reporting\SPO\FileSizeRpt.csv"

#Function to convert bytes to GB
function Convert-ToGB {
    param([double]$bytes)
    return [math]::Round(($bytes / 1GB), 2)
}
   
#Connect to SharePoint Online site
#Connect-PnPOnline $SiteURL -Interactive
$env:PNPPOWERSHELL_UPDATECHECK = "Off"
Connect-PnPOnline `
-Url $SiteURL `
-ClientId $env:DigitalSupportAppID `
-Tenant 'mhud.onmicrosoft.com' `
-Thumbprint $env:DigitalSupportCertificateThumbprint
 
#Array to store results
$Results = @()

#Get all document libraries from the site
$DocLibraries = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 }
Write-host "Found $($DocLibraries.Count) document libraries"

foreach($Library in $DocLibraries) {
    try {
        Write-host "`nProcessing library: $($Library.Title)" -ForegroundColor Cyan
        
        #Get all Items from the document library
        $ListItems = Get-PnPListItem -List $Library.Title -PageSize 500 | Where-Object {$_.FileSystemObjectType -eq "File"}
        $TotalItems = $ListItems.Count
        Write-host "Found $TotalItems items in $($Library.Title)"
        
        $ItemCounter = 0 
        $ErrorCount = 0
        #Iterate through each item
        Foreach ($Item in $ListItems)
        {
            try {
                $FileSizeBytes = [double]$Item.FieldValues.File_x0020_Size
                $TotalSizeBytes = [double]$Item.FieldValues.SMTotalSize.LookupId
                
                $Results += New-Object PSObject -Property ([ordered]@{
                    SiteUrl            = $SiteURL
                    Library           = $Library.Title
                    ID               = $Item.Id
                    UniqueId         = $Item.FieldValues.GUID
                    ParentFolderUniqueId = $Item.FieldValues.ParentUniqueId
                    Version          = $Item.FieldValues._UIVersionString
                    FolderPath       = $Item.FieldValues.FileDirRef
                    Title            = $Item.FieldValues.FileLeafRef
                    ServerRelativePath = $Item.FieldValues.FileRef
                    RetentionLabel   = $Item.FieldValues._ComplianceTag
                    SensitivityLabel = $Item.FieldValues._DisplayName
                    CreatedBy       = $Item["Author"].LookupValue
                    Created         = $Item["Created"]
                    LastModified     = $Item["Last_x0020_Modified"]
                    ModifiedBy      = $Item["Editor"].LookupValue
                    FileSizeGB       = Convert-ToGB $FileSizeBytes
                    TotalFileSizeGB  = Convert-ToGB $TotalSizeBytes
                })
                $ItemCounter++
            }
            catch {
                $ErrorCount++
                Write-Warning "Error processing item $($Item.FieldValues.FileLeafRef): $_"
            }
            Write-Progress -PercentComplete ($ItemCounter / $TotalItems * 100) -Activity "Processing Items $ItemCounter of $TotalItems" -Status "Getting data from Item '$($Item.FieldValues.FileLeafRef)'"
        }
        Write-Host "Completed processing $($Library.Title). Processed: $ItemCounter, Errors: $ErrorCount" -ForegroundColor Green
    }
    catch {
        Write-Error "Error processing library $($Library.Title): $_"
    }
}
  
#Export the results to CSV, sorted by size
$Results | Sort-Object FileSizeGB -Descending | Export-Csv -Path $ReportOutput -NoTypeInformation
Write-host "File Size Report Exported to CSV Successfully!"

#Display summary of total size by library
$Results | Group-Object Library | Select-Object Name, @{N='TotalSizeGB';E={($_.Group | Measure-Object FileSizeGB -Sum).Sum}} | 
    Sort-Object TotalSizeGB -Descending | Format-Table -AutoSize