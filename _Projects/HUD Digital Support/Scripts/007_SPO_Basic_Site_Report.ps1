<#
.SYNOPSIS
    SharePoint Online Reporting: Basic Site Report

.DESCRIPTION
    This script connects to a SharePoint Online environment and generates a report on site usage.
    The report includes details such as title, URL, owner, template name, status, locale ID, last modified date, sharing capability, storage allocated and used, and warning level for storage.

.AUTHOR
    Ashley Forde

.VERSION
    2.0

.DATE
    25.9.24

.EXAMPLE
    .\BasicSiteReport.ps1

    This command will run the script and generate a report which will then be displayed in a grid view. 
    The user can then save the report to an Excel file via a Save dialog.

.NOTES
    This script requires the SharePoint PnP PowerShell module to be installed and available.
    It is intended to be used by SharePoint administrators for reporting purposes.



#>

Clear-Host
Write-host "## SharePoint Online: Basic Site Report ##" -ForegroundColor Yellow


#Requires -Modules PNP.Powershell
# Connect to PnP PowerShell
$SiteURL = "https://mhud-admin.sharepoint.com"

try {
    Write-Host "Connecting to PnP PowerShell..."
    Connect-PnPOnline -Url $SiteURL `
                      -ClientId $env:DigitalSupportAppID `
                      -Tenant 'mhud.onmicrosoft.com' `
                      -Thumbprint $env:DigitalSupportCertificateThumbprint
    Write-Host "Connected" -ForegroundColor Green
    } catch {
	    Write-Host "Error connecting. Please check your credentials and network connection." -ForegroundColor Red
	    exit 1
        }

# Define the mapping from the key to the descriptive text
$TemplateMappings = @{
    'APPCATALOG#0'               = 'App Catalog Site'
    'BDR#0'                      = 'Document Center'
    'BICenterSite#0'             = 'Business Intelligence Center'
    'BLANKINTERNET#0'            = 'Publishing Site'
    'BLANKINTERNETCONTAINER#0'   = 'Publishing Portal'
    'COMMUNITY#0'                = 'Community Site'
    'COMMUNITYPORTAL#0'          = 'Community Portal'
    'DEV#0'                      = 'Developer Site'
    'EHS#1'                      = 'Team Site - SharePoint Online configuration'
    'ENTERWIKI#0'                = 'Enterprise Wiki'
    'GROUP#0'                    = 'Team site'
    'OFFILE#1'                   = 'Records Center'
    'POINTPUBLISHINGHUB#0'       = 'PointPublishing Hub'
    'POINTPUBLISHINGPERSONAL#0'  = 'Personal blog'
    'POINTPUBLISHINGTOPIC#0'     = 'PointPublishing Topic'
    'PRODUCTCATALOG#0'           = 'Product Catalog'
    'PROJECTSITE#0'              = 'Project Site'
    'PWA#0'                      = 'Project Web App Site'
    'RedirectSite#0'             = 'Redirect Site'
    'REVIEWCTR#0'                = 'Review Center'
    'SITEPAGEPUBLISHING#0'       = 'Communication site'
    'SPSMSITEHOST#0'             = 'My Site Host'
    'SRCHCEN#0'                  = 'Enterprise Search Center'
    'SRCHCENTERLITE#0'           = 'Basic Search Center'
    'STS#0'                      = 'Team site (classic experience)'
    'STS#3'                      = 'Team site (no Microsoft 365 group)'
    'TEAMCHANNEL#0'              = 'Team channel'
    'TEAMCHANNEL#1'              = 'Team channel'
    'visprus#0'                  = 'Visio Process Repository'
}

# Obtain Site Data
$AllSites = Get-PnPTenantSite -Detailed | Sort-Object Title

# Count for progress bar
$TotalCount = $AllSites.count
$progress = 0

Write-Host ""
Write-Host "Total number of sites is $($TotalCount)"

# Collection of Site Values
$Sites =@()

$AllSites | ForEach-Object {
    $Progress++
    Write-Progress -Activity "Fetching Data from SharePoint" -Status "$Progress of $TotalCount" -PercentComplete ($Progress/$TotalCount*100)

    $item = $_ | Select-Object Title, Url, Owner,
    @{Name='Template'; Expression={$TemplateMappings[$_.Template]}},
    Status, LastContentModifiedDate,SharingCapability,
    @{Name="Allocated Storage (TB)"; Expression={[math]::Round($_.StorageQuota / 1MB, 2)}},
    @{Name="Used Storage (MB)"; Expression={$_.StorageUsageCurrent}},
    @{Name="Storage Warning Level (TB)"; Expression={[math]::Round($_.StorageQuotaWarningLevel / 1MB, 2)}} 

    $Sites += $item
    }

Write-Host "Open Save Dialog"

$Date = Get-Date -Format "dd.MM.yyyy h.mm tt"
$FileName = "All SPO Sites Export"  
    
# Add assembly and import namespace  
Add-Type -AssemblyName System.Windows.Forms  
[System.Windows.Forms.Application]::EnableVisualStyles()  
$SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog  
      
# Configure the SaveFileDialog  
$SaveFileDialog.Filter = "XLSX Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"  
$SaveFileDialog.Title = "Save as"  
$SaveFileDialog.FileName = $FileName  
      
# Show the SaveFileDialog and get the selected file path  
$SaveFileResult = $SaveFileDialog.ShowDialog()  
if ($SaveFileResult -eq [System.Windows.Forms.DialogResult]::OK) {  
    $SelectedFilePath = $SaveFileDialog.FileName
    $Sites | Export-Excel $SelectedFilePath -AutoSize -AutoFilter -WorksheetName $Date -FreezeTopRow -BoldTopRow
        
    $excelPackage = Open-ExcelPackage -Path $SelectedFilePath
    $worksheet = $excelPackage.Workbook.Worksheets["$Date"]

    # Assuming headers are in row 1 and you start from row 2
    $startRow = 2
    $endRow = $worksheet.Dimension.End.Row
    $startColumn = 1
    $endColumn = $worksheet.Dimension.End.Column

    # Set horizontal alignment to left for all cells in the used range
    for ($col = $startColumn; $col -le $endColumn; $col++) {
        for ($row = $startRow; $row -le $endRow; $row++) {
            $worksheet.Cells[$row, $col].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
        }
    }

    # Now, continue with your hyperlink setting and other operations
    $rowIndex = 2
    foreach ($item in $Sites) {
        # Trim the hyperlink
        $hyperlink = $item.URL.Trim()

        # Set the hyperlink directly
        $cell = $worksheet.Cells[$rowIndex, 2]
        $cell.Hyperlink = New-Object -TypeName OfficeOpenXml.ExcelHyperLink -ArgumentList $hyperlink
        $cell.Value = "Link"  # Or any other display text you want for the hyperlink

        # Set the style for the hyperlink
        $cell.Style.Font.UnderLine = $true
        $cell.Style.Font.Color.SetColor([System.Drawing.Color]::Blue)
    
        # Increment the row index for the next item
        $rowIndex++
    }
    
        # Autosize columns if needed
        foreach ($column in $worksheet.Dimension.Start.Column..$worksheet.Dimension.End.Column) {
            $worksheet.Column($column).AutoFit()
        }
    
        # Save and close the Excel package
        $excelPackage.Save()
        Close-ExcelPackage $excelPackage -Show
    
        Write-Host "The report $FileName has been saved in $($SelectedFilePath)" -ForegroundColor Green  
        } else {  
            Write-Host "Save cancelled" -ForegroundColor Yellow  
            }  
    
# Hide the progress bar when done
Write-Progress -Activity "Fetching Data from SharePoint" -Completed