<#
.SYNOPSIS
    SharePoint Online Reporting: Export List Item Report for Specified Site

.DESCRIPTION
    This script prompts the user for a SharePoint Online site URL and list name, then exports a detailed report of list items to an Excel file.
    The report includes various metadata such as item ID, file name, breadcrumb trail, created/modified dates, authors, and more.

.AUTHOR
    Ashley Forde

.VERSION
    1.0

.DATE
    6 Nov 2023

.EXAMPLE
    .\ExportListItemReport.ps1

    Running the script will prompt the user interface to input the site URL and list name. Upon submission,
    it fetches the list items and generates an Excel report, offering the user a save dialog to store the report on the local system.

.NOTES
    - This script requires the SharePoint PnP PowerShell module and the ImportExcel module to be installed.
    - System.Windows.Forms assembly is used for the dialog interface; hence, the script should be run on a machine with a GUI.
    - Excel is not required to be installed as the Export-Excel cmdlet from the ImportExcel module is used.

# Requires -Modules SharePointPnPPowerShellOnline, ImportExcel
#>

Clear-Host
Write-host "## SharePoint Online Reporting: Export List Item Report for Specified Site ##" -ForegroundColor Yellow
<#
# Load necessary assembly
Add-Type -AssemblyName System.Windows.Forms

# Create a new form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'SPO Context'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'

# Create labels
$label1 = New-Object System.Windows.Forms.Label
$label1.Location = New-Object System.Drawing.Point(10,20)
$label1.Size = New-Object System.Drawing.Size(280,20)
$label1.Text = 'Site URL:'
$form.Controls.Add($label1)

$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(10,70)
$label2.Size = New-Object System.Drawing.Size(280,20)
$label2.Text = 'List Name:'
$form.Controls.Add($label2)

# Create textboxes
$textbox1 = New-Object System.Windows.Forms.TextBox
$textbox1.Location = New-Object System.Drawing.Point(10,40)
$textbox1.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textbox1)

$textbox2 = New-Object System.Windows.Forms.TextBox
$textbox2.Location = New-Object System.Drawing.Point(10,90)
$textbox2.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textbox2)

# Create OK button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(10,120)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

# Show the form
$form.Topmost = $true
$form.Add_Shown({$textbox1.Select()})
$result = $form.ShowDialog()

$uri = [System.Uri]$textbox1.Text
$RootURL = "$($uri.Scheme)://$($uri.Host)"
$Site = $uri.Segments[-1]

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $SiteURL = $textbox1.Text
    $ListName = $textbox2.Text
    # Output or use the strings as needed
    Write-Host ""
    Write-Host "Root URL    : $RootURL"
    Write-Host "Site URL    : $SiteURL"
    Write-Host "List Name   : $ListName"
}

#>

$SiteURL = Read-Host "Please enter the site URL"
$ListName = Read-Host "Please enter the Library Name"

# Create a Uri object from the string
$uriObject = [System.Uri]::new($SiteURL)

# Combine the scheme and the host to get the base URL
$baseUrl = $uriObject.Scheme + "://" + $uriObject.Host
$lastPart = $uriObject.Segments[-1]

$env:PNPPOWERSHELL_UPDATECHECK="false"
Connect-PnPOnline -Url $SiteURL -Interactive

Write-Host "Connecting to $SiteURL" -ForegroundColor Yellow

# Check if the specified list exists
$List = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue
if ($null -eq $List) {
    Write-Host "The specified list '$ListName' does not exist." -ForegroundColor Red
    # Optionally break out of the script or prompt the user to try again
    break
}

# Get the list
$ListItems = (Get-PnPListItem -List $ListName -PageSize 100 | Where-Object{$_.FieldValues.FSObjType -ne 1}) | Sort-Object id

# Count for progress bar
$TotalCount = $ListItems.count

Write-Host ""
Write-Host "Total number of items in list $($TotalCount)"


$Progress = 0

# Create an array to hold the data
$LibraryList = @()

#Get-PnPListItem -List $ListName -PageSize 100 | Where-Object{$_.FieldValues.FSObjType -ne 1} 
$ListItems | ForEach-Object {
    $Progress++
    Write-Progress -Activity "Fetching Data from SharePoint" -Status "$Progress of $TotalCount" -PercentComplete ($Progress/$TotalCount*100)
    
    $item = $_ | Select-Object id,
        @{label="File Name";expression={$_.FieldValues.FileLeafRef}},
        @{label="File Type";expression={$_.FieldValues.File_x0020_Type}},      
        @{label="File Size (KB)";Expression={[math]::Round($_.FieldValues.File_x0020_Size / 1KB, 0)}},
        @{label="Breadcrumb Trail";expression={$_.FieldValues.FileDirRef}},
        @{label="Created Date";expression={$_.FieldValues.Created}},
        @{label="Modified Date";expression={$_.FieldValues.Modified}},
        @{Label="Created By";e={$($_.FieldValues.Created_x0020_By).split("|")[2]}},
        @{Label="Modified By";e={$($_.FieldValues.Modified_x0020_By).split("|")[2]}},
        @{Label="Link";e={$baseUrl + $_.FieldValues.FileRef}},
        @{Label="Current Version";e={$($_.FieldValues._IsCurrentVersion)}},
        @{Label="Content Version";e={$_.FieldValues.ContentVersion}},
        @{Label="Compliance Flags";e={$_.FieldValues._ComplianceFlags}},
        @{Label="Compliance Tag";e={$_.FieldValues._ComplianceTag}},
        @{Label="Compliance Tag Written Time";e={
                $dateString = $_.FieldValues._ComplianceTagWrittenTime
                if ($dateString -ne $null) {
                    $date = [DateTime]::ParseExact($dateString, "yyyy-MM-ddTHH:mm:ssZ", [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal)
                    $date.ToString("dd/MM/yyyy h:mm")
                } else {
                    $null
                }
            }},
        @{Label="UniqueId";e={$_.FieldValues.UniqueId}},
        @{Label="GUID";e={$_.FieldValues.GUID}} 

        
    $LibraryList += $item
}

Write-Host "Open Save Dialog"

$Date = Get-Date -Format "dd.MM.yyyy h.mm tt"
$FileName = "SPOSite-$($lastPart) Library-$($ListName)"  

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
    $LibraryList | Export-Excel $SelectedFilePath -AutoSize -AutoFilter -WorksheetName $Date -FreezeTopRow -BoldTopRow
    
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
    foreach ($item in $LibraryList) {
        # Trim the hyperlink
        $hyperlink = $item.Link.Trim()

        # Set the hyperlink directly
        $cell = $worksheet.Cells[$rowIndex, 10]
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