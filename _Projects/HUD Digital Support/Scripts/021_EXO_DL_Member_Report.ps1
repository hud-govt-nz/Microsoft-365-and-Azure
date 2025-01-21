Clear-Host
Write-Host '## Exchange Online: DL Group Member Report ##' -ForegroundColor Yellow

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

do {

# Obtain List Name
Write-Host ""
$ListName = Read-Host "Enter the Distribution List name or press 'q' to quit"

if ($ListName -eq 'q') {
    break
}

# Check if the list is a Dynamic Distribution List
$IsDynamic = $null -ne (Get-DynamicDistributionGroup -Identity $ListName -ErrorAction SilentlyContinue)

if ($IsDynamic) {
    Write-Host ""
    Write-Host "Exporting members of dynamic distribution list $ListName." -ForegroundColor Green
    $Results = Get-DynamicDistributionGroupMember -Identity $ListName | Select-Object DisplayName, PrimarySMTPAddress
} elseif ($null -ne (Get-DistributionGroup -Identity $ListName -ErrorAction SilentlyContinue)) {
    Write-Host ""
    Write-Host "Exporting members of distribution list $ListName." -ForegroundColor Green
    $Results = Get-DistributionGroupMember -Identity $ListName | Select-Object DisplayName, PrimarySMTPAddress
} else {
    Write-Host ""
    Write-Host "No Distribution List or Dynamic Distribution List found with the name: $ListName"
}

# Now, we have all the results stored in $results  
$Date = Get-Date -Format "dd.MM.yyyy h.mm tt"
$FileName = "$($ListName)_Members"  

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
    $Results | Export-Excel $SelectedFilePath -AutoSize -AutoFilter -WorksheetName $Date -FreezeTopRow -BoldTopRow

    $excelPackage = Open-ExcelPackage -Path $SelectedFilePath
    $worksheet = $excelPackage.Workbook.Worksheets["$Date"]

    # Assuming headers are in row 1 and you start from row 2
    $startRow = 1
    $endRow = $worksheet.Dimension.End.Row
    $startColumn = 1
    $endColumn = $worksheet.Dimension.End.Column

    # Set horizontal alignment to left for all cells in the used range
    for ($col = $startColumn; $col -le $endColumn; $col++) {
        for ($row = $startRow; $row -le $endRow; $row++) {
            $worksheet.Cells[$row, $col].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
        }
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
        Write-Host "Save operation canceled." -ForegroundColor Yellow  
        }  
[System.GC]::Collect()

} while ($true)

Disconnect-ExchangeOnline -Confirm:$false | Out-Null