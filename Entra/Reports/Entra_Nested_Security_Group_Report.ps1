<#
Title: Nested Group Discovery Script
Author: Ashley Forde
Date: 12 October 2023
Version: 1.0
Description:
This PowerShell script is designed to discover and report on nested group structures within a Microsoft 365 environment. The script performs the following operations:

    1. Sets the working directory to C:\HUD\06_Reporting.
    2. Authenticates with Microsoft Graph.
    3. Connects to Exchange Online.
    4. Retrieves all Microsoft 365 Groups, filtering out Unified Groups and groups with specific naming conventions.
    5. Retrieves all Distribution Groups from Exchange Online.
    6. Filters out Distribution Groups from the list of Microsoft 365 Groups.
    7. Defines a function to recursively discover nested groups up to 5 layers deep.
    8. Iterates through each group, calling the nested group discovery function and collecting results.
    9. Displays a progress bar during processing.
    10. Once all groups have been processed, it marks the progress as completed.
    11. Stores the nested group information in a variable `$results`, ready for further processing or export.
    12. Appends the current date to the `$results` array for record-keeping.

#>

Connect-MgGraph -Scopes $Scopes -NoWelcome | Out-Null
$UPN = (Get-MgContext).Account

# Connect to Exchange Online
Connect-ExchangeOnline -UserPrincipalName $UPN -ShowBanner:$false

# Get all groups from Microsoft Graph
$allGroups = Get-MgGroup -All | Where-Object { $_.DisplayName -notlike "PIM - *" }

# Get all Distribution Groups from Exchange Online
$allDistGroups = Get-DistributionGroup | Select-Object -ExpandProperty DisplayName

# Filter out distribution groups from $allGroups
$allGroups = $allGroups | Where-Object { $allDistGroups -notcontains $_.DisplayName }

# Function to get nested groups  
function Get-NestedGroups ($rootGroupId,$groupId,$layer,$nestedGroupNames,$groupTypes) {
	$nestedGroups = Get-MgGroupMemberOf -GroupId $groupId
	$output = @()

	foreach ($group in $nestedGroups) {
		$nestedGroupName = $group.AdditionalProperties["displayName"]
		$nestedGroupNames[$layer] = $nestedGroupName
		$obj = [pscustomobject]@{
			"id" = $rootGroupId
			"Group" = $nestedGroupNames[0]
			"GroupTypes" = ($groupTypes -join ', ')
			"Nested Layer 1" = $nestedGroupNames[1]
			"Nested Layer 2" = $nestedGroupNames[2]
			"Nested Layer 3" = $nestedGroupNames[3]
			"Nested Layer 4" = $nestedGroupNames[4]
		}
		$output += $obj

		if ($layer -lt 5) {
			$output += Get-NestedGroups -rootGroupId $rootGroupId -GroupId $group.id -layer ($layer + 1) -nestedGroupNames $nestedGroupNames -groupTypes $group.GroupTypes
		}
	}
	return $output
}

# Display all groups and their nested groups  
$results = @()
$groupCount = $allGroups.count

foreach ($group in $allGroups) {
	$nestedGroupNames = @($group.DisplayName,"","","","","")
	$groupResult = Get-NestedGroups -rootGroupId $group.id -GroupId $group.id -layer 1 -nestedGroupNames $nestedGroupNames -groupTypes $group.GroupTypes

	if ($groupResult.count -eq 0) {
		$obj = [pscustomobject]@{
			"id" = $group.id
			"Group" = $group.DisplayName
			"GroupTypes" = ($group.GroupTypes -join ', ')
			"Nested Layer 1" = ""
			"Nested Layer 2" = ""
			"Nested Layer 3" = ""
			"Nested Layer 4" = ""
		}
		$results += $obj
	} else {
		$results += $groupResult
	}

	$progress = [math]::Round((($results.count) / $groupCount) * 100,2)
	$progress = [math]::Min($progress,100)
	Write-Progress -Activity "Processing Groups" -Status "Processing $($group.DisplayName)" -PercentComplete $progress
}

Write-Progress -Activity "Processing Groups" -Status "Completed" -Completed

# Now, we have all the results stored in $results  
$Date = Get-Date -Format "dd.MM.yyyy h.mm tt"
$FileName = "Nested_Entra_Groups"

# Add assembly and import namespace  
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
$SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog

# Configure the SaveFileDialog  
$SaveFileDialog.Filter = "XLSX Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
$SaveFileDialog.Title = "Save the Report"
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
