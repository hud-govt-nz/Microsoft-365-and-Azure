Clear-Host
Write-host "## Intune: App Assignment Report ##" -ForegroundColor Yellow

# Requirements
#Requires -Modules Microsoft.Graph.Authentication

# Connect to Graph
try {
    Connect-MgGraph `
        -ClientId $env:DigitalSupportAppID `
        -TenantId $env:DigitalSupportTenantID `
        -CertificateThumbprint $env:DigitalSupportCertificateThumbprint `
        -NoWelcome
    Write-Host "Connected" -ForegroundColor Green
        
    } catch {
	    Write-Host "Error connecting to Microsoft Graph. Please check your credentials and network connection." -ForegroundColor Red
	    exit 1
}

Write-Host "Generating Report..." -ForegroundColor Yellow

# Define filters for Windows and iOS
$filterWindows = "(isof('microsoft.graph.windowsStoreApp') or isof('microsoft.graph.microsoftStoreForBusinessApp') or isof('microsoft.graph.officeSuiteApp') or isof('microsoft.graph.win32LobApp') or isof('microsoft.graph.windowsMicrosoftEdgeApp') or isof('microsoft.graph.windowsPhone81AppX') or isof('microsoft.graph.windowsPhone81StoreApp') or isof('microsoft.graph.windowsPhoneXAP') or isof('microsoft.graph.windowsAppX') or isof('microsoft.graph.windowsMobileMSI') or isof('microsoft.graph.windowsUniversalAppX') or isof('microsoft.graph.webApp') or isof('microsoft.graph.windowsWebApp') or isof('microsoft.graph.winGetApp')) and (microsoft.graph.managedApp/appAvailability eq null or microsoft.graph.managedApp/appAvailability eq 'lineOfBusiness' or isAssigned eq true)"
$filterIOS = "((isof('microsoft.graph.managedIOSStoreApp') and microsoft.graph.managedApp/appAvailability eq 'lineOfBusiness') or isof('microsoft.graph.iosLobApp') or isof('microsoft.graph.iosStoreApp') or isof('microsoft.graph.iosVppApp') or isof('microsoft.graph.managedIOSLobApp') or (isof('microsoft.graph.managedIOSStoreApp') and microsoft.graph.managedApp/appAvailability eq 'global') or isof('microsoft.graph.webApp') or isof('microsoft.graph.iOSiPadOSWebClip')) and (microsoft.graph.managedApp/appAvailability eq null or microsoft.graph.managedApp/appAvailability eq 'lineOfBusiness' or isAssigned eq true)"
$filterAndroid = "((isof('microsoft.graph.androidManagedStoreApp') and microsoft.graph.androidManagedStoreApp/isSystemApp eq true) or isof('microsoft.graph.androidLobApp') or isof('microsoft.graph.androidStoreApp') or (isof('microsoft.graph.managedAndroidStoreApp') and microsoft.graph.managedApp/appAvailability eq microsoft.graph.managedAppAvailability'lineOfBusiness') or isof('microsoft.graph.managedAndroidLobApp') or (isof('microsoft.graph.managedAndroidStoreApp') and microsoft.graph.managedApp/appAvailability eq microsoft.graph.managedAppAvailability'global') or (isof('microsoft.graph.androidManagedStoreApp') and microsoft.graph.androidManagedStoreApp/isSystemApp eq false) or isof('microsoft.graph.webApp')) and (microsoft.graph.managedApp/appAvailability eq null or microsoft.graph.managedApp/appAvailability eq 'lineOfBusiness' or isAssigned eq true)"

# Get Windows and iOS apps
$WindowsApps = Get-MgBetaDeviceAppManagementMobileApp -Filter $filterWindows
$IOSApps = Get-MgBetaDeviceAppManagementMobileApp -Filter $filterIOS
$AndroidApps = Get-MgBetaDeviceAppManagementMobileApp -Filter $filterAndroid

# Function to process apps
function Process-Apps ($apps,$appType) {
	$customObjects = @()
	$totalApps = $apps.count
	$currentProgress = 0

	foreach ($app in $apps) {
		# Update progress bar
		$currentProgress++
		$remainingApps = $totalApps - $currentProgress
		Write-Progress -Activity "Processing $appType Apps" -Status "Processing $($app.DisplayName) - $remainingApps apps remaining" -PercentComplete ($currentProgress / $totalApps * 100)

		$appAssignments = Get-MgDeviceAppManagementMobileAppAssignment -MobileAppId $app.id
		if ($null -eq $appAssignments) {
			# Create a custom object with no group assignment
			$customObject = New-Object -TypeName PSObject -Property @{
				DisplayName = $app.DisplayName
				AssignmentGroup = "No Group Assignment"
				Intent = $null
				AppID = $app.id
				LastModifiedDateTime = $app.LastModifiedDateTime
			} | Select-Object DisplayName,AssignmentGroup,Intent,AppID,LastModifiedDateTime | Sort-Object -Property DisplayName
			$customObjects += $customObject
		} else {
			foreach ($assignment in $appAssignments) {
				$formattedGroupId = $assignment.id -replace '_.+$'
				if ($formattedGroupId -eq "adadadad-808e-44e2-905a-0b7873a8a531") {
					$formattedGroupId = "All Devices"
				} elseif ($formattedGroupId -eq "acacacac-9df4-4c7d-9d50-4ef0226f57a9") {
					$formattedGroupId = "All Users"
				} else {
					try {
						$group = Get-MgGroup -GroupId $formattedGroupId -ErrorAction Stop
						$formattedGroupId = $group.DisplayName
					} catch {
						# If there's an error, keep the formattedGroupId as is
					}
				}
				$customObject = New-Object -TypeName PSObject -Property @{
					DisplayName = $app.DisplayName
					AssignmentGroup = $formattedGroupId
					Intent = $assignment.Intent
					AppID = $app.id
					LastModifiedDateTime = $app.LastModifiedDateTime
				} | Select-Object DisplayName,AssignmentGroup,Intent,AppID,LastModifiedDateTime | Sort-Object -Property DisplayName
				$customObjects += $customObject
			}
		}
	}
	return $customObjects
}

# Process Windows and iOS apps separately
$WindowsReport = Process-Apps -apps $WindowsApps -appType "Windows"
$IOSReport = Process-Apps -apps $IOSApps -appType "iOS"
$AndroidReport = Process-Apps -apps $AndroidApps -appType "Android"

# Export Windows Report
$Date = Get-Date -Format "dd.MM.yyyy"
$FileName = "Intune_App_Assignment_Report_$($Date).xlsx"

Write-Host "Open Save Dialog"

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
	$WindowsReport | Export-Excel -Path $SelectedFilePath -WorksheetName "WinApps" -AutoFilter -FreezeTopRow -BoldTopRow -AutoNameRange
	$IOSReport | Export-Excel -Path $SelectedFilePath -WorksheetName "iOSApps" -AutoFilter -FreezeTopRow -BoldTopRow -AutoNameRange
	$AndroidReport | Export-Excel -Path $SelectedFilePath -WorksheetName "AndroidApps" -AutoFilter -FreezeTopRow -BoldTopRow -AutoNameRange

	$excelPackage = Open-ExcelPackage -Path $SelectedFilePath

	foreach ($Sheet in $excelPackage.Workbook.Worksheets) {
		$worksheet = $Sheet

		# Check if the worksheet is not null
		if ($null -eq $worksheet) {
			Write-Host "Worksheet not found." -ForegroundColor Red
			continue
		}

		# Get the range of used cells in the worksheet
		$startRow = 1 # Assuming headers are in row 1, start from row 1
		$endRow = $worksheet.Dimension.End.Row
		$startColumn = 1
		$endColumn = $worksheet.Dimension.End.Column

		# Set horizontal alignment to left for all cells in the used range
		for ($row = $startRow; $row -le $endRow; $row++) {
			for ($col = $startColumn; $col -le $endColumn; $col++) {
				$worksheet.Cells[$row,$col].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
			}
		}

		# Autosize columns if needed
		foreach ($column in $startColumn..$endColumn) {
			$worksheet.Column($column).AutoFit()
		}
	}

	# Save and close the Excel package
	$excelPackage.Save()
	Close-ExcelPackage $excelPackage -Show

	Write-Host "The report $FileName has been saved in $($SelectedFilePath)" -ForegroundColor Green
} else {
	Write-Host "Save cancelled" -ForegroundColor Yellow
}
