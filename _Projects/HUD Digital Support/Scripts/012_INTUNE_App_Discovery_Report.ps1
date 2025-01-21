<#
.SYNOPSIS
    Intune Discover Applications Report
.DESCRIPTION
    Lists all discovered Applications and associated devices.

.PARAMETER OSType
    - Either Windows, Android, iOS

.EXAMPLE 
    powershell.exe -executionpolicy bypass -file .\Enable_PIM_Assignment.ps1 -Platform Windows 

.NOTES
    - AUTHOR: Ashley Forde
    - Version: 2.0
    - Date: 12 Oct 2023

#>
# Region Parameters
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("windows","androidWorkProfile","ios")]
    [string]$Platform
)

if (-not $Platform) {
    $Platform = Read-Host -Prompt "Please select a platform (windows, androidWorkProfile, ios)"
    if ($Platform -notin @("windows", "androidWorkProfile", "ios")) {
        Write-Host "Invalid platform selected. Exiting script." -ForegroundColor Red
        exit
    }
}

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
	try {

		# Gather Device Info bas on OSType and Ownership
		switch -Wildcard ($Platform) {
			"Windows" { $PlatformStr = "windows" }
			"iOS" { $PlatformStr = "ios" }
			"androidWorkProfile" { $PlatformStr = "androidWorkProfile" }
			default { $PlatformStr = "Unknown" }
		}

		Write-Host "Platform Selected: $PlatformStr"

		# Discovered Apps
		$discoveredapps = Get-MgDeviceManagementDetectedApp -All | Where-Object { $_.Platform -eq $PlatformStr }
		$SelectedApps = $discoveredapps | Sort-Object DeviceCount -Descending | Out-GridView -Title "Discovered Apps" -PassThru

		Write-Host "Apps Selected: $SelectedApps"

		$Result = @() # Create an array to hold the result

		# Iterate through each selected app
		foreach ($app in $SelectedApps) {
			$appID = $app.id
			$appName = $app.DisplayName # Capture the display name of the app
			$uri = "https://graph.microsoft.com/v1.0/deviceManagement/detectedApps('$appID')/managedDevices" #no PS cmdlet equivalent 

			# Handle pagination
			do {
				$InstalledDevices = Invoke-MgGraphRequest -Method Get -Uri $uri -OutputType PSObject
				foreach ($device in $InstalledDevices.value) {
					$deviceInfo = [pscustomobject]@{
						IntuneDeviceID = $device.id
						DeviceName = $device.DeviceName
						OS = $device.OperatingSystem
						UserPrincipalname = $device.emailAddress # Assuming the email address is the UPN
						AppID = $appID
						AppDisplayName = $appName # Include the app display name
					}
					$Result += $deviceInfo
				}
				$uri = $InstalledDevices. '@odata.nextLink'
			} while ($null -ne $uri)
		}

		# Output the result
		$Selection = $Result | Out-GridView -Title "Selected Apps" -PassThru

		# Add assembly and import namespace  
		Add-Type -AssemblyName System.Windows.Forms
		[System.Windows.Forms.Application]::EnableVisualStyles()
		$SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog

		# Configure the SaveFileDialog  
		$SaveFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
		$SaveFileDialog.Title = "Save the Report"
		$SaveFileDialog.FileName = "$($AppName)_$Date.CSV"

		# Show the SaveFileDialog and get the selected file path  
		$SaveFileResult = $SaveFileDialog.ShowDialog()
		if ($SaveFileResult -eq [System.Windows.Forms.DialogResult]::OK) {
			$SelectedFilePath = $SaveFileDialog.FileName
			$Selection | Export-Csv $SelectedFilePath -NoTypeInformation -Encoding UTF8

			Write-Host "The report $FileName has been saved in $($SelectedFilePath)" -ForegroundColor Green
		} else {
			Write-Host "Save operation canceled." -ForegroundColor Yellow
		}

	} catch [System.Exception]{
		Write-Host -Value "Unable to Gather Information. Errormessage: $($_.Exception.Message)"
	}