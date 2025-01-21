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

# Begin Region Functions
function Write-LogEntry {
	param(
		[Parameter(Mandatory = $true,HelpMessage = "Value added to the log file.")]
		[ValidateNotNullOrEmpty()]
		[string]$Value,
		[Parameter(Mandatory = $true,HelpMessage = "Severity for the log entry. 1 for Informational, 2 for Warning and 3 for Error.")]
		[ValidateNotNullOrEmpty()]
		[ValidateSet("1","2","3")]
		[string]$Severity,
		[Parameter(Mandatory = $false,HelpMessage = "Name of the log file that the entry will written to.")]
		[ValidateNotNullOrEmpty()]
		[string]$FileName = $LogFileName
	)
	# Determine log file location
	$LogFilePath = Join-Path -Path $logsFolderVar -ChildPath $FileName

	# Construct time stamp for log entry
	$Time = -join @((Get-Date -Format "HH:mm:ss.fff")," ",(Get-WmiObject -Class Win32_TimeZone | Select-Object -ExpandProperty Bias))

	# Construct date for log entry
	$Date = (Get-Date -Format "MM-dd-yyyy")

	# Construct context for log entry
	$Context = $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)

	# Construct final log entry
	$LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""$($LogFileName)"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"

	# Add value to log file
	try {
		Out-File -InputObject $LogText -Append -NoClobber -Encoding Default -FilePath $LogFilePath -ErrorAction Stop
		if ($Severity -eq 1) {
			Write-Verbose -Message $Value
		}
		elseif ($Severity -eq 3) {
			Write-Warning -Message $Value
		}
	}
	catch [System.Exception]{
		Write-Warning -Message "Unable to append log entry to $LogFileName.log file. Error message at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
	}
}

# End Region Functions

# Begin Region Variables
$LogsFolderVar = "C:\HUD\01_Logs\"
$Date = Get-Date -Format "MM-dd-yyyy"
$AppName = "Discovered_Apps"
$LogFileName = "$($AppName)_$Date.log"

Write-LogEntry -Value "Connecting to Microsoft Graph" -Severity 1

try {
	# Define Graph scopes
	$scopes = @("DeviceManagementApps.Read.All",
		"DeviceManagementConfiguration.Read.All",
		"DeviceManagementManagedDevices.Read.All",
		"DeviceManagementServiceConfig.Read.All",
		"Directory.Read.All",
		"User.Read.All",
		"User.ReadBasic.All",
		"User.ReadWrite"
	)

	# Connect to Graph 
	Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop

	try {

		# Gather Device Info bas on OSType and Ownership
		switch -Wildcard ($Platform) {
			"Windows" { $PlatformStr = "windows" }
			"iOS" { $PlatformStr = "ios" }
			"androidWorkProfile" { $PlatformStr = "androidWorkProfile" }
			default { $PlatformStr = "Unknown" }
		}

		Write-LogEntry "Platform Selected: $PlatformStr" -Severity 1

		# Discovered Apps
		$discoveredapps = Get-MgDeviceManagementDetectedApp -All | Where-Object { $_.Platform -eq $PlatformStr }
		$SelectedApps = $discoveredapps | Sort-Object DeviceCount -Descending | Out-GridView -Title "Discovered Apps" -PassThru

		Write-LogEntry "Apps Selected: $SelectedApps" -Severity 1

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
		Write-LogEntry -Value "Unable to Gather Information. Errormessage: $($_.Exception.Message)" -Severity 3
	}

} catch [System.Exception]{
	Write-LogEntry -Value "Unable to connect to Microsoft Graph. Errormessage: $($_.Exception.Message)" -Severity 3
}
