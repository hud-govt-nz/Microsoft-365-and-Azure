<#
.SYNOPSIS
    Perform Intune Device Sync on Devices in bulk

.DESCRIPTION
    SScript can be used to bulk sync devices

.EXAMPLE 
    powershell.exe -executionpolicy bypass -file .\Sync-IntuneDevices.ps1 -Platform 'Windows' -Ownership 'Company'
    powershell.exe -executionpolicy bypass -file .\Sync-IntuneDevices.ps1 -Platform 'iOS' -Ownership 'Personal'
    powershell.exe -executionpolicy bypass -file .\Sync-IntuneDevices.ps1 -ShowAll

.NOTES
    - AUTHOR: Ashley Forde
    - Version: 2.0
    - Date: 25 Oct 2023

#>

# Region Parameters
[CmdletBinding()]
param(
	[Parameter(Mandatory = $false)]
	[ValidateNotNullOrEmpty()]
	[switch]$ShowAll,
	[ValidateSet("Windows","Android","iOS")]
	[string]$Platform,
	[ValidateSet("Company","Personal")]
	[string]$Ownership
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
$AppName = "Device_Sync"
$LogFileName = "$($AppName)_$Date.log"
$SyncedDeviceCount = 0

Write-LogEntry -Value "Connecting to Microsoft Graph" -Severity 1

try {
	# Define Graph scopes
	$scopes = @("DeviceManagementManagedDevices.PrivilegedOperations.All","DeviceManagementManagedDevices.ReadWrite.All","DeviceManagementManagedDevices.Read.All")

	# Connect to Graph 
	Connect-MgGraph -Scopes $scopes -NoWelcome

	try {
		# Gather Device Info bas on OSType and Ownership
		if ($ShowAll) {
			# Collect all devices
			$DiscoveredDevices = Get-MgDeviceManagementManagedDevice -All
		} else {
			# Gather Device Info based on OSType and Ownership
			switch -Wildcard ($Platform) {
				"Windows" { $PlatformStr = "Windows" }
				"iOS" { $PlatformStr = "iOS" }
				"androidWorkProfile" { $PlatformStr = "Android" }
				default { $PlatformStr = "Unknown" }
			}

			Write-LogEntry "Platform Selected: $PlatformStr" -Severity 1

			# Gather Device Info based on OSType and Ownership
			switch -Wildcard ($Ownership) {
				"personal" { $OwnershipStr = "personal" }
				"company" { $OwnershipStr = "company" }
				default { $OwnershipStr = "Unknown" }
			}

			Write-LogEntry "Platform Selected: $OwnershipStr" -Severity 1

			# Collect devices in scope
			$DiscoveredDevices = Get-MgDeviceManagementManagedDevice -Filter "contains(operatingsystem,'$PlatformStr') and ManagedDeviceOwnerType eq '$OwnershipStr'" -All
		}

		$SelectedDevices = $DiscoveredDevices | Sort-Object LastSyncDateTime -Descending | Out-GridView -Title "Managed Devices" -PassThru

		foreach ($Device in $SelectedDevices)
		{

			Sync-MgDeviceManagementManagedDevice -ManagedDeviceId $Device.id

			Write-LogEntry "Sending Sync request to Device with Device name $($Device.DeviceName)" -Severity 1

			$SyncedDeviceCount++ # Increment the counter
		}

		Write-LogEntry "Total Devices Synced: $SyncedDeviceCount" -Severity 1

		Disconnect-Graph | Out-Null




	}
	catch {
		<#Do this if a terminating exception happens#>
	}



}
catch [System.Exception]{
	Write-LogEntry -Value "Unable to connect to Microsoft Graph. Errormessage: $($_.Exception.Message)" -Severity 3
}
