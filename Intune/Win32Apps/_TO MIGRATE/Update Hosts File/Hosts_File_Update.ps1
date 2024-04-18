<#
.SYNOPSIS
    Scipt updates Host file on specified devices

.DESCRIPTION

.PARAMETER Mode
    Sets the mode of operation. Supported modes are Install or Uninstall.

.EXAMPLE 
    powershell.exe -executionpolicy bypass -file .\Hosts_File_Update.ps1 -Mode Install
    powershell.exe -executionpolicy bypass -file .\Hosts_File_Update.ps1 -Mode Uninstall

.NOTES
    - AUTHOR: Ashley Forde
    - Version: 1.0
    - Date: 28 Nov 2023
#>

#Region Parameters
[CmdletBinding()]
param(
	[Parameter(Mandatory = $false)]
	[ValidateNotNullOrEmpty()]
	[ValidateSet("Add","Remove")]
	[string]$Mode
)
# EndRegion Parameters
# Region Functions
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
function Initialize-Directories {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$HomeFolder
	)

	# Check if the path exists
	if (Test-Path -Path $HomeFolder) {
		Write-Verbose "Home folder exists..."
		# Force creating 00_Staging folder at a minimum if it is missing
		New-Item -Path "$HomeFolder" -Name "00_Staging" -ItemType "directory" -Force -Confirm:$false | Out-Null
	}
	else {
		Write-Verbose "Creating root folder..."
		New-Item -Path $HomeFolder -ItemType "directory" -Force -Confirm:$false
		if (-not $?) {
			Write-Verbose "Failed to create $HomeFolder"
		}

		# Create subfolders
		foreach ($subFolder in "00_Staging","01_Logs","02_Validation") {
			New-Item -Path "$HomeFolder\" -Name $subFolder -ItemType "directory" -Force -Confirm:$false
			if (-not $?) {
				Write-Verbose -Message "Failed to create sub-folder $subFolder under $HomeFolder"
			}
		}
	}

	# Calculate subfolder paths
	$StagingFolder = Join-Path -Path $HomeFolder -ChildPath "00_Staging"
	$LogsFolder = Join-Path -Path $HomeFolder -ChildPath "01_Logs"
	$ValidationFolder = Join-Path -Path $HomeFolder -ChildPath "02_Validation"

	# Return the folder paths as a custom object
	return @{
		HomeFolder = $HomeFolder
		StagingFolder = $StagingFolder
		LogsFolder = $LogsFolder
		ValidationFolder = $ValidationFolder
	}
}
# EndRegion Functions

# Comment: This region contains initialisations and variable assignments required for the script.   
# Region Initialisations
$HomeFolder = "C:\HUD"
$folderPaths = Initialize-Directories -HomeFolder $HomeFolder
# EndRegion Initialisations

# Assign the returned values to individual variables
$stagingFolderVar = $folderPaths.StagingFolder
$logsFolderVar = $folderPaths.LogsFolder
$validationFolderVar = $folderPaths.ValidationFolder

# Variables
$Date = Get-Date -Format "MM-dd-yyyy"
$AppName = "Hosts_File_Update"
$AppValidationFile = "$validationFolderVar\$AppName.txt"
$AppVersion = "1.1"
$LogFileName = "$($AppName)_${Mode}_$Date.log"

# Comment: The script initiates the install process, logs the initiation, and performs cleanup of the setup folder if it exists.
# Initate Install
Write-LogEntry -Value "Initializing Script" -Severity 1

# Attempt Cleanup of SetupFolder
if (Test-Path "$stagingFolderVar") {
	Remove-Item -Path "$($stagingFolderVar)\*" -Recurse -Force -ErrorAction SilentlyContinue
}

#$SetupFolder = (New-Item -ItemType "directory" -Path $stagingFolderVar -Name $AppName -Force).FullName


if ($Mode -eq "Add") {
	try {
		# Define the path to the hosts file and the backup file
		$hostsFilePath = "$env:SystemRoot\System32\drivers\etc\hosts"
		$backupFilePath = "$env:SystemRoot\System32\drivers\etc\hosts.old"

		# Create an empty hashtable to store IP addresses and hostnames
		$hostEntries = @{}

		# Add entries to the hashtable
		$hostEntries.Add("10.0.4.10","property.database.windows.net")
		#$hostEntries.Add("10.0.4.20","sql-corpreporting-dev.database.windows.net")
		#$hostEntries.Add("10.0.2.10","sql-easparx-prod.database.windows.net")
		$hostEntries.Add("10.0.4.30","sql-fpdreporting-dev.database.windows.net")
		#$hostEntries.Add("10.0.2.20","sql-projectregister-prod.database.windows.net")
		$hostEntries.Add("10.0.4.40","sql-reporting-prod.database.windows.net")
		$hostEntries.Add("10.0.2.30","sqldb-printix-prod.database.windows.net")
		$hostEntries.Add("10.0.5.5","dlprojectsdataprod.blob.core.windows.net")
		$hostEntries.Add("10.0.5.6","dlreportingdataprod.blob.core.windows.net")


		# Check if the backup file exists and if it doesn't, create it
		if (-not (Test-Path -Path $backupFilePath)) {
			Copy-Item -Path $hostsFilePath -Destination $backupFilePath
			Write-LogEntry -Value "Created a backup of the hosts file at $backupFilePath." -Severity 1
		}

		# Add the entries from the hashtable to the hosts file
		foreach ($entry in $hostEntries.GetEnumerator()) {
			$ipAddress = $entry.Key
			$hostname = $entry.Value
			$newHostsEntry = "$ipAddress   $hostname"

			# Check if the hosts file exists
			if (Test-Path -Path $hostsFilePath) {
				# Check if the entry already exists in the hosts file
				if ((Get-Content -Path $hostsFilePath) -notcontains $newHostsEntry) {
					# Append the new entry to the hosts file
					Add-Content -Path $hostsFilePath -Value $newHostsEntry
					Write-LogEntry -Value "Added entry: $newHostsEntry" -Severity 1
				} else {
					Write-LogEntry -Value "The entry already exists in the hosts file: $newHostsEntry" -Severity 1
				}
			} else {
				Write-LogEntry -Value "Hosts file not found at $hostsFilePath." -Severity 1
			}
		}

		# Add Validation File
		New-Item -ItemType File -Path $AppValidationFile -Force -Value $AppVersion
		Write-LogEntry -Value "Validation file has been created at $AppValidationFile" -Severity 1
		Write-LogEntry -Value "Install of $AppName is complete" -Severity 1

	} catch {
		Write-LogEntry -Value "error updating hosts file" -Severity 3
		return # Stop execution of the script after logging a critical error
	}


} elseif ($Mode -eq "Remove") {
	try {
		# Define the path to the hosts file and the backup file
		$hostsFilePath = "$env:SystemRoot\System32\drivers\etc\hosts"
		$backupFilePath = "$env:SystemRoot\System32\drivers\etc\hosts.old"

		# Function to restore the backup hosts file
		function RestoreHostsFile {
			if (Test-Path -Path $backupFilePath) {
				Copy-Item -Path $backupFilePath -Destination $hostsFilePath -Force
				Write-LogEntry -Value "Restored the backup hosts file to its original state." -Severity 1
			} else {
				Write-LogEntry -Value "Backup hosts file not found at $backupFilePath." -Severity 3
			}
		}

		RestoreHostsFile
		
		# Add Validation File
		Remove-Item -Path $AppValidationFile -Force -Confirm:$false
		Write-LogEntry -Value "Validation file has been removed at $AppValidationFile" -Severity 1
		Write-LogEntry -Value "Uninstall of $AppName is complete" -Severity 1


	} catch {
		Write-LogEntry -Value "error updating hosts file" -Severity 3
		return # Stop execution of the script after logging a critical error
	}

}

#ENDS