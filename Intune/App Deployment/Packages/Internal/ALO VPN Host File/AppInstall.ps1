<#
.SYNOPSIS
    ALOVPN Host File Update

.DESCRIPTION
    Script to install ALOVPN Host File Update

.PARAMETER Mode
    Sets the mode of operation. Supported modes are Install or Uninstall

.EXAMPLE 
    powershell.exe -executionpolicy bypass -file .\appinstall.ps1 -Mode Install
    powershell.exe -executionpolicy bypass -file .\appinstall.ps1 -Mode Uninstall

.NOTES
    - AUTHOR: 
    - Version: 
    - Date: 
#>
# Region Parameters
[CmdletBinding()]
param(
	[Parameter(Mandatory = $false)]
	[ValidateNotNullOrEmpty()]
	[ValidateSet("Add","Remove")]
	[string]$Mode
)

# Reference functions.ps1
. "$PSScriptRoot\functions.ps1"

# Application Variables
$AppName = "Host File Update"
$AppVersion = "3.0"

# Initialize Directories
$folderpaths = Initialize-Directories -HomeFolder C:\HUD\

# Template Variables
$Date = Get-Date -Format "MM-dd-yyyy"
$stagingFolderVar = $folderPaths.StagingFolder
$logsFolderVar = $folderPaths.LogsFolder
$LogFileName = "$($AppName)_${Mode}_$Date.log"
$validationFolderVar = $folderPaths.ValidationFolder
$AppValidationFile = "$validationFolderVar\$AppName.txt"

# Begin Setup
Write-LogEntry -Value "Initiating setup process" -Severity 1

# Add/Remove
switch ($Mode) {
    "Add" {
        try {
            # Define the path to the hosts file and the backup file
            $hostsFilePath = "$env:SystemRoot\System32\drivers\etc\hosts"
            $backupDirPath = "$env:SystemRoot\System32\drivers\etc\old host files"
            $backupFilePath = "$backupDirPath\hosts.old"

            # Create the backup directory if it doesn't exist
            if (-not (Test-Path -Path $backupDirPath)) {
                New-Item -ItemType Directory -Path $backupDirPath
                Write-LogEntry -Value "Created backup directory at $backupDirPath." -Severity 1
            }

            # Move the current hosts file to the backup directory
            if (Test-Path -Path $hostsFilePath) {
                Move-Item -Path $hostsFilePath -Destination $backupFilePath -Force
                Write-LogEntry -Value "Moved current hosts file to $backupFilePath." -Severity 1
            }

            # Create a new array with the desired entries
            $newHostEntries = @(
                "# Azure Private Endpoints"
                "10.0.4.10    property.database.windows.net"
                "10.0.4.30    sql-fpdreporting-dev.database.windows.net"
                "10.0.4.40    sql-reporting-prod.database.windows.net"
                "10.0.5.5     dlprojectsdataprod.blob.core.windows.net"
                "10.0.5.6     dlreportingdataprod.blob.core.windows.net"
                "10.0.5.10    dlreportingdataprod.dfs.core.windows.net"
                ""
            )

            # Write the new entries to the hosts file
            Add-Content -Path $hostsFilePath -Value $newHostEntries
            Write-LogEntry -Value "Created new hosts file with updated entries." -Severity 1

            # Add Validation File
            New-Item -ItemType File -Path $AppValidationFile -Force -Value $AppVersion
            Write-LogEntry -Value "Validation file has been created at $AppValidationFile" -Severity 1
            Write-LogEntry -Value "Update of $AppName is complete" -Severity 1

        } catch {
            Write-LogEntry -Value "Error updating hosts file" -Severity 3
            return # Stop execution of the script after logging a critical error
        }
    }
	"Remove" {
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
			Write-LogEntry -Value "Error updating hosts file" -Severity 3
			return # Stop execution of the script after logging a critical error
		}

		Write-LogEntry -Value "Uninstall of $AppName is complete" -Severity 1
	}

	default {
		Write-Output "Invalid mode: $Mode"
	}
}