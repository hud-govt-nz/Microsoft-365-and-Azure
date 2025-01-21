<#
.SYNOPSIS
    Printers Host File Update

.DESCRIPTION
    Script to install Printers Host File Update

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
$AppName = "Printer Host File Update"
$AppVersion = "1.0"
#$Installer = "<INSTALLERS>" # assumes the .exe or .msi installer is in the Files folder of the app package.
#$InstallArguments = "<INSTALLARGUMENTS>" # Optional
#$UninstallArguments = "<UNINSTALLARGUMENTS>" # Optional

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
            $backupFilePath = "$env:SystemRoot\System32\drivers\etc\Pre_Printers_hosts.old"

		# Create an empty hashtable to store IP addresses and hostnames
		$hostEntries = @{}

		# Add entries to the hashtable
		$hostEntries.Add("10.128.4.5","HUD-7WQ-L8-02-SOUTH.hud.govt.nz")
		$hostEntries.Add("10.128.4.6","HUD-7WQ-L8-01-NORTH.hud.govt.nz")
		$hostEntries.Add("10.128.4.7","HUD-7WQ-L9-01-NORTH.hud.govt.nz")
		$hostEntries.Add("10.128.4.8","HUD-7WQ-L7-01-NORTH.hud.govt.nz")
		$hostEntries.Add("10.128.4.9","HUD-7WQ-L6-02-SOUTH.hud.govt.nz")
		$hostEntries.Add("10.128.4.10","HUD-7WQ-L6-01-NORTH.hud.govt.nz")
		$hostEntries.Add("10.64.1.12","HUD-45QUEEN-L7-01.hud.govt.nz")

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
            Write-LogEntry -Value "error updating hosts file" -Severity 3
            return # Stop execution of the script after logging a critical error
        }

        Write-LogEntry -Value "Uninstall of $AppName is complete" -Severity 1
    }

    default {
        Write-Output "Invalid mode: $Mode"
    }
}
