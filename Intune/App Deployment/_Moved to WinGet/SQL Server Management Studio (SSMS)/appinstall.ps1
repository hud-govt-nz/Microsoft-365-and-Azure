<#
.SYNOPSIS
    SQL Server Management Studio (SSMS).

.DESCRIPTION
    Script to install SQL Server Management Studio (SSMS).

.PARAMETER Mode
Sets the mode of operation. Supported modes are Install or Uninstall.

.EXAMPLE 
    powershell.exe -executionpolicy bypass -file .\appinstall.ps1 -Mode Install
    powershell.exe -executionpolicy bypass -file .\appinstall.ps1 -Mode Uninstall

.NOTES
    - AUTHOR: Ashley Forde
    - Version: 1.0
    - Date: 16.04.2024
#>

#Region Parameters
[CmdletBinding()]
param(
	[Parameter(Mandatory = $false)]
	[ValidateNotNullOrEmpty()]
	[ValidateSet("Install","Uninstall")]
	[string]$Mode
)

# Reference functions.ps1
. "$PSScriptRoot\functions.ps1"

# Initialize Directories
$folderpaths = Initialize-Directories -HomeFolder C:\HUD\

# Assign the returned values to individual variables
$stagingFolderVar = $folderPaths.StagingFolder
$logsFolderVar = $folderPaths.LogsFolder
$validationFolderVar = $folderPaths.ValidationFolder
$Date = Get-Date -Format "MM-dd-yyyy"
$AppName = "SQL Server Management Studio (SSMS)"
$AppValidationFile = "$validationFolderVar\$AppName.txt"
$AppVersion = "20.1"
$LogFileName = "$($AppName)_${Mode}_$Date.log"

# Begin Setup
Write-LogEntry -Value "Initiating setup process" -Severity 1
$SetupFolder = (New-Item -ItemType "directory" -Path $stagingFolderVar -Name $AppName -Force).FullName
Write-LogEntry -Value "Setup folder has been created at: $Setupfolder." -Severity 1

# Install/Uninstall
if ($Mode -eq "Install") {
    try {
        # Copy files to staging folder
        Copy-Item -Path "$PSScriptRoot\Files\*" -Destination $SetupFolder -Recurse -Force -ErrorAction Stop
        Write-LogEntry -Value "Setup files have been copied to $Setupfolder." -Severity 1

		# Test if there is a setup file
		$SetupFilePath = (Join-Path -Path $SetupFolder -ChildPath "SSMS-Setup-ENU.exe").ToString()

		if (-not (Test-Path $SetupFilePath)) { 
			throw "Error: Setup file not found" 
		}
		Write-LogEntry -Value "Setup file ready at $($SetupFilePath)" -Severity 1

		try {
			# Run setup with custom arguments and create validation file
			Write-LogEntry -Value "Starting $Mode of SQL Server Management Studio (SSMS)" -Severity 1
			[string]$Arguments = "/install /quiet /norestart"
			$Process = Start-Process $SetupFilePath -ArgumentList $Arguments -Wait -PassThru -ErrorAction Stop

			# Post Install Actions
			if ($Process.ExitCode -eq "0") {
				# Create validation file
				New-Item -ItemType File -Path $AppValidationFile -Force -Value $AppVersion
				Write-LogEntry -Value "Validation file has been created at $AppValidationFile" -Severity 1
				Write-LogEntry -Value "Install of $AppName is complete" -Severity 1
			} else {
				Write-LogEntry -Value "Install of $AppName failed with ExitCode: $($Process.ExitCode)" -Severity 3
			}

			# Remove previous versions if they exist
			$OldApp = Get-InstalledApps -App "Microsoft SQL Server Management Studio*" | Where-Object {$_.DisplayVersion -ne "20.1.10.0"} | Select-Object -First 1
			Start-Process $OldApp.BundleCachePath -ArgumentList '/uninstall /quiet /norestart' -PassThru -Wait -ErrorAction stop

			# Cleanup 
			if (Test-Path "$SetupFolder") {
				Remove-Item -Path "$SetupFolder" -Recurse -Force -ErrorAction Continue
				Write-LogEntry -Value "Cleanup completed successfully" -Severity 1
			}

			} catch {
				Write-LogEntry -Value "Error running installer. Errormessage: $($_.Exception.Message)" -Severity 3
				return # Stop execution of the script after logging a critical error
			}
		} catch [System.Exception]{ Write-LogEntry -Value "Error preparing installation $FileName $($mode). Errormessage: $($_.Exception.Message)" -Severity 3 }

} elseif ($Mode -eq "Uninstall") {
	try {
		$MyApp = Get-InstalledApps -App "Microsoft SQL Server Management Studio*"

		# Uninstall App
		$uninstallProcess = Start-Process $MyApp.BundleCachePath -ArgumentList '/uninstall /quiet /norestart' -PassThru -Wait -ErrorAction stop

		# Post Uninstall Actions
		if ($uninstallProcess.ExitCode -eq "0") {
			# Delete validation file
			try {
				Remove-Item -Path $AppValidationFile -Force -Confirm:$false
				Write-LogEntry -Value "Validation file has been removed at $AppValidationFile" -Severity 1

				# Cleanup 
				if (Test-Path "$SetupFolder") {
					Remove-Item -Path "$SetupFolder" -Recurse -Force -ErrorAction Continue
					Write-LogEntry -Value "Cleanup completed successfully" -Severity 1
				}
			} catch [System.Exception] {
				Write-LogEntry -Value "Error deleting validation file. Errormessage: $($_.Exception.Message)" -Severity 3
			}
		} else {
			throw "Uninstallation failed with exit code $($uninstallProcess.ExitCode)"
		}
	} catch [System.Exception] {
		Write-LogEntry -Value "Error completing uninstall. Errormessage: $($_.Exception.Message)" -Severity 3
		throw "Uninstallation halted due to an error"
	}

	Write-LogEntry -Value "Uninstall of $AppName is complete" -Severity 1
}
