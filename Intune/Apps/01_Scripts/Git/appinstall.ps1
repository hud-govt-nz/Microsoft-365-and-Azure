<#
.SYNOPSIS
    Git.

.DESCRIPTION
    Script to install Git.

.PARAMETER Mode
Sets the mode of operation. Supported modes are Install or Uninstall.

.EXAMPLE 
    powershell.exe -executionpolicy bypass -file .\appinstall.ps1 -Mode Install
    powershell.exe -executionpolicy bypass -file .\appinstall.ps1 -Mode Uninstall

.NOTES
    - AUTHOR: Ashley Forde
    - Version: 2.0
    - Date: 18.4.2024
    - NOTES
        - Refactored to use functions.ps1
        - Latest version packaged.

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
$AppName = "Git"
$AppValidationFile = "$validationFolderVar\$AppName.txt"
$AppVersion = "2.44.0"
$LogFileName = "$($AppName)_${Mode}_$Date.log"

# Begin Setup
Write-LogEntry -Value "Initiating setup process" -Severity 1
$SetupFolder = (New-Item -ItemType "directory" -Path $stagingFolderVar -Name $AppName -Force).FullName
Write-LogEntry -Value "Setup folder has been created at: $Setupfolder." -Severity 1

# Install/Uninstall
if ($Mode -eq "Install") {
    try {
        # Copy files to staging folder
        Copy-Item -Path "$PSScriptRoot\Installer\*" -Destination $SetupFolder -Recurse -Force -ErrorAction Stop
        Write-LogEntry -Value "Setup files have been copied to $Setupfolder." -Severity 1

		# Test if there is a setup file
		$SetupFilePath = (Join-Path -Path $SetupFolder -ChildPath "Git-2.44.0-64-bit.exe").ToString()

		if (-not (Test-Path $SetupFilePath)) { 
			throw "Error: Setup file not found" 
		}
		Write-LogEntry -Value "Setup file ready at $($SetupFilePath)" -Severity 1

		try {
			# Run setup with custom arguments and create validation file
			Write-LogEntry -Value "Starting $Mode of Git" -Severity 1
			[string]$Arguments = "/VERYSILENT /NORESTART /NOCANCEL"
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
		# Uninstall App
		$uninstall_command = Get-ChildItem -Path "C:\Program Files\Git\" -Filter "unins00*" | Where-Object { $_.Extension -eq ".exe" } | Select-Object -ExpandProperty FullName
		[string]$uninstall_args = '/SILENT'
		$uninstallProcess = Start-Process $uninstall_command -ArgumentList $uninstall_args -PassThru -Wait -ErrorAction stop

		# Post Uninstall Actions
		if ($uninstallProcess.ExitCode -eq "0") {		
			# Delete validation file
			Remove-Item -Path $AppValidationFile -Force -Confirm:$false
			Write-LogEntry -Value "Validation file has been removed at $AppValidationFile" -Severity 1

			# Cleanup 
			if (Test-Path "$SetupFolder") {
				Remove-Item -Path "$SetupFolder" -Recurse -Force -ErrorAction Continue
				Write-LogEntry -Value "Cleanup completed successfully" -Severity 1
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
