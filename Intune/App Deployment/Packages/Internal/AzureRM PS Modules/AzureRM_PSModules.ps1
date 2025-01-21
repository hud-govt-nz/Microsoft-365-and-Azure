<#
.SYNOPSIS
    Install/Uninstall of PS modules with name AzureRM* .

.DESCRIPTION
    Script installs the PowerShell Modules for all modules that begin with name AzureRM*

.PARAMETER Mode
    Sets the mode of operation. Supported modes are Install or Uninstall.

.EXAMPLE 
    powershell.exe -executionpolicy bypass -file .\AzureRM_PSModules.ps1 -Mode Install
    powershell.exe -executionpolicy bypass -file .\AzureRM_PSModules.ps1 -Mode Uninstall

.NOTES
    - AUTHOR: Ashley Forde
    - Script can be run in both SYSTEM and USER Context however if running as USER then they need local administrator rights. 
    - Version: 2.0
        - Refactored script to use functions.ps1



#>

#Region Parameters
[CmdletBinding()]
param(
	[Parameter(Mandatory = $true)]
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
$AppName = "AzureRM_PSModules"
$AppValidationFile = "$validationFolderVar\$AppName.txt"
$AppVersion = "2.0"
$LogFileName = "$($AppName)_${Mode}_$Date.log"

# Begin Setup
Write-LogEntry -Value "Initiating setup process" -Severity 1

# Attempt Cleanup of SetupFolder
if (Test-Path "$stagingFolderVar") {
	Remove-Item -Path "$($stagingFolderVar)\*" -Recurse -Force -ErrorAction SilentlyContinue
}

# Install/Uninstall Modules
if ($Mode -eq "Install") {
	# Ensure prerequisites are installed
	try {
		# NuGet
		Install-PackageProvider -Name NuGet -Scope AllUsers -Force -Confirm:$false -ForceBootstrap

		# PowerShellGet
		Install-Module -Name PowerShellGet -Force -Confirm:$false
		Install-Module -Name PackageManagement -Force -Confirm:$false

		Write-LogEntry -Value "Prerequisites have been installed/updated successfully." -Severity 1

		try {
			# Define the modules to be installed
			$modules = Find-Module -Name "AzureRM*"

			# Install the defined modules
			foreach ($module in $modules) {
				if (-not (Get-Module -ListAvailable -Name $module.Name)) {
					Install-Module -Name $module.Name -Scope AllUsers -Force -AllowClobber -Verbose
					Write-LogEntry -Value "$($module.Name) module has been installed successfully." -Severity 1
				} else {
					Write-LogEntry -Value "$($module.name) module is already installed." -Severity 1
				}
			}

			try {
				if (Get-Module -ListAvailable -Name "AzureRM*") {
					# Create validation file if script completes without issue
					New-Item -ItemType File -Path $AppValidationFile -Force -Value $AppVersion
					Write-LogEntry -Value "Validation file has been created at $AppValidationFile" -Severity 1
				} else {
					Write-LogEntry -Value "One or both of the modules are not available." -Severity 1
				}

			} catch [System.Exception]{
				Write-LogEntry -Value "Error verifying module install, Errormessage: $($_.Exception.Message)" -Severity 3
				return
			}

		} catch [System.Exception]{
			Write-LogEntry -Value "Error installing Module, Errormessage: $($_.Exception.Message)" -Severity 3
			return
		}
	}
	catch [System.Exception]{
		Write-LogEntry -Value "Error installing prerequisites Package Provider, Errormessage: $($_.Exception.Message)" -Severity 3
		return
	}

}
elseif ($Mode -eq "Uninstall") {
	try {
		# Get all available modules that match the pattern 'AzureRM*'
		$modules = Get-Module -ListAvailable -Name "AzureRM*"

		# Check if modules were found
		if ($modules.count -eq 0) {
			Write-LogEntry -Value "No modules matching 'AzureRM*' were found." -Severity 1
			return
		}

		# Loop through each module in the array
		foreach ($module in $modules) {
			try {
				# Uninstall the module
				Uninstall-Module -Name $module.Name -Force -Confirm:$false -ErrorAction SilentlyContinue

				# Display a confirmation message
				Write-LogEntry -Value "$($module.Name) module has been uninstalled." -Severity 1
			} catch [System.Exception]{
				Write-LogEntry -Value "Error uninstalling $($module.Name) module. Error message: $($_.Exception.Message)" -Severity 3
				return
			}
		}

		# Remove validation files
		Remove-Item -Path $AppValidationFile -Force -Confirm:$false
		Write-LogEntry -Value "Validation file has been removed at $AppValidationFile" -Severity 1

	} catch [System.Exception]{
		Write-LogEntry -Value "Error retrieving modules. Error message: $($_.Exception.Message)" -Severity 3
	}

	# Complete
	Write-LogEntry -Value "Uninstall of $AppName is complete" -Severity 1
}
