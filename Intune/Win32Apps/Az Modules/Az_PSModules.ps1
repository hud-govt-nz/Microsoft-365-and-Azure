<#
.SYNOPSIS
    Install/Uninstall of PS modules with name Az_PSModules* .

.DESCRIPTION
    Script installs the PowerShell Modules for all modules that begin with name Az_PSModules*

.PARAMETER Mode
    Sets the mode of operation. Supported modes are Install or Uninstall.

.EXAMPLE 
    powershell.exe -executionpolicy bypass -file .\Az_PSModules Modules.ps1 -Mode Install
    powershell.exe -executionpolicy bypass -file .\Az_PSModules Modules.ps1 -Mode Uninstall

.NOTES
    - AUTHOR: Ashley Forde
    - Script can be run in both SYSTEM and USER Context however if running as USER then they need local administrator rights. 
    - Version: 1.0
        - Initial Release
    
#>

#Region Parameters
[CmdletBinding()]
param(
	[Parameter(Mandatory = $true)]
	[ValidateNotNullOrEmpty()]
	[ValidateSet("Install","Uninstall")]
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
$AppName = "Az_PSModules"
$AppValidationFile = "$validationFolderVar\$AppName.txt"
$AppVersion = "1.0"
$LogFileName = "$($AppName)_${Mode}_$Date.log"

# Comment: The script initiates the install process, logs the initiation, and performs cleanup of the setup folder if it exists.
# Initate Install
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

			# Force Uninstall existing Modules

			# Define the modules to be installed
			$modules = Find-Module -Name "Az.*"

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
				if (Get-Module -ListAvailable -Name "Az*") {
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
		# Get all available modules that match the pattern 'Az_PSModules*'
		$modules = Get-Module -ListAvailable -Name "Az.*"

		# Check if modules were found
		if ($modules.count -eq 0) {
			Write-LogEntry -Value "No modules matching 'Az*' were found." -Severity 1
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
		#Remove-Item -Path $AppValidationFile -Force -Confirm:$false
		Write-LogEntry -Value "Validation file has been removed at $AppValidationFile" -Severity 1

	} catch [System.Exception]{
		Write-LogEntry -Value "Error retrieving modules. Error message: $($_.Exception.Message)" -Severity 3
	}

	# Complete
	Write-LogEntry -Value "Uninstall of $AppName is complete" -Severity 1
}
