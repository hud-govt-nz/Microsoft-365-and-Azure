<#
.SYNOPSIS
    Install/Uninstall of all currently installed PS Modules. [use with caution]

.DESCRIPTION
    Script installs the PowerShell Modules for all currently installed PS modules*

.EXAMPLE 
    powershell.exe -executionpolicy bypass -file .\Purge_Installed_PSModules.ps1
    powershell.exe -executionpolicy bypass -file .\Purge_Installed_PSModules.ps1

.NOTES
    - AUTHOR: Ashley Forde
    - Script can be run in both SYSTEM and USER Context however if running as USER then they need local administrator rights. 
    - Version: 1.0
        - Initial Release
    
#>

#Region Parameters
[CmdletBinding()]
param(
	[Parameter(Mandatory = $false)]
	[string]$Action
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
$AppName = "Purge_Installed_Modules"
$AppValidationFile = "$validationFolderVar\$AppName.txt"
$AppVersion = "1.0"
$LogFileName = "$($AppName)_$Date.log"

# Comment: The script initiates the install process, logs the initiation, and performs cleanup of the setup folder if it exists.
# Initate Install
Write-LogEntry -Value "Initiating setup process" -Severity 1

# Attempt Cleanup of SetupFolder
if (Test-Path "$stagingFolderVar") {
	Remove-Item -Path "$($stagingFolderVar)\*" -Recurse -Force -ErrorAction SilentlyContinue
}


try {
	# Get all available modules that match the pattern 'Microsoft.Graph*'
	$modules = Get-InstalledModule

	# Check if modules were found
	if ($modules.count -eq 0) {
		Write-LogEntry -Value "No modules installed' were found." -Severity 1
		return
	}

	# Loop through each module in the array
	foreach ($module in $modules) {
		try {
			# Check if the current module is PowerShellGet,PSReadline, or PackageManagement and skip uninstallation if it is
			if ($module.Name -ne "PowerShellGet" -and $module.Name -ne "PSReadline" -and $module.Name -ne "PackageManagement") {
				# Uninstall the module
				Uninstall-Module -Name $module.Name -Force -Confirm:$false -ErrorAction Stop

				# Display a confirmation message
				Write-LogEntry -Value "$($module.Name) module has been uninstalled." -Severity 1
			} else {
				# Log skipping of module
				Write-LogEntry -Value "$($module.Name) module was not uninstalled because it is excluded by the script's conditions." -Severity 1
			}
		} catch [System.Exception]{
			Write-LogEntry -Value "Error uninstalling $($module.Name) module. Error message: $($_.Exception.Message)" -Severity 3
			return
		}
	}

	# Create validation file if script completes without issue
	New-Item -ItemType File -Path $AppValidationFile -Force -Value $AppVersion
	Write-LogEntry -Value "Validation file has been created at $AppValidationFile" -Severity 1

} catch [System.Exception]{
	Write-LogEntry -Value "Error retrieving modules. Error message: $($_.Exception.Message)" -Severity 3
}

