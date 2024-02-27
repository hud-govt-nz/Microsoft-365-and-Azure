<#
.SYNOPSIS
    WinSCP Application Deployment Script.

.DESCRIPTION
    Script to install WinSCP
    Runs install by downloading setup.exe file.

.NOTES
    - AUTHOR: Ashley Forde
    - Version: 1.0
    - Date: 12 Oct 2023
#>

#Region Parameters
[CmdletBinding()]
param(
	[Parameter(Mandatory = $false)]
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
$AppName = "WinSCP"
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

$SetupFolder = (New-Item -ItemType "directory" -Path $stagingFolderVar -Name $AppName -Force).FullName
Write-LogEntry -Value "Setup folder has been created at: $Setupfolder." -Severity 1

# Install/Uninstall M365 Apps
if ($Mode -eq "Install") {

	try {
		# Copy files to staging folder
		Copy-Item -Path "$PSScriptRoot\Installer\*" -Destination $SetupFolder -Recurse -Force -ErrorAction Stop
		Write-LogEntry -Value "Installer files have been copied to $Setupfolder." -Severity 1

		# create location for configuration files
		$config = New-Item -Path "$HomeFolder" -Name "09_WINSCP" -ItemType "directory" -Force -Confirm:$false
		Copy-Item -Path "$PSScriptRoot\Config\*" -Destination $config -Recurse -Force -ErrorAction Stop
		Write-LogEntry -Value "configuration files have been copied to $Configurations." -Severity 1

		# Test if file exists
		$SetupFilePath = (Join-Path -Path $SetupFolder -ChildPath "WinSCP-6.1.2-Setup.exe").ToString()
		if (-not (Test-Path $SetupFilePath)) { throw "Error: Setup file not found" }

		Write-LogEntry -Value "Setup file ready at $($SetupFilePath)" -Severity 1

		try {
			# Check setup.exe has valid file signature
			$VersionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($SetupFilePath).FileVersion
			Write-LogEntry -Value "Setup is running version $VersionInfo" -Severity 1

			try {
				# Run  setup.exe per configuration file
				Write-LogEntry -Value "Starting $Mode of WinSCP" -Severity 1
				[string]$Arguments = "/VERYSILENT /ALLUSERS /ALLUSERS"
				$Process = Start-Process $SetupFilePath -ArgumentList $Arguments -Wait -PassThru -ErrorAction Stop

				# Post Install Actions
				if ($Process.ExitCode -eq "0") {
					# add application configuration .ini to installation directory
					#Copy-Item -Path "C:\HUD\09_WINSCP\winSCP.ini" -Destination "C:\Program Files (x86)\WinSCP" -Force -Confirm:$false


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

	} catch [System.Exception]{ Write-LogEntry -Value "Error finding setup.exe Possible download error. Errormessage: $($_.Exception.Message)" -Severity 3 }

}
elseif ($Mode -eq "Uninstall") {

	try {
		# Run in-built Uninstaller
		Write-LogEntry -Value "Starting $Mode of WinSCP" -Severity 1
		$uninstallPath = "C:\Program Files (x86)\WinSCP\unins000.exe"
		[string]$Arguments = "/VERYSILENT /ALLUSERS /ALLUSERS"
		$Process = Start-Process $uninstallPath -ArgumentList $Arguments -Wait -PassThru -ErrorAction Stop

		# Post Install Actions
		if ($Process.ExitCode -eq "0") {
			# Remove leftover data
			Remove-Item -Path "C:\HUD\09_WINSCP" -Force -Recurse -Confirm:$false
			Write-LogEntry -Value "C:\HUD\09_WINSCP directory has been deleted" -Severity 1

			# Remove Validation File
			Remove-Item -Path $AppValidationFile -Force -Confirm:$false
			Write-LogEntry -Value "Validation file has been removed at $AppValidationFile" -Severity 1
			Write-LogEntry -Value "Uninstall of $AppName is complete" -Severity 1
		}

		try {
			#Set PS Drive for HKEY_Users and Obtain Current User System Identifier
			New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS -Scope Global
			Write-LogEntry -Value "PS_Drive HKU:\ has been created" -Severity 1

			# Obtain currently logged in user
			$currentUser = (Get-Process -IncludeUserName -Name explorer | Select-Object -First 1 | Select-Object -ExpandProperty UserName).Split("\")[1]
			Write-LogEntry -Value "Current signed in user is $currentUser" -Severity 1
			$Keys = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" -Recurse
			foreach ($Key in $Keys) {
				if (($key.GetValueNames() | ForEach-Object { $key.GetValue($_) }) -match $CurrentUser) {
					$sid = $key
				}
			}
			#SID for current user
			$UserSID = $sid.PSChildName
			Write-LogEntry -Value "Current user SID is $UserSID" -Severity 1

			# Registry Paths
			$regPath = "HKU:\$UserSID\Software\Martin Prikryl"

			# Add Custom dictionary to M365 Apps
			Remove-Item -Path $regPath -Force -Recurse -Confirm:$false
			Write-LogEntry -Value "Registry keys updated" -Severity 1

			Remove-PSDrive -Name HKU -Force -Confirm:$false
			Write-LogEntry -Value "PS Drive has been unmounted." -Severity 1

			# Cleanup 
			if (Test-Path "$SetupFolder") {
				Remove-Item -Path "$SetupFolder" -Recurse -Force -ErrorAction Continue
				Write-LogEntry -Value "Cleanup completed successfully" -Severity 1
			}

		} catch [System.Exception]{
			Write-LogEntry -Value "Error updating registry files. Errormessage: $($_.Exception.Message)" -Severity 3
		}

	} catch [System.Exception]{ Write-LogEntry -Value "Error completing uninstall. Errormessage: $($_.Exception.Message)" -Severity 3

	}
}
#ENDS
