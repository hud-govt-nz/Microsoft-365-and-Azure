<#
.SYNOPSIS
    Updates the currently signed in users Custom.DIC file for M365.

.DESCRIPTION
    The script appends to the users current Custom Dictionary file (DIC) used by Microsoft Office. 

.PARAMETER Mode
    Sets the mode of operation. Supported modes are Install or Uninstall.

.EXAMPLE 
    powershell.exe -executionpolicy bypass -file .\Add_To_Custom_Dictionary.ps1 -Mode Install
    powershell.exe -executionpolicy bypass -file .\Add_To_Custom_Dictionary.ps1 -Mode Uninstall

.NOTES
    - AUTHOR: Ashley Forde
    - Version: 1.0
        - Initial Release

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
$AppName = "Custom Dictionary Update"
$AppValidationFile = "$validationFolderVar\$AppName.txt"
$AppVersion = "1.0"
$LogFileName = "$($AppName)_${Mode}_$Date.log"

# Comment: The script initiates the install process, logs the initiation, and performs cleanup of the setup folder if it exists.
# Initate Install
Write-LogEntry -Value "Initiating update process" -Severity 1

# Attempt Cleanup of SetupFolder
if (Test-Path "$stagingFolderVar") {
	Remove-Item -Path "$($stagingFolderVar)\*" -Recurse -Force -ErrorAction SilentlyContinue

}

$SetupFolder = (New-Item -ItemType "directory" -Path $stagingFolderVar -Name $AppName -Force).FullName
Write-LogEntry -Value "Setup folder has been created at: $Setupfolder." -Severity 1

# Obtain Current User SID to enter into 
$currentUser = (Get-Process -IncludeUserName -Name explorer | Select-Object -First 1 | Select-Object -ExpandProperty UserName).Split("\")[1]
Write-LogEntry -Value "Currently signed in user is $currentUser." -Severity 1


# Install/Uninstall
if ($Mode -eq "Install") {

	try {
		# Copy Update DIC file to staging folder
		Copy-Item -Path "$PSScriptRoot\Payload\*" -Destination $SetupFolder -Recurse -Force -ErrorAction Stop
		Write-LogEntry -Value "Merge DIC files has been copied to $Setupfolder." -Severity 1

		# Define the file paths
		$dictionaryPath = "C:\Users\$currentUser\AppData\Roaming\Microsoft\UProof\CUSTOM.dic"
		$customWordsPath = "$PSScriptRoot\Payload\ToMerge.dic"

		# Copy existing Custom.DIC file and rename to .old
		if (Test-Path $dictionaryPath) {
			# Copy the existing dictionary file to a new file with the name custom.dic.old
			$Backup = Copy-Item -Path $dictionaryPath -Destination "C:\Users\$currentUser\AppData\Roaming\Microsoft\UProof\CUSTOM.dic.old"
			Write-LogEntry -Value "Copied existing CUSTOM.DIC file and renamed to CUSTOM.DIC.OLD, located at $Backup." -Severity 1

			# Now proceed to update the original dictionary file
			Get-Content $customWordsPath | ForEach-Object {
				Add-Content -Path $dictionaryPath -Value $_
				Write-LogEntry -Value "$_ has been added to Custom.DIC" -Severity 1
			}
		} else {
			Write-LogEntry -Value "Dictionary file not found." -Severity 1
		}
		try {
			# Load dictionary to office applications

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
			$regPath = "HKU:\$UserSID\Software\Microsoft\Shared Tools\Proofing Tools\1.0\Custom Dictionaries"
			$regPath2 = "HKU:\$UserSID\Software\Microsoft\Office\16.0\Common\Identity"

			# Obtain Current User Entra User Object ID 
			[string]$EntraID = (Get-ItemProperty -Path $regPath2).ConnectedAccountWamAad
			Write-LogEntry -Value "Current user Entra User Object ID is $EntraID" -Severity 1

			# Add Custom dictionary to M365 Apps
			Set-ItemProperty -Path $regPath -Name "1" -Value "CUSTOM.DIC"
			Set-ItemProperty -Path $regPath -Name "UpdateComplete" -Value 1 -Type DWord
			Set-ItemProperty -Path $regPath -Name "1_16_$($EntraID)_ADAL" -Value "C:\Users\$currentUser\AppData\Roaming\Microsoft\UProof\CUSTOM.dic"
			Set-ItemProperty -Path $regPath -Name "1_16_$($EntraID)_ADAL_state" -Value ([byte[]](1,0,0,0))
			Set-ItemProperty -Path $regPath -Name "1_16_$($EntraID)_ADAL_roamed" -Value ([byte[]](0,0,0,0))
			Set-ItemProperty -Path $regPath -Name "1_16_$($EntraID)_ADAL_external" -Value ([byte[]](0,0,0,0))

			Write-LogEntry -Value "Custom dictionary file mapped to users local office profile" -Severity 1

			Remove-PSDrive -Name HKU -Force -Confirm:$false
			Write-LogEntry -Value "PS Drive has been unmounted." -Severity 1


		} catch [System.Exception]{
			Write-LogEntry -Value "Error updating registry files. Errormessage: $($_.Exception.Message)" -Severity 3
		}

		# Add Validation File
		New-Item -ItemType File -Path $AppValidationFile -Force -Value $AppVersion
		Write-LogEntry -Value "Validation file has been created at $AppValidationFile" -Severity 1
		Write-LogEntry -Value "Install of $AppName is complete" -Severity 1

		# Cleanup 
		if (Test-Path "$SetupFolder") {
			Remove-Item -Path "$SetupFolder" -Recurse -Force -ErrorAction Continue
			Write-LogEntry -Value "Cleanup completed successfully" -Severity 1
		}

	} catch [System.Exception]{
		Write-LogEntry -Value "Error setting up custom dictionary. Errormessage: $($_.Exception.Message)" -Severity 3
	}
}

elseif ($Mode -eq "Uninstall") {
	# Define the file path
	$dictionaryPath = "C:\Users\$currentUser\AppData\Roaming\Microsoft\UProof\CUSTOM.dic"

	# Copy existing Custom.DIC file and rename to .old
	if (Test-Path $dictionaryPath) {
		# Copy the existing dictionary file to a new file with the name custom.dic.old
		Remove-Item -Path $dictionaryPath -Force -Confirm:$false
		Write-LogEntry -Value "Custom dictionary file removed. Restoring previous file." -Severity 1

		$Backup = "C:\Users\$currentUser\AppData\Roaming\Microsoft\UProof\CUSTOM.dic.old"
		Rename-Item -Path $Backup -NewName "CUSTOM.dic" -Force -Confirm:$false
		Write-LogEntry -Value "Old custom dictionary file restored" -Severity 1

		# Remove Validation File
		Remove-Item -Path $AppValidationFile -Force -Confirm:$false
		Write-LogEntry -Value "Validation file has been removed at $AppValidationFile" -Severity 1
		Write-LogEntry -Value "Uninstall of $AppName is complete" -Severity 1
	}

	try {
		# Load dictionary to office applications

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
		$regPath = "HKU:\$UserSID\Software\Microsoft\Shared Tools\Proofing Tools\1.0\Custom Dictionaries"
		$regPath2 = "HKU:\$UserSID\Software\Microsoft\Office\16.0\Common\Identity"

		# Obtain Current User Entra User Object ID 
		[string]$EntraID = (Get-ItemProperty -Path $regPath2).ConnectedAccountWamAad
		Write-LogEntry -Value "Current user Entra User Object ID is $EntraID" -Severity 1

		# Add Custom dictionary to M365 Apps
		Remove-ItemProperty -Path $regPath -Name "1"
		Remove-ItemProperty -Path $regPath -Name "1_16_$($EntraID)_ADAL"
		Remove-ItemProperty -Path $regPath -Name "1_16_$($EntraID)_ADAL_state"
		Remove-ItemProperty -Path $regPath -Name "1_16_$($EntraID)_ADAL_roamed"
		Remove-ItemProperty -Path $regPath -Name "1_16_$($EntraID)_ADAL_external"

		Write-LogEntry -Value "Registry keys updated" -Severity 1

		Remove-PSDrive -Name HKU -Force -Confirm:$false
		Write-LogEntry -Value "PS Drive has been unmounted." -Severity 1

	} catch [System.Exception]{
		Write-LogEntry -Value "Error updating registry files. Errormessage: $($_.Exception.Message)" -Severity 3
	}

	# Cleanup 
	if (Test-Path "$SetupFolder") {
		Remove-Item -Path "$SetupFolder" -Recurse -Force -ErrorAction Continue
		Write-LogEntry -Value "Cleanup completed successfully" -Severity 1
	}
}
