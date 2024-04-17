<#
.SYNOPSIS
    Updates the currently signed in users Custom.DIC file for M365.

.DESCRIPTION
    The script appends to the users current Custom Dictionary file (DIC) used by Microsoft Office.

.PARAMETER Mode
	Sets the mode of operation. Supported modes are Install or Uninstall.

.EXAMPLE 
    powershell.exe -executionpolicy bypass -file .\appinstall.ps1 -Mode Install
    powershell.exe -executionpolicy bypass -file .\appinstall.ps1 -Mode Uninstall

.NOTES
    - AUTHOR: Ashley Forde
    - Version: 1.0
    - Date: 25.03.2024

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
$AppName = "Office 365 Custom Dictionary"
$AppValidationFile = "$validationFolderVar\$AppName.txt"
$AppVersion = "2.0"
$LogFileName = "$($AppName)_${Mode}_$Date.log"
$SID = Get-CurrentUserSID

# Get current user information
$user=(Get-WmiObject -Class win32_computersystem).UserName.split('\')[1]
$AppData = "c:\users\$user\Appdata\Roaming"

# Begin Setup
Write-LogEntry -Value "Initiating setup process" -Severity 1

# Install/Uninstall
if ($Mode -eq "Install") {
    try {
		# Define file paths
		$dictionaryPath = "$AppData\Microsoft\UProof\CUSTOM.dic"
		$customWordsPath = "$PSScriptRoot\Files\ToMerge.dic"

		# Copy existing Custom.DIC file and rename to .old
		if (Test-Path $dictionaryPath) {
			# Copy the existing dictionary file to a new file with the name custom.dic.old
			$Backup = Copy-Item -Path $dictionaryPath -Destination "$AppData\Microsoft\UProof\CUSTOM.dic.old"
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
		Write-LogEntry -Value "Current user SID is $SID" -Severity 1

		# Registry Paths
		$regPath = "Registry::HKEY_USERS\$SID\Software\Microsoft\Shared Tools\Proofing Tools\1.0\Custom Dictionaries"
		$regPath2 = "Registry::HKEY_USERS\$SID\Software\Microsoft\Office\16.0\Common\Identity"

		# Obtain Current User Entra User Object ID 
		[string]$EntraID = (Get-ItemProperty -Path $regPath2).ConnectedAccountWamAad
		Write-LogEntry -Value "Current user Entra User Object ID is $EntraID" -Severity 1

		# Add Custom dictionary to M365 Apps
		Set-ItemProperty -Path $regPath -Name "1" -Value "CUSTOM.DIC"
		Set-ItemProperty -Path $regPath -Name "UpdateComplete" -Value 1 -Type DWord
		Set-ItemProperty -Path $regPath -Name "1_16_$($EntraID)_ADAL" -Value "$AppData\Microsoft\UProof\CUSTOM.dic"
		Set-ItemProperty -Path $regPath -Name "1_16_$($EntraID)_ADAL_state" -Value ([byte[]](1,0,0,0))
		Set-ItemProperty -Path $regPath -Name "1_16_$($EntraID)_ADAL_roamed" -Value ([byte[]](0,0,0,0))
		Set-ItemProperty -Path $regPath -Name "1_16_$($EntraID)_ADAL_external" -Value ([byte[]](0,0,0,0))
		Write-LogEntry -Value "Custom dictionary file mapped to users local office profile" -Severity 1

	} catch [System.Exception]{
		Write-LogEntry -Value "Error updating registry files. Errormessage: $($_.Exception.Message)" -Severity 3
	}

	# Add Validation File
	New-Item -ItemType File -Path $AppValidationFile -Force -Value $AppVersion
	Write-LogEntry -Value "Validation file has been created at $AppValidationFile" -Severity 1
	Write-LogEntry -Value "Install of $AppName is complete" -Severity 1

	} catch [System.Exception]{
		Write-LogEntry -Value "Error setting up custom dictionary. Errormessage: $($_.Exception.Message)" -Severity 3
	}
}

elseif ($Mode -eq "Uninstall") {
	# Define the file path
	$dictionaryPath = "$AppData\Microsoft\UProof\CUSTOM.dic"

	# Copy existing Custom.DIC file and rename to .old
	if (Test-Path $dictionaryPath) {
		# Copy the existing dictionary file to a new file with the name custom.dic.old
		Remove-Item -Path $dictionaryPath -Force -Confirm:$false
		Write-LogEntry -Value "Custom dictionary file removed. Restoring previous file." -Severity 1

		$Backup = "$AppData\Microsoft\UProof\CUSTOM.dic.old"
		Rename-Item -Path $Backup -NewName "CUSTOM.dic" -Force -Confirm:$false
		Write-LogEntry -Value "Old custom dictionary file restored" -Severity 1

		# Remove Validation File
		Remove-Item -Path $AppValidationFile -Force -Confirm:$false
		Write-LogEntry -Value "Validation file has been removed at $AppValidationFile" -Severity 1
		Write-LogEntry -Value "Uninstall of $AppName is complete" -Severity 1
	}

	try {
		# Obtain currently logged in user
		Write-LogEntry -Value "Current user SID is $SID" -Severity 1

		# Registry Paths
		$regPath = "Registry::HKEY_USERS\$SID\Software\Microsoft\Shared Tools\Proofing Tools\1.0\Custom Dictionaries"
		$regPath2 = "Registry::HKEY_USERS\$SID\Software\Microsoft\Office\16.0\Common\Identity"

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

	} catch [System.Exception]{
		Write-LogEntry -Value "Error updating registry files. Errormessage: $($_.Exception.Message)" -Severity 3
	}

}
