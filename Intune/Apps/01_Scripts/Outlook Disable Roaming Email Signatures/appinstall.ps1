<#
.SYNOPSIS
    Outlook Desktop Client Disable Roaming Email Signatures.

.DESCRIPTION
    Script runs in SYSTEM context, retreives current users SID and makes changes to their Outlook Client configuration. Sets DWORD value in "HKCU:\Software\Microsoft\Office\16.0\Outlook\Setup"
	for "DisableRoamingSignatures" to either 1 or 0.
	
	1 = Disable Roaming Email Signatures
	0 = Enable Roaming Email Signatures

.PARAMETER Mode
    Sets the mode of operation. Supported modes are Install or Uninstall.

.EXAMPLE 
    powershell.exe -executionpolicy bypass -file .\appinstall.ps1 -Mode Install
    powershell.exe -executionpolicy bypass -file .\appinstall.ps1 -Mode Uninstall

.NOTES
    - AUTHOR: Ashley Forde
    - Version: 1.0

#>

# Parameters
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
$AppName = "Disable Roaming Email Signatures"
$AppValidationFile = "$validationFolderVar\$AppName.txt"
$AppVersion = "1.0"
$LogFileName = "$($AppName)_${Mode}_$Date.log"

# Begin Setup
Write-LogEntry -Value "Initiating setup process" -Severity 1

# Install/Uninstall
if ($Mode -eq "Install") {
	try {

		# Disable Roaming Signatures
		Set-DisableRoamingSignatures -Action Add -ValueData 1
		Write-LogEntry -Value "Roaming Signatures have been disabled." -Severity 1

		# Add Validation File
		New-Item -ItemType File -Path $AppValidationFile -Force -Value $AppVersion
		Write-LogEntry -Value "Validation file has been created at $AppValidationFile" -Severity 1

	} catch [System.Exception]{
		Write-LogEntry -Value "Error disabling roaming signatures. Errormessage: $($_.Exception.Message)" -Severity 3
	}

	Write-LogEntry -Value "Install of $AppName is complete" -Severity 1
}

elseif ($Mode -eq "Uninstall") {

	try {
		# Enable Roaming Signatures
		Set-DisableRoamingSignatures -Action Add -ValueData 0
		Write-LogEntry -Value "Roaming Signatures have been enabled." -Severity 1
		
		# Add Validation File
		Remove-Item -Path $AppValidationFile -Force -Confirm:$false
		Write-LogEntry -Value "Validation file has been removed at $AppValidationFile" -Severity 1

	} catch [System.Exception]{
		Write-LogEntry -Value "Error enabling roaming signatures. Errormessage: $($_.Exception.Message)" -Severity 3
	}
	
	Write-LogEntry -Value "Uninstall of $AppName is complete" -Severity 1

}










