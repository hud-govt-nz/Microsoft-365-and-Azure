<#
.SYNOPSIS
    Outlook Desktop Client Email Signature.

.DESCRIPTION
    Script deploys Outlook signature to local device. Uses Graph API to obtain users information from EntraID and fills in the template files.
    Scipt uses Graph API to obtain the currently logged on users EntraID details. This requires the deployment to be run in the USER context.

.PARAMETER Mode
    Sets the mode of operation. Supported modes are Install or Uninstall.

.EXAMPLE 
    powershell.exe -executionpolicy bypass -file .\appinstall.ps1 -Mode Install
    powershell.exe -executionpolicy bypass -file .\appinstall.ps1 -Mode Uninstall

.NOTES
    - AUTHOR: Ashley Forde
    - Version: 3.0
        - Pre-reqs: Microsoft Graph PS Module
        - Previous iteration of this deployment used the AzureAD PowerShell Module whereas this script leverages Graph

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
$AppName = "HUD Email Signature"
$AppValidationFile = "$validationFolderVar\$AppName.txt"
$AppVersion = "3.0"
$LogFileName = "$($AppName)_${Mode}_$Date.log"
$UPN = whoami /upn

# Begin Setup
Write-LogEntry -Value "Initiating setup process" -Severity 1
$SetupFolder = (New-Item -ItemType "directory" -Path $stagingFolderVar -Name $AppName -Force).FullName
Write-LogEntry -Value "Setup folder has been created at: $Setupfolder." -Severity 1

# Install/Uninstall
if ($Mode -eq "Install") {
	try {

		# Copy Signature template files to staging folder
		Copy-Item -Path "$PSScriptRoot\Files\*" -Destination $SetupFolder -Recurse -Force -ErrorAction Stop
		Write-LogEntry -Value "Signature files have been copied to $Setupfolder." -Severity 1

		try {
			# Remove old signature files
			$signaturePath = "$env:Appdata\Microsoft\Signatures"
			$oldSignaturePattern = "*HUD_Default_files*"
			$oldSignatureItems = Get-ChildItem -Path $signaturePath -Filter $oldSignaturePattern -Recurse

			if ($oldSignatureItems) {
				Write-LogEntry -Value "HUD_Default Signature items found on this device, removing" -Severity 1
				foreach ($item in $oldSignatureItems) {
					Remove-Item -Path $item.FullName -Recurse -Force -Confirm:$false
				}
			} else {
				Write-LogEntry -Value "No HUD_Default Signature items present on the device" -Severity 1
			}

			# Create signatures folder if it does not exist
			if (-not (Test-Path $signaturePath)) {
				try {
					New-Item -Path $signaturePath -ItemType Directory
					Write-LogEntry -Value "Signature path created at $signaturePath." -Severity 1

				} catch [System.Exception]{
					Write-LogEntry -Value "Error creating signature path. Errormessage: $($_.Exception.Message)" -Severity 3
				}
			}

			try {

				# Create signature files
				Connect-MgGraph -Scopes "Directory.Read.All,Directory.ReadWrite.All,User.Read,User.Read.All,User.ReadBasic.All,User.ReadWrite,User.ReadWrite.All" -NoWelcome
				Write-LogEntry -Value "Connecting to EntraID using Graph SDK" -Severity 1

				# Get user details from EntraID
				$User = Get-MgBetaUser -UserId $UPN
				Write-LogEntry -Value "Obtaining account details of $UPN" -Severity 1

				$Group = (Get-MgBetaUser -UserId $UPN | Select-Object AdditionalProperties -ExpandProperty AdditionalProperties).extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserOrganisationalGroup
				# Get all signature files
				$signatureFiles = Get-ChildItem -Path $SetupFolder

				foreach ($signatureFile in $signatureFiles) {
					if ($signatureFile.Name -like "*.htm" -or $signatureFile.Name -like "*.rtf" -or $signatureFile.Name -like "*.txt") {

						# Get file content with placeholder values
						$signatureFileContent = Get-Content -Path $signatureFile.FullName

						# Replace placeholder values
						$signatureFileContent = $signatureFileContent -replace "%DisplayName%",$User.DisplayName
						$signatureFileContent = $signatureFileContent -replace "%GivenName%",$User.GivenName
						$signatureFileContent = $signatureFileContent -replace "%Surname%",$User.Surname
						$signatureFileContent = $signatureFileContent -replace "%Mail%",$User.Mail
						$signatureFileContent = $signatureFileContent -replace "%Mobile%",$User.MobilePhone
						$signatureFileContent = $signatureFileContent -replace "%TelephoneNumber%",$User.BusinessPhones
						$signatureFileContent = $signatureFileContent -replace "%JobTitle%",$User.Jobtitle
						$signatureFileContent = $signatureFileContent -replace "%Department%",$User.Department
						$signatureFileContent = $signatureFileContent -replace "%Group%",$Group
						$signatureFileContent = $signatureFileContent -replace "%City%",$User.City
						$signatureFileContent = $signatureFileContent -replace "%Country%",$User.Country
						$signatureFileContent = $signatureFileContent -replace "%StreetAddress%",$User.StreetAddress
						$signatureFileContent = $signatureFileContent -replace "%PostalCode%",$User.PostalCode
						$signatureFileContent = $signatureFileContent -replace "%Country%",$User.Country
						$signatureFileContent = $signatureFileContent -replace "%State%",$User.State
						$signatureFileContent = $signatureFileContent -replace "%PhysicalDeliveryOfficeName%",$User.OfficeLocation

						# Set file content with values retrieved from Get-MgBetaUser command
						Set-Content -Path "$setupfolder\$($signatureFile.Name)" -Value $signatureFileContent -Force


					}
				}

				# Load new signature files to Outlook signatures folder
				Copy-Item -Path $SetupFolder\* -Destination $signaturePath -Recurse -Force -Confirm:$false
				Write-LogEntry -Value "Signature template copied to $signaturePath" -Severity 1

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
				Write-LogEntry -Value "Error creating new signature. Errormessage: $($_.Exception.Message)" -Severity 3
			}

		} catch [System.Exception]{
			Write-LogEntry -Value "Error removing old signature files. Errormessage: $($_.Exception.Message)" -Severity 3
		}

	} catch [System.Exception]{
		Write-LogEntry -Value "Error copying payload to staging folder. Errormessage: $($_.Exception.Message)" -Severity 3
	}
}

elseif ($Mode -eq "Uninstall") {
	# Purge signature files
	$signaturePath = "$env:Appdata\Microsoft\Signatures"
	$oldSignaturePattern = "*Default*"
	$oldSignatureItems = Get-ChildItem -Path $signaturePath -Filter $oldSignaturePattern -Recurse

	if ($oldSignatureItems) {
		Write-LogEntry -Value "Default Signature items found on this device, removing" -Severity 1
		foreach ($item in $oldSignatureItems) {
			Remove-Item -Path $item.FullName -Recurse -Force -Confirm:$false
		}
	} else {
		Write-LogEntry -Value "No Default Signature items present on the device" -Severity 1
	}

	# Add Validation File
	Remove-Item -Path $AppValidationFile -Force -Confirm:$false
	Write-LogEntry -Value "Validation file has been removed at $AppValidationFile" -Severity 1
	Write-LogEntry -Value "Uninstall of $AppName is complete" -Severity 1

	# Cleanup 
	if (Test-Path "$SetupFolder") {
		Remove-Item -Path "$SetupFolder" -Recurse -Force -ErrorAction Continue
		Write-LogEntry -Value "Cleanup completed successfully" -Severity 1
	}

}
