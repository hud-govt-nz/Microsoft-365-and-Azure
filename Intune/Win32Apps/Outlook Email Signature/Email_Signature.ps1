<#
.SYNOPSIS
    Outlook Desktop Client Email Signature.

.DESCRIPTION
    Script deploys Outlook signature to local device. Uses Graph API to obtain users information from EntraID and fills in the template files.
    Scipt uses Graph API to obtain the currently logged on users EntraID details. This requires the deployment to be run in the USER context.

.PARAMETER Mode
    Sets the mode of operation. Supported modes are Install or Uninstall.

.EXAMPLE 
    powershell.exe -executionpolicy bypass -file .\Email_Signature.ps1 -Mode Install
    powershell.exe -executionpolicy bypass -file .\Email_Signature.ps1 -Mode Uninstall

.NOTES
    - AUTHOR: Ashley Forde
    - Version: 2.0
        - Pre-reqs: Microsoft Graph PS Module
        - Previous iteration of this deployment used the AzureAD PowerShell Module whereas this script leverages Graph
        - Roaming signature feature in Outlook Click-to-Run has been disabled https://techcommunity.microsoft.com/t5/outlook/outlook-roaming-signature-vs-signatures-on-this-device/m-p/3672866

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
function Set-DisableRoamingSignatures {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[ValidateSet('Add','Remove')]
		[string]$Action,

		[Parameter(Mandatory = $false)]
		[ValidateSet(0,1)]
		[int]$ValueData
	)

	$Hive = "HKEY_CURRENT_USER"
	$KeyPath = "Software\Microsoft\Office\16.0\Outlook\Setup"
	$ValueName = "DisableRoamingSignatures"
	$ValueType = "DWORD"

	if ($Action -eq "Add") {
		New-ItemProperty -Path "Registry::$Hive\$KeyPath" -Name $ValueName -PropertyType $ValueType -Value $ValueData -Force
		Write-Verbose -Message "Registry value added successfully."
	}
	elseif ($Action -eq "Remove") {
		Remove-ItemProperty -Path "Registry::$Hive\$KeyPath" -Name $ValueName -Force
		Write-Verbose -Message "Registry value removed successfully."
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
$AppName = "Outlook_Email_Signature"
$AppValidationFile = "$validationFolderVar\$AppName.txt"
$AppVersion = "2.0"
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
		# Disable roaming signatures
		Set-DisableRoamingSignatures -Action Add -ValueData 1 -Verbose
		Write-LogEntry -Value "Roaming Signatures has been disabled in Outlook" -Severity 1

		# Copy Signature template files to staging folder
		Copy-Item -Path "$PSScriptRoot\Payload\*" -Destination $SetupFolder -Recurse -Force -ErrorAction Stop
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
				$upn = Whoami /upn
				Connect-MgGraph -NoWelcome
				Write-LogEntry -Value "Connecting to EntraID using Graph SDK" -Severity 1


				# Get user details from EntraID
				$User = Get-MgBetaUser -UserId $upn
				Write-LogEntry -Value "Obtaining account details of $upn" -Severity 1


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
	# Enable roaming signatures
	Set-DisableRoamingSignatures -Action Remove -Verbose
	Write-LogEntry -Value "Roaming Signatures has been re-enabled in Outlook" -Severity 1

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
