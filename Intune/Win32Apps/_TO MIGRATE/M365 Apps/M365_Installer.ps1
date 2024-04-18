<#
.SYNOPSIS
    M365 Application Deployment Script.

.DESCRIPTION
    Script to install custom M365 App deployments using .XML files and parameter switches depending on scenario.
    Runs install by downloading Evergreen setup.exe file.

.PARAMETER Type
    Sets the type of installation. Supported types are OfficeStd, OfficeInclAccess, Visio, Project, LangPack, ProofTools.
    These configurations are stored in the .\configuration folder inside the Win32app and are pushed out to the device depending on which configuration is selected.

    Note all configurations skip MS Teams installation as that is governed by a separate Win32app deployment.
.PARAMETER Mode
    Sets the mode of operation. Supported modes are Install or Uninstall.

.EXAMPLE 
    powershell.exe -executionpolicy bypass -file .\M365_Installer.ps1 -Type OfficeStd -Mode Install
    powershell.exe -executionpolicy bypass -file .\M365_Installer.ps1 -Type OfficeStd -Mode Uninstall

.NOTES
    - AUTHOR: Ashley Forde
    - Version: 1.0
    - Special Thanks to https://github.com/MSEndpointMgr/M365Apps which provded the starting point for this deployment. 
        - Improvements made:
            - Included uninstall logic for all app variations.
            - 1 script used to manage Office, Visio, Project, Language Pack, Proofing tool install and uninstalls. 
            - Validation/detection script based on script completion where a validation file is generated only if the install completes with exit code 0 (successful)
            - Removed XMLURL as it was not needed in this scenario


#>

#Region Parameters
[CmdletBinding()]
param(
	[Parameter(Mandatory = $false)]
	[ValidateNotNullOrEmpty()]
	[ValidateSet("OfficeStd","OfficeInclAccess","Visio","Project","LangPack","ProofTools")]
	[string]$Type,
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
function Start-DownloadFile {
	param(
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$URL,

		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$Path,

		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$Name
	)
	begin {
		# Construct WebClient object
		$WebClient = New-Object -TypeName System.Net.WebClient
	}
	process {
		# Create path if it doesn't exist
		if (-not (Test-Path -Path $Path)) {
			New-Item -Path $Path -ItemType Directory -Force | Out-Null
		}

		# Start download of file
		$WebClient.DownloadFile($URL,(Join-Path -Path $Path -ChildPath $Name))
	}
	end {
		# Dispose of the WebClient object
		$WebClient.Dispose()
	}
}
function Invoke-FileCertVerification {
	param(
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$FilePath
	)
	# Get a X590Certificate2 certificate object for a file
	$Cert = (Get-AuthenticodeSignature -FilePath $FilePath).SignerCertificate
	$CertStatus = (Get-AuthenticodeSignature -FilePath $FilePath).Status
	if ($Cert) {
		#Verify signed by Microsoft and Validity
		if ($cert.Subject -match "O=Microsoft Corporation" -and $CertStatus -eq "Valid") {
			#Verify Chain and check if Root is Microsoft
			$chain = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Chain
			$chain.Build($cert) | Out-Null
			$RootCert = $chain.ChainElements | ForEach-Object { $_.Certificate } | Where-Object { $PSItem.Subject -match "CN=Microsoft Root" }
			if (-not [string ]::IsNullOrEmpty($RootCert)) {
				#Verify root certificate exists in local Root Store
				$TrustedRoot = Get-ChildItem -Path "Cert:\LocalMachine\Root" -Recurse | Where-Object { $PSItem.Thumbprint -eq $RootCert.Thumbprint }
				if (-not [string]::IsNullOrEmpty($TrustedRoot)) {
					Write-LogEntry -Value "Verified setupfile signed by : $($Cert.Issuer)" -Severity 1
					return $True
				}
				else {
					Write-LogEntry -Value "No trust found to root cert - aborting" -Severity 2
					return $False
				}
			}
			else {
				Write-LogEntry -Value "Certificate chain not verified to Microsoft - aborting" -Severity 2
				return $False
			}
		}
		else {
			Write-LogEntry -Value "Certificate not valid or not signed by Microsoft - aborting" -Severity 2
			return $False
		}
	}
	else {
		Write-LogEntry -Value "Setup file not signed - aborting" -Severity 2
		return $False
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
$AppName = "Microsoft_365_Apps_${Type}"
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

switch -Wildcard ($Type) {
	"OfficeStd" {
		$TypeStr = "OfficeStd"
	}
	"OfficeInclAccess" {
		$TypeStr = "OfficeInclAccess"
	}
	"Visio" {
		$TypeStr = "Visio"
	}
	"Project" {
		$TypeStr = "Project"
	}
	"LangPack" {
		$TypeStr = "LangPack"
	}
	"ProofTools" {
		$TypeStr = "ProofTools"
	}
}

# Install/Uninstall M365 Apps
if ($Mode -eq "Install") {
	$FileName = "${TypeStr}_Install.xml"
}
elseif ($Mode -eq "Uninstall") {
	$FileName = "${TypeStr}_Uninstall.xml"
}

# Run install or uninstall depending on XML configuration.

try {
	# Set TLS1.2
	[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls13

	# Download latest office setup.exe
	$SetupFile = "setup.exe"
	$SetupEvergreenUrl = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
	Write-LogEntry -Value "Attempting to download latest Office setup executable" -Severity 1
	Start-DownloadFile -URL $SetupEvergreenUrl -Path $SetupFolder -Name $SetupFile

	try {
		# Test if file exists
		$SetupFilePath = (Join-Path -Path $SetupFolder -ChildPath "setup.exe").ToString()
		if (-not (Test-Path $SetupFilePath)) { throw "Error: Setup file not found" }

		Write-LogEntry -Value "Setup file ready at $($SetupFilePath)" -Severity 1

		try {
			# Check setup.exe has valid file signature
			$OfficeC2RVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($SetupFilePath).FileVersion
			Write-LogEntry -Value "Office C2R Setup is running version $OfficeC2RVersion" -Severity 1

			if (Invoke-FileCertVerification -FilePath $SetupFilePath) {
				Write-LogEntry -Value "Setup.exe file has a valid signature." -Severity 1
				# Copying configuration XML file
				Copy-Item -Path "$PSScriptRoot\Configurations\$FileName" -Destination $SetupFolder -Recurse -Force -ErrorAction Stop
				Write-LogEntry -Value "$FileName has been copied to $Setupfolder." -Severity 1
			} else { throw "Error: Unable to verify setup file signature" }

			try {
				# Run M365 setup.exe per configuration file
				Write-LogEntry -Value "Starting $Type $Mode with configuration file $FileName" -Severity 1
				[string]$Arguments = "/configure `"$SetupFolder\$FileName`""
				$Process = Start-Process $SetupFilePath -ArgumentList $Arguments -NoNewWindow -Wait -PassThru -ErrorAction Stop

				if ($mode -eq "Install") {
					# Add Validation File
					if ($Process.ExitCode -eq "0") {
						New-Item -ItemType File -Path $AppValidationFile -Force -Value $AppVersion
						Write-LogEntry -Value "Validation file has been created at $AppValidationFile" -Severity 1
						Write-LogEntry -Value "Install of $AppName is complete" -Severity 1
					} else { Write-LogEntry -Value "Install of $AppName failed with ExitCode: $($Process.ExitCode)" -Severity 3 }
				} elseif ($Mode -eq "Uninstall") {
					if ($Process.ExitCode -eq "0") {
						Remove-Item -Path $AppValidationFile -Force -Confirm:$false
						Write-LogEntry -Value "Validation file has been removed at $AppValidationFile" -Severity 1
						Write-LogEntry -Value "Uninstall of $AppName is complete" -Severity 1
					} else { Write-LogEntry -Value "Install of $AppName failed with ExitCode: $($Process.ExitCode)" -Severity 3 }
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

} catch [System.Exception]{ Write-LogEntry -Value "Error downloading setup.exe from evergreen url. Errormessage: $($_.Exception.Message)" -Severity 3
}

#ENDS
