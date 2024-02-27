<#
.SYNOPSIS
    Tableau Reader Deployment Script.

.DESCRIPTION

.PARAMETER Mode
    Sets the mode of operation. Supported modes are Install or Uninstall.

.EXAMPLE 
    powershell.exe -executionpolicy bypass -file .\Tableau_Reader.ps1 -Mode Install
    powershell.exe -executionpolicy bypass -file .\Tableau_Reader.ps1 -Mode Uninstall

.NOTES
    - AUTHOR: Ashley Forde
    - Version: 1.0
    - Date: 25 Oct 2023


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
$AppName = "Tableau_Reader"
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

# Download and Provision Installer File
try {
	# Download latest setup.exe
	$SetupFile = "Setup.exe"
	$SetupEvergreenUrl = "https://www.tableau.com/downloads/reader/pc64"
	Write-LogEntry -Value "Attempting to download latest Tableau_Reader setup executable" -Severity 1
	Start-DownloadFile -URL $SetupEvergreenUrl -Path $SetupFolder -Name $SetupFile

	try {
		# Test if file exists
		$SetupFilePath = (Join-Path -Path $SetupFolder -ChildPath "Setup.exe").ToString()
		if (-not (Test-Path $SetupFilePath)) { throw "Error: Setup file not found" }

		Write-LogEntry -Value "Setup file ready at $($SetupFilePath)" -Severity 1

	} catch [System.Exception]{ Write-LogEntry -Value "Error finding Tableau Reader Possible download error. Errormessage: $($_.Exception.Message)" -Severity 3 }

} catch [System.Exception]{ Write-LogEntry -Value "Error downloading Tableau Reader from evergreen url. Errormessage: $($_.Exception.Message)" -Severity 3
}


if ($Mode -eq "Install") {
	try {
		# Run setup.exe per configuration file
		Write-LogEntry -Value "Starting install with file $SetupFile" -Severity 1
		[string]$Arguments = "ACCEPTEULA=1 DESKTOPSHORTCUT=1 /quiet /norestart"
		$Process = Start-Process $SetupFilePath -ArgumentList $Arguments -Wait -PassThru -ErrorAction Stop

		# Add Validation File
		if ($Process.ExitCode -eq "0") {
			New-Item -ItemType File -Path $AppValidationFile -Force -Value $AppVersion
			Write-LogEntry -Value "Validation file has been created at $AppValidationFile" -Severity 1
			Write-LogEntry -Value "Install of $AppName is complete" -Severity 1
		} else {
			Write-LogEntry -Value "Install of $AppName failed with ExitCode: $($Process.ExitCode)" -Severity 3
		}

	} catch {
		Write-LogEntry -Value "Error running installer. Errormessage: $($_.Exception.Message)" -Severity 3
		return # Stop execution of the script after logging a critical error
	}


} elseif ($Mode -eq "Uninstall") {
	try {
		# Run setup.exe per configuration file
		Write-LogEntry -Value "Starting uninstall with configuration file $SetupFile" -Severity 1
		[string]$Arguments = "/uninstall /quiet /norestart"
		$Process = Start-Process $SetupFilePath -ArgumentList $Arguments -Wait -PassThru -ErrorAction Stop

		# Add Validation File
		if ($Process.ExitCode -eq "0") {
			Remove-Item -Path $AppValidationFile -Force -Confirm:$false
			Write-LogEntry -Value "Validation file has been removed at $AppValidationFile" -Severity 1
			Write-LogEntry -Value "Uninstall of $AppName is complete" -Severity 1
		} else {
			Write-LogEntry -Value "Install of $AppName failed with ExitCode: $($Process.ExitCode)" -Severity 3
		}

	} catch {
		Write-LogEntry -Value "Error running installer. Errormessage: $($_.Exception.Message)" -Severity 3
		return # Stop execution of the script after logging a critical error
	}

}

# Cleanup 
if (Test-Path "$SetupFolder") {
	Remove-Item -Path "$SetupFolder" -Recurse -Force -ErrorAction Continue
	Write-LogEntry -Value "Cleanup completed successfully" -Severity 1
}

#ENDS
