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
    - Version: 2.0
    - Special Thanks to https://github.com/MSEndpointMgr/M365Apps which provded the starting point for this deployment. 
        - Improvements made:
            - Included uninstall logic for all app variations.
            - 1 script used to manage Office, Visio, Project, Language Pack, Proofing tool install and uninstalls. 
            - Validation/detection script based on script completion where a validation file is generated only if the install completes with exit code 0 (successful)
            - Removed XMLURL as it was not needed in this scenario
	- Date: 11.06.2024
		- refactored

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

# Reference functions.ps1
. "$PSScriptRoot\functions.ps1"

# Application Variables
$AppName = "Microsoft_365_Apps_${Type}"
$AppVersion = "2.0"
$Installer = "setup.exe" # assumes the .exe or .msi installer is in the Files folder of the app package.
$InstallArguments = "<INSTALLARGUMENTS>" # Optional
$UninstallArguments = "<UNINSTALLARGUMENTS>" # Optional
$SetupEvergreenUrl = "https://officecdn.microsoft.com/pr/wsus/setup.exe"

# Initialize Directories
$folderpaths = Initialize-Directories -HomeFolder C:\HUD\

# Template Variables
$Date = Get-Date -Format "MM-dd-yyyy"
$stagingFolderVar = $folderPaths.StagingFolder
$logsFolderVar = $folderPaths.LogsFolder
$LogFileName = "$($AppName)_${Mode}_$Date.log"
$validationFolderVar = $folderPaths.ValidationFolder
$AppValidationFile = "$validationFolderVar\$AppName.txt"

# Assign the returned values to individual variables
$stagingFolderVar = $folderPaths.StagingFolder
$logsFolderVar = $folderPaths.LogsFolder
$validationFolderVar = $folderPaths.ValidationFolder

# Begin Setup
Write-LogEntry -Value "Initiating setup process" -Severity 1

# Create Setup Folder
$SetupFolder = (New-Item -ItemType "directory" -Path $stagingFolderVar -Name $AppName -Force).FullName
Write-LogEntry -Value "Setup folder has been created at: $Setupfolder." -Severity 1

# Select XML Configuration
switch -Wildcard ($Type) {
	"OfficeStd" {$TypeStr = "OfficeStd"}
	"OfficeInclAccess" {$TypeStr = "OfficeInclAccess"}
	"Visio" {$TypeStr = "Visio"}
	"Project" {$TypeStr = "Project"}
	"LangPack" {$TypeStr = "LangPack"}
	"ProofTools" {$TypeStr = "ProofTools"}
}

# Install/Uninstall
if ($Mode -eq "Install") {$FileName = "${TypeStr}_Install.xml"}
    elseif ($Mode -eq "Uninstall") {$FileName = "${TypeStr}_Uninstall.xml"
    }

# Download setup file
try {
    # Set TLS1.3
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls13

    # Download latest office setup.exe
    Write-LogEntry -Value "Attempting to download latest Office setup executable" -Severity 1
    Start-DownloadFile -URL $SetupEvergreenUrl -Path $SetupFolder -Name $Installer

    # Test if there is a setup file
    $SetupFilePath = (Join-Path -Path $SetupFolder -ChildPath $Installer).ToString()
    if (-not (Test-Path $SetupFilePath)) { throw "Error: Setup file not found" }

    Write-LogEntry -Value "Setup file ready at $($SetupFilePath)" -Severity 1

    } catch [System.Exception]{
        Write-LogEntry -Value "Error downloading setup.exe from evergreen url. Errormessage: $($_.Exception.Message)" -Severity 3
        }

# Setup.exe File Verification
try {
    # Check setup.exe has valid file signature
    $OfficeC2RVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($SetupFilePath).FileVersion
    Write-LogEntry -Value "Office C2R Setup is running version $OfficeC2RVersion" -Severity 1

    if (Invoke-FileCertVerification -FilePath $SetupFilePath) {
	    Write-LogEntry -Value "Setup.exe file has a valid signature." -Severity 1 } else { throw "Error: Unable to verify setup file signature" }

    } catch [System.Exception]{
        Write-LogEntry -Value "Error preparing installation $FileName $($mode). Errormessage: $($_.Exception.Message)" -Severity 3
        }

# Copy configuration XML file to Setup Folder
try {
    Copy-Item -Path "$PSScriptRoot\Files\Configurations\$FileName" -Destination $SetupFolder -Recurse -Force -ErrorAction Stop
    Write-LogEntry -Value "$FileName has been copied to $Setupfolder." -Severity 1
   	} catch [System.Exception]{
        Write-LogEntry -Value "Error finding setup.exe Possible download error. Errormessage: $($_.Exception.Message)" -Severity 3
        }
 
# Execute Install or Uninstall of M365 Apps
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