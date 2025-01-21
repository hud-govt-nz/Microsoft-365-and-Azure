<#
.SYNOPSIS
    NZ Regional Settings

.DESCRIPTION
    Script to install NZ Regional Settings

.PARAMETER Mode
    Sets the mode of operation. Supported modes are Install or Uninstall

.EXAMPLE 
    powershell.exe -executionpolicy bypass -file .\appinstall.ps1 -Mode Install
    powershell.exe -executionpolicy bypass -file .\appinstall.ps1 -Mode Uninstall

.NOTES
    - AUTHOR: 
    - Version: 
    - Date: 
#>
# Region Parameters
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("Install","Uninstall")]
    [string]$Mode
)

# Reference functions.ps1
. "$PSScriptRoot\functions.ps1"

# Application Variables
$AppName = "NZ Regional Settings"
$AppVersion = "2.0"
$Installer = "<INSTALLERS>" # assumes the .exe or .msi installer is in the Files folder of the app package.
$InstallArguments = "<INSTALLARGUMENTS>" # Optional
$UninstallArguments = "<UNINSTALLARGUMENTS>" # Optional

# Initialize Directories
$folderpaths = Initialize-Directories -HomeFolder C:\HUD\

# Template Variables
$Date = Get-Date -Format "MM-dd-yyyy"
$stagingFolderVar = $folderPaths.StagingFolder
$logsFolderVar = $folderPaths.LogsFolder
$LogFileName = "$($AppName)_${Mode}_$Date.log"
$validationFolderVar = $folderPaths.ValidationFolder
$AppValidationFile = "$validationFolderVar\$AppName.txt"

# Begin Setup
Write-LogEntry -Value "Initiating setup process" -Severity 1

# Create Setup Folder
$SetupFolder = (New-Item -ItemType "directory" -Path $stagingFolderVar -Name $AppName -Force).FullName
Write-LogEntry -Value "Setup folder has been created at: $Setupfolder." -Severity 1

# Install/Uninstall
switch ($Mode) {
    "Install" {
        # Set preferred language to English (New Zealand)
        Set-WinUserLanguageList -LanguageList en-NZ, mi-latn -Force -Confirm:$false | Out-Null
        Write-LogEntry -Value "Set preferred languages to English (New Zealand)and Maori." -Severity 1
        # Set country or region to New Zealand
        Set-WinSystemLocale -SystemLocale "en-NZ" | Out-Null
        Write-LogEntry -Value "Set country or region to New Zealand." -Severity 1

        # Set the home location of the current user
        Set-WinHomeLocation -GeoId 0xb7 | Out-Null
        Write-LogEntry -Value "Set the home location of the current user." -Severity 1

        # Set regional format and all its settings to be the default NZ settings
        Set-WinUILanguageOverride -Language en-NZ | Out-Null
        Write-LogEntry -Value "Set regional format and all its settings to be the default NZ settings" -Severity 1

        # Set culture and all its settings to be the default NZ settings
        Set-Culture -CultureInfo en-NZ | Out-Null
        Write-LogEntry -Value "Set culture and all its settings to be the default NZ settings" -Severity 1

        # Create validation file
        New-Item -ItemType File -Path $AppValidationFile -Force -Value $AppVersion
        Write-LogEntry -Value "Validation file has been created at $AppValidationFile" -Severity 1
        Write-LogEntry -Value "Install of $AppName is complete" -Severity 1
        }

    "Uninstall" {
        # Remove Validation File
        Remove-Item -Path $AppValidationFile -Force -Confirm:$false
        Write-LogEntry -Value "Validation file has been removed at $AppValidationFile" -Severity 1
        }

    default {
        Write-Output "Invalid mode: $Mode"
    }
}
