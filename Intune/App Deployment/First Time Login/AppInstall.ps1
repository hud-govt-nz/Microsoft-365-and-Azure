<#
.SYNOPSIS
    This script installs the First Time Login Script (Device), version 1.0, in the "HUD Tools" folder within $env:homedrive\HUD Tools directory.

.DESCRIPTION
    General script for updating user profile settings on devices during first login.

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
$AppName = "First Time Login Script"
$AppVersion = "2.0"

# Initialize Directories
$folderPaths = Initialize-Directories -HomeFolder "C:\HUD\"

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
$SetupFolder = New-Item -ItemType "directory" -Path $stagingFolderVar -Name $AppName -Force | Select-Object -ExpandProperty FullName
Write-LogEntry -Value "Setup folder has been created at: $SetupFolder." -Severity 1

try {
    $Appx = @{
        0 = "XboxIdentityProvider"
        1 = "XboxSpeechToTextOverlay"
        2 = "XboxGamingOverlay"
        3 = "XboxGameOverlay"
        4 = "Xbox.TCUI"
        5 = "MicrosoftTeams"
    }

    foreach ($Key in $Appx.Keys) {
        $AppName = $Appx[$Key]
        $AppPackage = Get-AppxPackage -Name "*$AppName*" -AllUsers
        if ($AppPackage) {
            $AppPackage | Remove-AppPackage -AllUsers -Confirm:$false -ErrorAction Stop -Verbose
            Write-LogEntry -Value "[$(Get-Date)] $AppName has been successfully uninstalled" -Severity 1
        } else {
            Write-LogEntry -Value "[$(Get-Date)] $AppName has not been found" -Severity 1
        }
    }

    # Enabling .NET 3.5
    if ((Get-WindowsOptionalFeature -FeatureName NetFx3 -Online).State -ne "Enabled") {
        Write-LogEntry -Value "[$(Get-Date)] NetFx3 is currently Disabled, enabling Component" -Severity 1
        Enable-WindowsOptionalFeature -FeatureName NetFx3 -Online
    } else {
        Write-LogEntry -Value "[$(Get-Date)] NetFx3 is currently Enabled, Skipping Installation" -Severity 1
    }

    # Add Serial to Settings Menu
    $Path0 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation"
    if ((Get-ItemProperty -Path $Path0).SerialNumberIsValid -ne 1) {
        Write-LogEntry -Value "[$(Get-Date)] SerialNumberIsValid is not present. Adding Key & Value to show Serial Number on Settings page" -Severity 1
        Set-ItemProperty -Path $Path0 -Name SerialNumberIsValid -Value 1 -Force
    }

    # Set preferred language to English (New Zealand)
    Set-WinUserLanguageList -LanguageList en-NZ, mi-latn -Force -Confirm:$false | Out-Null
    Write-LogEntry -Value "[$(Get-Date)] Set preferred languages to English (New Zealand) and Maori." -Severity 1

    # Set country or region to New Zealand
    Set-WinSystemLocale -SystemLocale "en-NZ" | Out-Null
    Write-LogEntry -Value "[$(Get-Date)] Set country or region to New Zealand." -Severity 1

    # Set the home location of the current user
    Set-WinHomeLocation -GeoId 0xb7 | Out-Null
    Write-LogEntry -Value "[$(Get-Date)] Set the home location of the current user." -Severity 1

    # Set regional format and all its settings to be the default NZ settings
    Set-WinUILanguageOverride -Language en-NZ | Out-Null
    Write-LogEntry -Value "[$(Get-Date)] Set regional format and all its settings to be the default NZ settings" -Severity 1

    # Set culture and all its settings to be the default NZ settings
    Set-Culture -CultureInfo en-NZ | Out-Null
    Write-LogEntry -Value "[$(Get-Date)] Set culture and all its settings to be the default NZ settings" -Severity 1

    # Post Install Actions
    # Create validation file
    New-Item -ItemType File -Path $AppValidationFile -Force -Value $AppVersion
    Write-LogEntry -Value "Validation file has been created at $AppValidationFile" -Severity 1
    Write-LogEntry -Value "Install of $AppName is complete" -Severity 1

    # Cleanup
    if (Test-Path "$SetupFolder") {
        Remove-Item -Path "$SetupFolder" -Recurse -Force -ErrorAction Continue
        Write-LogEntry -Value "Cleanup completed successfully" -Severity 1
    }

} catch [System.Exception] {
    Write-LogEntry -Value "Errormessage: $($_.Exception.Message)" -Severity 3
}
