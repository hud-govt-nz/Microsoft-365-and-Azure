<#
.SYNOPSIS
    This script installs the 22H2 Search Function Fix

.DESCRIPTION
    Script to install 22H2 Search Function Fix

.NOTES
    - AUTHOR: 
    - Version: 
    - Date: 
#>

# Reference functions.ps1
. "$PSScriptRoot\functions.ps1"

# Application Variables
$AppName = "22H2 Search Function Fix"
$AppVersion = "2.0"

# Initialize Directories
$folderPaths = Initialize-Directories -HomeFolder "C:\HUD\"

# Template Variables
$Date = Get-Date -Format "MM-dd-yyyy"
$stagingFolderVar = $folderPaths.StagingFolder
$logsFolderVar = $folderPaths.LogsFolder
$LogFileName = "$AppName`_Install_$Date.log"
$validationFolderVar = $folderPaths.ValidationFolder
$AppValidationFile = "$validationFolderVar\$AppName.txt"

# Begin Setup
Write-LogEntry -Value "Initiating setup process" -Severity 1

# Create Setup Folder
$SetupFolder = New-Item -ItemType "directory" -Path $stagingFolderVar -Name $AppName -Force
Write-LogEntry -Value "Setup folder has been created at: $SetupFolder." -Severity 1

# Install
try {
    $validKeys = @("Loc_0409", "Loc_0481", "Loc_0804")
    $Keys = Get-ChildItem -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Input\Locales"
    $Keys | ForEach-Object {
        if ($validKeys -contains $_.Name.Split("\")[-1]) {                   
            # Log the result of the installation
            Write-LogEntry -Value "$($_.Name) exists on device" -Severity 1
        } else {
            Remove-Item -Path "Registry::$($_.Name)" -Recurse -Force -Confirm:$false

            # Log the result of the installation
            Write-LogEntry -Value "$($_.Name) has been removed from device" -Severity 1
        }
    }

    New-Item -ItemType File -Path $AppValidationFile -Force -Value $AppVersion
    Write-LogEntry -Value "Validation file has been created at $AppValidationFile" -Severity 1
    Write-LogEntry -Value "Install of $AppName is complete" -Severity 1

    # Cleanup 
    if (Test-Path $SetupFolder) {
        Remove-Item -Path $SetupFolder -Recurse -Force -ErrorAction Continue
        Write-LogEntry -Value "Cleanup completed successfully" -Severity 1
    }
} catch {
    Write-LogEntry -Value "Error running installer. ErrorMessage: $($_.Exception.Message)" -Severity 3
    return # Stop execution of the script after logging a critical error
}
