<#
.SYNOPSIS
    This script installs the Outlook Remove Exchange Cache Slider Policy script.

.DESCRIPTION
    The script first checks if the Outlook Remove Exchange Cache Slider Policy script is already installed on the computer and exits the script if it is. 
    It then checks if the specified directories and files already exist before trying to create them. 
    The script then downloads and installs the Outlook Remove Exchange Cache Slider Policy script.

.PARAMETER Mode
    Sets the mode of operation. Supported modes are Install or Uninstall.

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
$AppName = "Outlook Remove Exchange Cache Slider Policy"
$AppVersion = "2.0"

# Initialize Directories
$folderPaths = Initialize-Directories -HomeFolder "C:\HUD\"

# Template Variables
$Date = Get-Date -Format "MM-dd-yyyy"
$stagingFolderVar = $folderPaths.StagingFolder
$logsFolderVar = $folderPaths.LogsFolder
$LogFileName = "$AppName" + "_" + "$Mode" + "_" + "$Date.log"
$validationFolderVar = $folderPaths.ValidationFolder
$AppValidationFile = Join-Path -Path $validationFolderVar -ChildPath "$AppName.txt"

# Begin Setup
Write-LogEntry -Value "Initiating setup process" -Severity 1

# Create Setup Folder
$SetupFolder = New-Item -ItemType "directory" -Path $stagingFolderVar -Name $AppName -Force | Select-Object -ExpandProperty FullName
Write-LogEntry -Value "Setup folder has been created at: $SetupFolder." -Severity 1

# Get current user information
$SID = Get-CurrentUserSID
$User = (Get-Process -IncludeUserName -Name explorer | Select-Object -First 1 | Select-Object -ExpandProperty UserName).Split("\")[1]
Write-LogEntry -Value "Current user SID is $SID and username is $User" -Severity 1

# Install/Uninstall
switch ($Mode) {
    "Install" {
        try {
            # Note: These values are not dynamically removed by Intune when a user is added to this policy, hence this script manually removes them after they are added.   
            $RemoveRegPath = "Registry::HKEY_USERS\$SID\Software\Policies\Microsoft\office\16.0\outlook\cached mode\"
            Remove-ItemProperty -Path $RemoveRegPath -Name syncwindowsettingdays -Force -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
            Remove-ItemProperty -Path $RemoveRegPath -Name syncwindowsetting -Force -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
       
            # Restart Outlook
            Get-Process Outlook* | Stop-Process -Force -Confirm:$false -ErrorAction SilentlyContinue
            Write-LogEntry -Value "Outlook Remove Exchange Cache Slider Policy script version $($AppVersion) was installed successfully." -Severity 1
           
            # Create validation file
            New-Item -ItemType File -Path $AppValidationFile -Force -Value $AppVersion
            Write-LogEntry -Value "Validation file has been created at $AppValidationFile" -Severity 1
            Write-LogEntry -Value "Install of $AppName is complete" -Severity 1

            # Cleanup 
            if (Test-Path -Path $SetupFolder) {
                Remove-Item -Path $SetupFolder -Recurse -Force -ErrorAction Continue
                Write-LogEntry -Value "Cleanup completed successfully" -Severity 1
            }

        } catch {
            Write-LogEntry -Value "Error running installer. ErrorMessage: $($_.Exception.Message)" -Severity 3
            return # Stop execution of the script after logging a critical error
        }
    }

    "Uninstall" {
        try {
            # Add Registry Values back
            $addRegPath = "Registry::HKEY_USERS\$SID\Software\Policies\Microsoft\office\16.0\outlook\cached mode\"
            New-ItemProperty -Path $addRegPath -Name syncwindowsetting -PropertyType DWORD -Value 3 -Force -Confirm:$false | Out-Null

            # Delete validation file
            try {
                Remove-Item -Path $AppValidationFile -Force -Confirm:$false
                Write-LogEntry -Value "Validation file has been removed at $AppValidationFile" -Severity 1

                # Cleanup 
                if (Test-Path -Path $SetupFolder) {
                    Remove-Item -Path $SetupFolder -Recurse -Force -ErrorAction Continue
                    Write-LogEntry -Value "Cleanup completed successfully" -Severity 1
                }
            } catch [System.Exception] {
                Write-LogEntry -Value "Error deleting validation file. ErrorMessage: $($_.Exception.Message)" -Severity 3
            }

        } catch [System.Exception] {
            Write-LogEntry -Value "Error completing uninstall. ErrorMessage: $($_.Exception.Message)" -Severity 3
            throw "Uninstallation halted due to an error"
        }

        Write-LogEntry -Value "Uninstall of $AppName is complete" -Severity 1
    }

    default {
        Write-Output "Invalid mode: $Mode"
    }
}