<#
.SYNOPSIS
    WinSCP

.DESCRIPTION
    Script to install WinSCP

.PARAMETER Mode
    Sets the mode of operation. Supported modes are Install or Uninstall

.EXAMPLE 
    powershell.exe -executionpolicy bypass -file .\appinstall.ps1 -Mode Install
    powershell.exe -executionpolicy bypass -file .\appinstall.ps1 -Mode Uninstall

.NOTES
    - AUTHOR: Ashley Forde
    - Version: 2.0
    - Date: 12.06.2024
#>

#Region Parameters
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
$AppName = "WinSCP"
$AppVersion = "6.3.5"
$Installer = "WinSCP-6.3.5-Setup.exe" # assumes the .exe or .msi installer is in the Files folder of the app package.
$InstallArguments = "/VERYSILENT /ALLUSERS /ALLUSERS /NORESTART" # Optional
$UninstallArguments = "/VERYSILENT /ALLUSERS /ALLUSERS /NORESTART" # Optional

# Define the $HomeFolder variable
$HomeFolder = "C:\HUD" # Requied so that we can create new folder in the root directory later.

# Initialize Directories
$folderpaths = Initialize-Directories -HomeFolder $HomeFolder 

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
        try {
            # Copy files to staging folder
            Copy-Item -Path "$PSScriptRoot\Files\*" -Destination $SetupFolder -Recurse -Force -ErrorAction Stop
            Write-LogEntry -Value "Setup files have been copied to $Setupfolder." -Severity 1

            # create location for configuration files
            $config = New-Item -Path "$HomeFolder" -Name "09_WINSCP" -ItemType "directory" -Force -Confirm:$false
            Copy-Item -Path "$PSScriptRoot\Files\Config\*" -Destination $config -Recurse -Force -ErrorAction Stop
            Write-LogEntry -Value "configuration files have been copied to $Configurations." -Severity 1 # Test if there is a setup file
            $SetupFilePath = (Join-Path -Path $SetupFolder -ChildPath $Installer).ToString()

            if (-not (Test-Path $SetupFilePath)) { 
                throw "Error: Setup file not found" 
            }
            Write-LogEntry -Value "Setup file ready at $($SetupFilePath)" -Severity 1

            try {
                # Run setup with custom arguments and create validation file
                Write-LogEntry -Value "Starting $Mode of $AppName" -Severity 1
                $Process = Start-Process $SetupFilePath -ArgumentList $InstallArguments -Wait -PassThru -ErrorAction Stop

                # Post Install Actions
                if ($Process.ExitCode -eq "0") {
                    # add application configuration .ini to installation directory
					Copy-Item -Path "$config\winSCP_All.ini" -Destination "C:\Program Files (x86)\WinSCP\winSCP.ini" -Force -Confirm:$false # this can change depending on the .ini file used.               
                    # Create validation file
                    New-Item -ItemType File -Path $AppValidationFile -Force -Value $AppVersion
                    Write-LogEntry -Value "Validation file has been created at $AppValidationFile" -Severity 1
                    Write-LogEntry -Value "Install of $AppName is complete" -Severity 1
                } else {
                    Write-LogEntry -Value "Install of $AppName failed with ExitCode: $($Process.ExitCode)" -Severity 3
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
    }

    "Uninstall" {
        try {
            $AppToUninstall = Get-InstalledApps -App $AppName

            # Uninstall App
            $uninstallProcess = Start-Process $AppToUninstall.UninstallString -ArgumentList $UninstallArguments -PassThru -Wait -ErrorAction stop

            # Post Uninstall Actions
            if ($uninstallProcess.ExitCode -eq "0") {
                
                try {
                    # Delete validation file
                    Remove-Item -Path $AppValidationFile -Force -Confirm:$false
                    Write-LogEntry -Value "Validation file has been removed at $AppValidationFile" -Severity 1
                   
                    # Remove leftover data
                    Remove-Item -Path "$HomeFolder\09_WINSCP" -Force -Recurse -Confirm:$false
                    Write-LogEntry -Value "$HomeFolder\09_WINSCP directory has been deleted" -Severity 1
                    
                    # Get the current user SID
                    $SID = Get-currentUserSID
                    Write-LogEntry -Value "Current user SID is $UserSID" -Severity 1

                    # Remove registry keys         
                    $regPath = "Registry::HKEY_USERS\$SID\Software\Martin Prikryl"
                    Remove-Item -Path $regPath -Force -Recurse -Confirm:$false
                    Write-LogEntry -Value "Registry keys updated" -Severity 1

                    # Cleanup 
                    if (Test-Path "$SetupFolder") {
                        Remove-Item -Path "$SetupFolder" -Recurse -Force -ErrorAction Continue
                        Write-LogEntry -Value "Cleanup completed successfully" -Severity 1
                        }

                } catch [System.Exception] {
                    Write-LogEntry -Value "Error deleting validation file. Errormessage: $($_.Exception.Message)" -Severity 3
                    }

                } else {throw "Uninstallation failed with exit code $($uninstallProcess.ExitCode)"
                    }
                    
        } catch [System.Exception] {
            Write-LogEntry -Value "Error completing uninstall. Errormessage: $($_.Exception.Message)" -Severity 3
            throw "Uninstallation halted due to an error"
        }

        Write-LogEntry -Value "Uninstall of $AppName is complete" -Severity 1
    }

    default {
        Write-Output "Invalid mode: $Mode"
    }
}
