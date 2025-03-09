<#
.SYNOPSIS
    Tableau Reader

.DESCRIPTION
    Script to install Tableau Reader

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
$AppName = "Tableau Reader"
$AppVersion = "24.2.931.0"
$Installer = "TableauReader-64bit-2024-2-2.exe" # assumes the .exe or .msi installer is in the Files folder of the app package.
$InstallArguments = "ACCEPTEULA=1 DESKTOPSHORTCUT=1 AUTOUPDATE=1 REMOVEINSTALLEDAPP=1 SILENTLYREGISTERUSER=true /quiet /norestart" # Optional
#$UninstallArguments = "/uninstall /quiet" # Optional

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
        try {
            # Copy files to staging folder
            Copy-Item -Path "$PSScriptRoot\Files\*" -Destination $SetupFolder -Recurse -Force -ErrorAction Stop
            Write-LogEntry -Value "Setup files have been copied to $Setupfolder." -Severity 1

            # Test if there is a setup file
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

            # Auto populate the registration keys
            New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS -Scope Global | Out-Null

            $currentUserSID = Get-CurrentUserSID
            $baseRegistryPath = "HKU:\$currentUserSID\Software\Tableau\Registration"

            # Define the new keys to be created  
            $newKeys = @("Data", "License")  
  
            # Loop through each key and create it  
            foreach ($key in $newKeys) {
                $newKeyPath = Join-Path -Path $baseRegistryPath -ChildPath $key  
                if (-not (Test-Path $newKeyPath)) {
                    New-Item -Path $newKeyPath -Force
                    } else {
                        Write-Output "Registry key already exists: $newKeyPath"  
                        }
                }  

            # Define the path for the License key  
            $licenseKeyPath = Join-Path -Path $baseRegistryPath -ChildPath "License"  
  
            # Define the name and data for the new value under License key  
            $licenseValueName = "a484699a"  
            $licenseValueData = "10"  
  
            # Create the new REG_SZ value under the License key  
            if (Test-Path $licenseKeyPath) {  
                New-ItemProperty -Path $licenseKeyPath -Name $licenseValueName -Value $licenseValueData -PropertyType String -Force
                } else {  
                    Write-Output "License key path does not exist: $licenseKeyPath"  
                    }  
  
            # Define the path for the Data key  
            $dataKeyPath = Join-Path -Path $baseRegistryPath -ChildPath "Data"  
  
            # Define the name and data for the new values under Data key  
            $dataValues = @{  
                "company" = "MHUD"  
                "company_employees" = "500"  
                "country" = "NZ"  
                "email" = "DigitalSupport@hud.govt.nz"  
                "first_name" = "Digital"  
                "last_name" = "TableauUser"  
                "opt_in" = "false"  
                "state" = "WGN"  
                "registration_date" = "09/24/2024 12:09:59.153 AM"  
            }  
  
            # Create the new REG_SZ values under the Data key  
            if (Test-Path $dataKeyPath) {  
                foreach ($name in $dataValues.Keys) {  
                    New-ItemProperty -Path $dataKeyPath -Name $name -Value $dataValues[$name] -PropertyType String -Force
                    }  
                } else {  
                    Write-Output "Data key path does not exist: $dataKeyPath"  
                    }  
            
            Remove-PSDrive HKU

        } catch [System.Exception]{ Write-LogEntry -Value "Error preparing installation $FileName $($mode). Errormessage: $($_.Exception.Message)" -Severity 3 }
    }

    "Uninstall" {
        try {
            $AppToUninstall = Get-InstalledApps -App $AppName

            # Uninstall App
            $uninstallcommand =  "cmd.exe"
            $uninstallProcess = Start-Process $uninstallcommand -ArgumentList "/c $($AppToUninstall.quietUninstallString)" -NoNewWindow -PassThru -Wait -ErrorAction stop

            # Post Uninstall Actions
            if ($uninstallProcess.ExitCode -eq "0") {
                # Delete validation file
                try {
                    Remove-Item -Path $AppValidationFile -Force -Confirm:$false
                    Write-LogEntry -Value "Validation file has been removed at $AppValidationFile" -Severity 1

                    # Cleanup 
                    if (Test-Path "$SetupFolder") {
                        Remove-Item -Path "$SetupFolder" -Recurse -Force -ErrorAction Continue
                        Write-LogEntry -Value "Cleanup completed successfully" -Severity 1
                    }
                } catch [System.Exception] {
                    Write-LogEntry -Value "Error deleting validation file. Errormessage: $($_.Exception.Message)" -Severity 3
                }
            } else {
                throw "Uninstallation failed with exit code $($uninstallProcess.ExitCode)"
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
