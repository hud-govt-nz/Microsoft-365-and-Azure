<#
.SYNOPSIS
    FortiClient Desktop VPN Client

.DESCRIPTION
    Script to install FortiClient Desktop VPN Client

.PARAMETER Mode
    Sets the mode of operation. Supported modes are Install or Uninstall

.EXAMPLE 
    powershell.exe -executionpolicy bypass -file .\appinstall.ps1 -Mode Install
    powershell.exe -executionpolicy bypass -file .\appinstall.ps1 -Mode Uninstall

.NOTES
    - AUTHOR: Ashley Forde
    - Version: 6.0.5
    - Date: 31.01.25
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
$AppName = "FortiClient"
$Installer = "FortiClientSetup_6.0.5.0209_x64.exe" # assumes the .exe or .msi installer is in the Files folder of the app package.
$InstallArguments = "/quiet /norestart" # Optional

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
            # Copy files to staging folder
            Copy-Item -Path "$PSScriptRoot\Files\*" -Destination $SetupFolder -Recurse -Force -ErrorAction Stop
            Write-LogEntry -Value "Setup files have been copied to $Setupfolder." -Severity 1
                      
            # Set the service to disabled
            $service = Get-Service -Name "FA_Scheduler" -ErrorAction SilentlyContinue
            if ($service) {
                Write-LogEntry -Value "Disabling service FA_Scheduler" -Severity 1
                Set-Service -Name "FA_Scheduler" -StartupType Disabled -ErrorAction SilentlyContinue
            }

            # Command to run a script block upon reboot
            $scriptBlock = {

                # Define paths
                $FCRemove = "C:\HUD\00_Staging\FortiClient\fcremove\fcremove_x64.exe"

                # Run FCRemove.exe if it exists
                if (Test-Path $FCRemove) {
                    $process = Start-Process -FilePath $FCRemove -ArgumentList "-silent -noreboot" -Wait -PassThru

                    if ($process.ExitCode -eq 0) {
                        Write-Output -Value "FortiClient uninstallation completed successfully"
                    } else {
                        Write-Output -Value "FortiClient uninstallation failed with ExitCode: $($process.ExitCode)" -Severity 3
                    }
                }

                # Remove the scheduled task after execution
                $taskName = "Continue FortiClient Uninstall"
                Unregister-ScheduledTask -TaskName $taskName -Confirm:$false

                # Remove setup files
                Remove-Item -Path "C:\HUD\00_Staging\FortiClient" -Recurse -Force -ErrorAction SilentlyContinue

                # Remove all FortiClient shortcuts
                $shortcutPath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\FortiClient"
                Remove-Item -Path $shortcutPath -Recurse -Force -ErrorAction SilentlyContinue

            }

            # Convert script block to a Base64 encoded string
            $encodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes(
                [ScriptBlock]::Create($scriptBlock).ToString()
            ))

            # Create a scheduled task to run the script block once upon reboot
            $settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 30)
            $action = New-ScheduledTaskAction -Execute 'Powershell.exe' -Argument "-ExecutionPolicy Bypass -EncodedCommand $encodedCommand"
            $trigger = New-ScheduledTaskTrigger -AtStartup

            # Register the scheduled task
            $task = Register-ScheduledTask `
                                -Action $action `
                                -Trigger $trigger `
                                -Settings $settings `
                                -TaskName "Continue FortiClient Uninstall" `
                                -Description "Completes the FCRemove.exe task to complete uninstall of FortiClient" `
                                -RunLevel Highest `
                                -User "System" `
                                -TaskPath "\HUD Digital"

            if ($task.State -eq "ready"){
                exit 0

            } else {
                exit 1
            
                } 

            # Restart the computer (uncomment in actual use)
                #perform reboot
                #Write-LogEntry -value $($task.exitcode) -Severity 1


            #Restart-Computer -Confirm:$false

        } catch [System.Exception] {
            Write-LogEntry -Value "Error completing uninstall. Errormessage: $($_.Exception.Message)" -Severity 3
            throw "Uninstallation halted due to an error"
        }
    }

    default {
        Write-Output "Invalid mode: $Mode"
    }
}
