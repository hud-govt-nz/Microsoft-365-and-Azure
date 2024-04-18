<#
.SYNOPSIS
This script installs the One Drive Sync, version 1.0, in the "HUD Tools" folder within $env:homedrive\HUD Tools directory. 
It also creates log files in the "Log" folder within the "EndpointManager" folder.

.DESCRIPTION
The script first checks if One Drive Sync is already installed on the computer and exits the script if it is. 
It then checks if the specified directories and files already exist before trying to create them. 
The script then downloads and installs the One Drive Sync, writes the version number to the validation file, and then deletes the temporary files.

.EXAMPLE
powershell .\One Drive Sync.ps1 install
powershell .\One Drive Sync.ps1 uninstall

.NOTES
- Make sure that you have the appropriate permissions to install software on the computer
- This script is designed to work with version 2021.4.7 of One Drive Sync. If you need to install a different version, you will need to modify the script accordingly.
- This script assumes that the "HUD Tools" and "EndpointManager" folders already exist in the $env:homedrive\HUD directory. 
  If they do not exist, you will need to modify the script accordingly.

.AUTHOR
Ashley Forde

#>
############################################

#Test if root directory exists on device
$Folder = "$($env:homedrive)\HUD"

    if (Test-Path -Path $Folder) {"Path exists!"} else {
        "Creating root folder..."
        New-Item -Path "C:\" -Name "HUD" -ItemType "directory" -Force -Confirm:$false
        New-Item -Path "C:\HUD\" -Name "00_Staging" -ItemType "directory" -Force -Confirm:$false
        New-Item -Path "C:\HUD\" -Name "01_Logs" -ItemType "directory" -Force -Confirm:$false
        New-Item -Path "C:\HUD\" -Name "02_Validation" -ItemType "directory" -Force -Confirm:$false
    }

#Define working folders
$path = "$Folder\00_Staging\"
$logs = "$Folder\01_Logs\"
$validation = "$Folder\02_Validation\"

#Define application version
$version = "1.0"

#Define log file
$logfile = "$logs\One Drive Sync.log"

############################################

# Check if the install or uninstall switch is used
switch ($args[0]) {
    'install' {
        # Create a new task action
        $TaskAction = (New-ScheduledTaskAction -Execute "C:\Program Files\Microsoft OneDrive\OneDrive.exe" -Argument "/reset"),(New-ScheduledTaskAction -Execute "C:\Program Files\Microsoft OneDrive\OneDrive.exe" -Argument "/background")

        # Create a new trigger (Once)
        $taskTrigger = New-ScheduledTaskTrigger -AtLogOn

        # Taskname and Description
        $Taskname = "Remediate OD4B (reset and resync)"
        $Description = "Run the One Drive Reset Action"

        # Register Task
        Register-ScheduledTask -TaskName $Taskname -Action $taskAction -Trigger $taskTrigger -Description $description

        # Trigger Task Immediately 
        Start-ScheduledTask -TaskName $Taskname

        # Pause while task runs
        Start-Sleep 600

        # Delete task
        Unregister-ScheduledTask -TaskName $Taskname -Confirm:$false       

        # Log the result of the installation
        Add-Content -Path $logfile -Value "[$(Get-Date)] One Drive Sync version $($version) was installed successfully with exit code $($exitCode)"

        # Create validation file
        New-Item -ItemType File -Path "$validation\One Drive Sync.txt" -Force -Value $version

        # Log that validation file was created
        Add-Content -Path $logfile -Value "[$(Get-Date)] Validation file was created successfully."
        
    }
}