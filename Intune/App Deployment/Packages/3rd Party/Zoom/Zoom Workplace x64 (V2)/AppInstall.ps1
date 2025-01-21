    <#
    .SYNOPSIS
    Zoom.

    .DESCRIPTION
    Script to install or uninstall Zoom.

    .PARAMETER Mode
    Sets the mode of operation. Supported modes are Install or Uninstall.

    .EXAMPLE 
    powershell.exe -executionpolicy bypass -file .\zoominstall.ps1 -Mode Install
    powershell.exe -executionpolicy bypass -file .\zoominstall.ps1 -Mode Uninstall

    .NOTES
    - AUTHOR : Ashley Forde & Janine Crous
    - Version: 2.0
    - Date   : 12.11.2024

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
    $AppName          = "Zoom Workplace x64"
    $AppVersion       = "6.2.7.49583"
    $Installer        = "ZoomInstallerFull.msi" # assumes the .exe or .msi installer is in the Files folder of the app package.
    $InstallArguments = "/quiet /qn /norestart ZConfig=`"nogoogle = 1;nofacebook = 1;AU2_EnableAutoUpdate = 1`"" # Optional
    $uninstaller      = "CleanZoom.exe"

    # Initialize Directories
    $folderpaths = Initialize-Directories -HomeFolder C:\HUD\

    # Template Variables
    $Date                = Get-Date -Format "MM-dd-yyyy"
    $stagingFolderVar    = $folderPaths.StagingFolder
    $logsFolderVar       = $folderPaths.LogsFolder
    $LogFileName         = "$($AppName)_${Mode}_$Date.log"
    $validationFolderVar = $folderPaths.ValidationFolder
    $AppValidationFile   = "$validationFolderVar\$AppName.txt"

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

                # test if there is a setup file
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
        } catch [System.Exception]{ Write-LogEntry -Value "Error preparing installation $FileName $($mode). Errormessage: $($_.Exception.Message)" -Severity 3 }
    }

    "Uninstall" {
        try {
                # Copy files to staging folder
                Copy-Item -Path "$PSScriptRoot\Files\*" -Destination $SetupFolder -Recurse -Force -ErrorAction Stop
                Write-LogEntry -Value "Setup files have been copied to $Setupfolder." -Severity 1

                # Test if there is a clean zoom exe file
                $CleanZoomExePath = (Join-Path -Path $SetupFolder -ChildPath $uninstaller).ToString()
                if (-not (Test-Path $CleanZoomExePath)) { 
                    throw "Error: Setup file not found" 
                }
                Write-LogEntry -Value "Clean zoom file ready at $($CleanZoomExePath)" -Severity 1
                
                # Remove Zoom files from validation folder
                Get-ChildItem -Path $validationFolderVar -Filter "*Zoom*" | Remove-Item -Force -ErrorAction SilentlyContinue
                Write-LogEntry -Value "Validation file has been removed at $AppValidationFile" -Severity 1

                # Run CleanZoom.exe
                $CleanZoomProcess = Start-Process $CleanZoomExePath -ArgumentList "/SILENT" -Wait -PassThru -ErrorAction Stop
                
                Write-LogEntry -Value "CleanZoom.exe process exit code: $($CleanZoomProcess.ExitCode)" -Severity 1
                
                if ($CleanZoomProcess.ExitCode -eq 0) {
                    Write-LogEntry -Value "CleanZoom.exe has been executed" -Severity 1
                } else {
                    throw "CleanZoom.exe failed with ExitCode: $($CleanZoomProcess.ExitCode)"
                }
                Start-sleep -Seconds 5
                
                # Cleanup 
                if (Test-Path "$SetupFolder") {
                    Remove-Item -Path "$SetupFolder" -Recurse -Force -ErrorAction Continue
                    Write-LogEntry -Value "Cleanup completed successfully" -Severity 1
                }

                } catch [System.Exception] {
                    Write-LogEntry -Value "Error completing uninstall. Errormessage: $($_.Exception.Message)" -Severity 3
                    throw "Uninstallation halted due to an error"
                }

                Write-LogEntry -Value "Uninstall of $AppName is complete" -Severity 1    }

    default {
        Write-Output "Invalid mode: $Mode"
    }
}

