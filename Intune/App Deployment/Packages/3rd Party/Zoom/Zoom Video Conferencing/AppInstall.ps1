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
    - AUTHOR: Ashley Forde & Janine Crous
    - Version: 2.0
    - Date: 12.11.2024

.UPDATES
    - 2.0 - 12.11.2024 - Include Clean Zoom exe unsinstaller tool 
#>

#Region Parameters
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("Install", "Uninstall", IgnoreCase = $true)]
    [string]$Mode
)

# Reference functions.ps1 (Assuming it contains necessary functions like Initialize-Directories and Write-LogEntry)
. "$PSScriptRoot\functions.ps1"

# Download function
function Start-DownloadFile {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$URL,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )
    begin {
        $WebClient = New-Object -TypeName System.Net.WebClient
    }
    process {
        if (-not (Test-Path -Path $Path)) {
            New-Item -Path $Path -ItemType Directory -Force | Out-Null
        }
        $WebClient.DownloadFile($URL, (Join-Path -Path $Path -ChildPath $Name))
    }
    end {
        $WebClient.Dispose()
    }
}

# Initialize Directories
$folderPaths = Initialize-Directories -HomeFolder C:\HUD\
$stagingFolderVar = $folderPaths.StagingFolder
$logsFolderVar = $folderPaths.LogsFolder
$validationFolderVar = $folderPaths.ValidationFolder
$Date = Get-Date -Format "MM-dd-yyyy"
$AppName = "Zoom"
$AppValidationFile = "$validationFolderVar\$AppName.txt"
$AppVersion = "6.2.7.49583"
$LogFileName = "$($AppName)_${Mode}_$Date.log"

# Begin Setup
Write-LogEntry -Value "Initiating setup process" -Severity 1
$SetupFolder = (New-Item -ItemType "directory" -Path $stagingFolderVar -Name $AppName -Force).FullName
Write-LogEntry -Value "Setup folder has been created at: $SetupFolder." -Severity 1

# Install/Uninstall
if ($Mode -eq "Install") {
    try {
        # Clear legacy Validation file
        Remove-Item -Path "C:\HUD\02_Validation\Zoom(64bit).txt" -Force -Confirm:$false -ErrorAction SilentlyContinue
        Write-LogEntry "Legacy validation file has been removed" -Severity 1

        # Download and run CleanZoom.exe
        $CleanZoomZip = "CleanZoom.zip"
        Start-DownloadFile -URL "https://assets.zoom.us/docs/msi-templates/CleanZoom.zip?_ga=2.190707371.1896308117.1731364113-2080433774.1731364113" -Path $SetupFolder -Name $CleanZoomZip -ErrorAction Stop

        start-sleep 3

        # Unzip the file
        Expand-Archive -Path "$SetupFolder\$CleanZoomZip" -Destination $SetupFolder -Force

        #run CleanZoom.exe
        $CleanZoomExe = "$SetupFolder\CleanZoom.exe"
        Start-Process -FilePath $CleanZoomExe -Wait -ErrorAction Stop

        Write-LogEntry -Value "CleanZoom.exe has been executed" -Severity 1

    }
    catch {
        Write-LogEntry -Value "Error downloading CleanZoom.exe. Errormessage: $($_.Exception.Message)" -Severity 3
        return # Stop execution of the script after logging a critical error
    }


    try {
        $installerFileName = "ZoomInstallerFull.msi"
        Start-DownloadFile -URL "https://cdn.zoom.us/prod/6.2.7.49583/x64/ZoomInstallerFull.msi" -Path $SetupFolder -Name $installerFileName -ErrorAction Stop
        Write-LogEntry -Value "Downloaded Zoom installer to $installerFileName." -Severity 1

        $SetupFilePath = "$SetupFolder\$installerFileName"
        if (-not (Test-Path $SetupFilePath)) {
            throw "Installer file not found."
        }
        Write-LogEntry -Value "Found installer file at $SetupFilePath." -Severity 1

        $Arguments = "/quiet /qn /norestart ZConfig=`"nogoogle=1;nofacebook=1;AU2_EnableAutoUpdate=1`""
        $Process = Start-Process -FilePath $SetupFilePath -ArgumentList $Arguments -Wait -PassThru -ErrorAction Stop

        if ($Process.ExitCode -eq 0) {
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

    } catch [System.Exception] {
        Write-LogEntry -Value "Error running installer. Errormessage: $($_.Exception.Message)" -Severity 3
        return # Stop execution of the script after logging a critical error
    }

} elseif ($Mode -eq "Uninstall") {
    try {
        # Find Zoom Uninstaller
        $MyApp = Get-InstalledApps -App "workplace*"

        # Uninstall App
        $uninstall_command = 'MsiExec.exe'
        $Result = (($MyApp.UninstallString -split ' ')[1] -replace '/I','/X ') + ' /quiet'
        $uninstall_args = [string]$Result
        $uninstallProcess = Start-Process $uninstall_command -ArgumentList $uninstall_args -PassThru -Wait -ErrorAction Stop

        # Post Uninstall Actions
        if ($uninstallProcess.ExitCode -eq 0) {
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
