<#
.SYNOPSIS
    HUD Template Documents.

.DESCRIPTION
    Deploys a set of template files to M365 Apps for use by HUD Staff


.PARAMETER Mode
	Sets the mode of operation. Supported modes are Install or Uninstall.

.EXAMPLE 
    powershell.exe -executionpolicy bypass -file .\appinstall.ps1 -Mode Install
    powershell.exe -executionpolicy bypass -file .\appinstall.ps1 -Mode Uninstall

.NOTES
    - AUTHOR: Ashley Forde
    - Version: 3.0
    - Date: 18.04.2024
    - NOTES: 
        - Refactored script to use functions.ps1

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

# Initialize Directories
$folderpaths = Initialize-Directories -HomeFolder C:\HUD\

# Assign the returned values to individual variables
$stagingFolderVar = $folderPaths.StagingFolder
$logsFolderVar = $folderPaths.LogsFolder
$validationFolderVar = $folderPaths.ValidationFolder
$Date = Get-Date -Format "MM-dd-yyyy"
$AppName = "HUD - Desktop Shortcuts"
$AppValidationFile = "$validationFolderVar\$AppName.txt"
$AppVersion = "3.0"
$LogFileName = "$($AppName)_${Mode}_$Date.log"

# Begin Setup
Write-LogEntry -Value "Initiating setup process" -Severity 1
$SetupFolder = (New-Item -ItemType "directory" -Path $stagingFolderVar -Name $AppName -Force).FullName
Write-LogEntry -Value "Setup folder has been created at: $Setupfolder." -Severity 1

# Get current user information
$SID = Get-CurrentUserSID
$User = (Get-Process -IncludeUserName -Name explorer | Select-Object -First 1 | Select-Object -ExpandProperty UserName).Split("\")[1]
Write-LogEntry -Value "Current user SID is $SID and username is $User" -Severity 1

# Install/Uninstall
if ($Mode -eq "Install") {
    
    #  Copy HUD Custom Shortcuts Icon files to local device
    $IcoFiles = New-Item -Path $folderpaths.HomeFolder -Name 03_Icons -ItemType Directory -Force -Confirm:$false
    $LinkFiles = New-Item -Path $folderpaths.HomeFolder -Name 04_Links -ItemType Directory -Force -Confirm:$false
    Copy-Item -Path ".\Shortcuts\*" -Destination $IcoFiles -Recurse -Force
    Write-LogEntry -Value "Shortcut icons coped to 03_Icons folder" -Severity 1

    # Create an array of shortcut names and URLs
    $shortcuts = @(
        @{name="Aho"; url="https://fa-evjy-saasfaprod1.fa.ocs.oraclecloud.com/fscmUI/faces/FuseWelcome"; icon="$IcoFiles\Aho.ico"},
        @{name="HUD Support Hub"; url="https://mhud.sharepoint.com/sites/im"; icon="$IcoFiles\HUD.ico"}
        )

    #Create Shortcuts
    foreach($shortcut in $shortcuts) {
        $link = "$($LinkFiles)\" + $shortcut.name + ".lnk"
        [System.IO.Path]::GetFileName($link)
        
        # Create shortcut that can be pinned to taskbar
        $wshShell = New-Object -ComObject WScript.Shell
        $objShortcut = $wshShell.CreateShortcut($link)
        $objShortcut.TargetPath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
        $objShortcut.Arguments = $shortcut.url
        $objShortcut.IconLocation = "$($shortcut.icon)"
        $objShortcut.Save()

        Copy-Item -Path $link -Destination "C:\Users\Public\Desktop" -Force -Confirm:$false -ErrorAction SilentlyContinue
        Write-LogEntry -Value "$($shortcut.name) shortcut successfully created" -Severity 1

	# Add Validation File
	New-Item -ItemType File -Path $AppValidationFile -Force -Value $AppVersion
	Write-LogEntry -Value "Validation file has been created at $AppValidationFile" -Severity 1
	Write-LogEntry -Value "Install of $AppName is complete" -Severity 1
    
    # Cleanup 
    if (Test-Path "$SetupFolder") {
        Remove-Item -Path "$SetupFolder" -Recurse -Force -ErrorAction Continue
        Write-LogEntry -Value "Cleanup completed successfully" -Severity 1
        }
    }
}
elseif ($Mode -eq "Uninstall") {
    # Delete Shortcuts and Icons
    Remove-Item -Path "C:\Users\Public\Desktop\Aho.lnk" -Force -Confirm:$false -ErrorAction SilentlyContinue
    Remove-Item -Path "C:\Users\Public\Desktop\HUD Support Hub.lnk" -Force -Confirm:$false -ErrorAction SilentlyContinue
    Remove-Item -Path "C:\HUD\03_Icons" -Recurse -Force -Confirm:$false -ErrorAction SilentlyContinue
    Remove-Item -Path "C:\HUD\04_Links" -Recurse -Force -Confirm:$false -ErrorAction SilentlyContinue
    Write-LogEntry -Value "Uninstall of $AppName is complete" -Severity 1

    # Add Validation File

    # Remove Validation Files
    Get-ChildItem -Path "C:\HUD\02_Validation\HUD - Custom Shortcuts.txt","C:\HUD\02_Validation\HUDCustomShortcuts.txt" -ErrorAction SilentlyContinue | Remove-Item -Force -Confirm:$false
	Remove-Item -Path $AppValidationFile -Force -Confirm:$false
	Write-LogEntry -Value "Validation file has been removed at $AppValidationFile" -Severity 1
	
	# Cleanup 
	if (Test-Path "$SetupFolder") {
		Remove-Item -Path "$SetupFolder" -Recurse -Force -ErrorAction Continue
		Write-LogEntry -Value "Cleanup completed successfully" -Severity 1
    	}
}
