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
$AppName = "HUD - Font Install"
$AppValidationFile = "$validationFolderVar\$AppName.txt"
$AppVersion = "3.0"
$LogFileName = "$($AppName)_${Mode}_$Date.log"

# Begin Setup
Write-LogEntry -Value "Initiating setup process" -Severity 1
$SetupFolder = (New-Item -ItemType "directory" -Path $stagingFolderVar -Name $AppName -Force).FullName
Write-LogEntry -Value "Setup folder has been created at: $Setupfolder." -Severity 1

# Install/Uninstall
if ($Mode -eq "Install") {
    
    # Copy files
    Copy-Item -Path "$PSScriptRoot\Files\*" -Destination $stagingFolderVar -Recurse -Force

    # Add font file names to an array
    $fonts =@()
    $fonts += Get-ChildItem -Path $stagingFolderVar

    # Load each file and assign registry string value
    Foreach ($font in $fonts) {
        copy-Item -Path "$path\$($font)" -Destination "$env:windir\Fonts" -Force -Confirm:$false
        New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts" -Name $font.Name -PropertyType String -Value $font.Name -Force -Confirm:$false
        }

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

elseif ($Mode -eq "Uninstall") {
    # Complile National font list array
    $installedfonts=@()
    $installedfonts += Get-ChildItem -Path C:\Windows\Fonts | Where-Object {$_.Name -ilike "*National*"}
    
    if ($null -ne $installedfonts) {
        try {
            #Uninstall HUD Font Install silently
            Foreach ($font in $installedfonts) {
                Remove-Item -Path "C:\Windows\Fonts\$font" -Force -Confirm:$false 
                Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts" -Name $font.name -Force -Confirm:$false
                # Log the result of the uninstall
                Write-LogEntry -Value "HUD Font $($font.name) version $($version) was uninstalled successfully." -Severity 1
            }
            } catch {
                Write-LogEntry -Value "Error uninstalling fonts: $_" -severity 3
                exit 1
                }
            
        # Add Validation File
        Remove-Item -Path $AppValidationFile -Force -Confirm:$false
        Write-LogEntry -Value "Validation file has been removed at $AppValidationFile" -Severity 1
        Write-LogEntry -Value "Uninstall of $AppName is complete" -Severity 1
        }
}
