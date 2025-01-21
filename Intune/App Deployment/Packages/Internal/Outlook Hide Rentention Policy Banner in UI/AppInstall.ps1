<#
.SYNOPSIS
    Suppress Retention Policy in Outlook (Classic) UI

.DESCRIPTION
    Script is for suppressing then relevant retention policy from appearing in the Outlook (Classic) UI

.PARAMETER Mode
    Sets the mode of operation. Supported modes are Install or Uninstall.

.EXAMPLE 
    powershell.exe -executionpolicy bypass -file .\appinstall.ps1 -Mode Install
    powershell.exe -executionpolicy bypass -file .\appinstall.ps1 -Mode Uninstall

.NOTES
    - AUTHOR: Ashley Forde
    - Version: 1.0
    - Date: 15.01.25
#>

# Parameters
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
$AppName = "Suppress Retention Policy in Outlook (Classic) UI"
$AppValidationFile = "$validationFolderVar\$AppName.txt"
$AppVersion = "1.0"
$LogFileName = "$($AppName)_${Mode}_$Date.log"


# Begin Setup
Write-LogEntry -Value "Initiating setup process" -Severity 1
$SetupFolder = (New-Item -ItemType "directory" -Path $stagingFolderVar -Name $AppName -Force).FullName
Write-LogEntry -Value "Setup folder has been created at: $Setupfolder." -Severity 1

# Install/Uninstall
if ($Mode -eq "Install") {
    try {
        # Get current user SID and username
        $SID = Get-CurrentUserSID
        $User = (Get-Process -IncludeUserName -Name explorer | Select-Object -First 1 | Select-Object -ExpandProperty UserName).Split("\")[1]
        Write-LogEntry -Value "Current user SID is $SID and username is $User" -Severity 1
        
        $SuppressRetentionUI = Get-ItemProperty Registry::HKEY_USERS\$SID\Software\Microsoft\Office\16.0\Outlook\Options
        $value = $SuppressRetentionUI.SuppressRetentionPolicyUI
        $path = $SuppressRetentionUI.pspath

        # Set value in Registry
        if ($value -eq 0) {
            Set-ItemProperty -Path $path -Name SuppressRetentionPolicyUI -Value 1 -Force
            Write-LogEntry -Value "Retention Policy Note in Email banner has been disabled" -Severity 1
        } elseif (!$value) {
            New-ItemProperty -Path $path -Name SuppressRetentionPolicyUI -PropertyType DWORD -Value 1 -Force            
            Write-LogEntry -Value "Retention Policy Note in Email banner has been created" -Severity 1
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

        } catch [System.Exception]{
            Write-LogEntry -Value "Error running installer. Errormessage: $($_.Exception.Message)" -Severity 3
            }

}

elseif ($Mode -eq "Uninstall") {
    try {
        # Get current user SID and username
        $SID = Get-CurrentUserSID
        $User = (Get-Process -IncludeUserName -Name explorer | Select-Object -First 1 | Select-Object -ExpandProperty UserName).Split("\")[1]
        Write-LogEntry -Value "Current user SID is $SID and username is $User" -Severity 1
        
        $SuppressRetentionUI = Get-ItemProperty Registry::HKEY_USERS\$SID\Software\Microsoft\Office\16.0\Outlook\Options
        $value = $SuppressRetentionUI.SuppressRetentionPolicyUI
        $path = $SuppressRetentionUI.pspath
        
        Remove-ItemProperty -Path $path -Name 'SuppressRetentionPolicyUI' -Force -ErrorAction SilentlyContinue
        Write-LogEntry -Value "Retention Policy Note in Email banner has been enabled" -Severity 1

        # Add Validation File
        Remove-Item -Path $AppValidationFile -Force -Confirm:$false
        Write-LogEntry -Value "Validation file has been removed at $AppValidationFile" -Severity 1
        Write-LogEntry -Value "Uninstall of $AppName is complete" -Severity 1

        # Cleanup 
        if (Test-Path "$SetupFolder") {
            Remove-Item -Path "$SetupFolder" -Recurse -Force -ErrorAction Continue
            Write-LogEntry -Value "Cleanup completed successfully" -Severity 1
            }

        } catch [System.Exception]{
            Write-LogEntry -Value "Error Removing registry key. Errormessage: $($_.Exception.Message)" -Severity 3
            }
}