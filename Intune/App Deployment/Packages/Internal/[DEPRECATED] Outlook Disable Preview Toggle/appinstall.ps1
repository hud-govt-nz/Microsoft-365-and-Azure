<#
.SYNOPSIS
    Disable Outlook Preview Script

.DESCRIPTION
    This script toggles off the Outlook Preview feature in M365 Office.

.PARAMETER Mode
    Sets the mode of operation. Supported modes are Install or Uninstall.

.EXAMPLE 
    powershell.exe -executionpolicy bypass -file .\appinstall.ps1 -Mode Install
    powershell.exe -executionpolicy bypass -file .\appinstall.ps1 -Mode Uninstall

.NOTES
    - AUTHOR: Ashley Forde
    - Version: 1.0
    - Date: 17.4.24
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
$AppName = "Disable Outlook Preview Script"
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
                
        #Auto Complete Registry Key
        $KeyValue = Get-ItemProperty Registry::HKEY_USERS\$SID\Software\Microsoft\Office\16.0\Outlook\Options\General
        $value = $KeyValue.HideNewOutlookToggle
        $path = $KeyValue.pspath

        if ($value -eq 1) {
            Set-ItemProperty -Path $path -Name HideNewOutlookToggle -Value 1 -Force
            Write-LogEntry -Value "Preview toggle has been disabled" -Severity 1        
        } elseif (!$value) {
                New-ItemProperty -Path $path -Name HideNewOutlookToggle -PropertyType DWORD -Value 1 -Force
                Write-LogEntry -Value "Key created and Preview toggle has been disabled" -Severity 1 
                }

         # Check if Outlook is open
         $isOutlookOpen = Get-Process outlook* -ErrorAction SilentlyContinue

         if ($null -eq $isOutlookOpen) {
         # Outlook is not open, run code here
         Write-LogEntry -value "Outlook is not open. Running code..." -Severity 1
         } else {
         # Outlook is open, close all instances
         Write-LogEntry -value "Outlook is open. Closing Outlook..." -Severity 1
 
         # Loop until all Outlook windows are closed
         while ($null -ne $isOutlookOpen) {
         # Close each Outlook window
         Get-Process outlook* | ForEach-Object {
             try {
                 # Attempt to get the active Outlook application object
                 $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
                 $inspector = $outlook.ActiveInspector()
                 if ($inspector -ne $null -and $inspector.IsWordMail() -eq $true) {
                     # Save and close the open message
                     $inspector.CurrentItem().Save()
                     $inspector.Close()
                 }
             } catch {
                 Write-Host "Error: $_"
             }
             
             # Close the Outlook window
             $_.CloseMainWindow() | Out-Null
         }
 
         # Wait for Outlook to close
         Start-Sleep -Seconds 3
 
         # Check if Outlook is still open
         $isOutlookOpen = Get-Process outlook* -ErrorAction SilentlyContinue
 
         # If Outlook is still open, try to force close it
         if ($null -ne $isOutlookOpen) {
             Write-Host "Outlook is still open. Forcing close..."
             $wshell = New-Object -ComObject WScript.Shell
             $wshell.AppActivate("Microsoft Outlook")
             $wshell.SendKeys("%(Y)")
         }
         }
         # Outlook has been closed, run code here
         Write-LogEntry -Value "Outlook has been closed. Running code..." -Severity 1
         }

        Start-Sleep -Seconds 3
        Start-Process outlook

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
            Write-LogEntry -Value "Error updating Outlook. Errormessage: $($_.Exception.Message)" -Severity 3
            }

}

elseif ($Mode -eq "Uninstall") {
    try {
        # Get current user SID and username
        $SID = Get-CurrentUserSID
                
        #Auto Complete Registry Key
        $KeyValue = Get-ItemProperty Registry::HKEY_USERS\$SID\Software\Microsoft\Office\16.0\Outlook\Options\General
        $value = $KeyValue.HideNewOutlookToggle
        $path = $KeyValue.pspath
        
        Remove-ItemProperty -Path $path -Name 'HideNewOutlookToggle' -Force -ErrorAction SilentlyContinue
        Write-LogEntry -Value "HideNewOutlookToggle has been removed" -Severity 1

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