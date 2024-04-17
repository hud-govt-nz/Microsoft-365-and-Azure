<#
.APP: Disable Outlook Preview Script
.DESCRIPTION: This script toggles off the Outlook Preview feature in M365 Office. 
.AUTHOR: Ashley Forde
.DATE: 14 June 2023
#>

# Root Folder
$Directory = 'HUD'

# Define functions
function Write-Log {
  param(
    [string]$Path,
    [string]$Value
    )
  Add-Content -Path $Path -Value $Value
  }
function Get-CurrentUserSID {
  # Set current user
  $currentUser = (Get-Process -IncludeUserName -Name explorer | Select-Object -First 1 | Select-Object -ExpandProperty UserName).Split("\")[1]
  $keys = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" -Recurse

  foreach ($key in $keys) {
    if (($key.GetValueNames() | ForEach-Object { $key.GetValue($_) }) -match $CurrentUser) {
      $sid = $key.PSChildName
      break
      }
    }
  # SID for current user
  return $sid
  }

# Create Directories
$HomeFolder = "$($env:homedrive)\$Directory"
  if (Test-Path -Path $HomeFolder) { 
    "Path exists!"
    } else { 
        "Creating root folder..."
        New-Item -Path $HomeFolder -ItemType "directory" -Force -Confirm:$false | Out-Null
            
        foreach($subFolder in "00_Staging", "01_Logs", "02_Validation") {
          New-Item -Path "$HomeFolder\" -Name $subFolder -ItemType "directory" -Force -Confirm:$false | Out-Null
          }
      }

#Set Variables
$path = "$HomeFolder\00_Staging"
$logs = "$HomeFolder\01_Logs"
$validation = "$HomeFolder\02_Validation"
$AppName=[string]'Disable Outlook Preview Script'
$AppVersion="1.0"
$AppValidationFile="$validation\$AppName.txt"
$AppLog="$logs\$AppName.log"
$UserSID = Get-CurrentUserSID

# Check if the install or uninstall switch is used
switch ($args[0]) {
  'install' {
    try {
      #Add key to CURRENT USER (HKEY_USERS\$USERSID)
      New-ItemProperty -Path "Registry::HKEY_USERS\$UserSID\Software\Microsoft\Office\16.0\Outlook\Options\General" -Name "HideNewOutlookToggle" -PropertyType Dword -Value 1 -Force -Confirm:$false    
      Write-Log -Path $AppLog -Value "[$(Get-Date)] Outlook Preview Toggle has been disabled"
      } catch {
          Write-Log -Path $AppLog -Value "[$(Get-Date)] Error adding registry key: $_"
          exit 1
          }
    try {
      # Create validation file
      New-Item -ItemType File -Path $AppValidationFile -Force -Value $AppVersion | Out-Null
      Write-Log -Path $AppLog -Value "[$(Get-Date)] Validation file was created successfully."
      } catch {
          Write-Log -Path $AppLog -Value "[$(Get-Date)] Error creating validation file: $_"
          exit 1
          }
    }
  'uninstall' {
    try {
      Remove-ItemProperty -Path "Registry::HKEY_USERS\$UserSID\Software\Microsoft\Office\16.0\Outlook\Options\General" -Name "HideNewOutlookToggle" -Force -Confirm:$false
      } catch {
          Write-Log -Path $AppLog -Value "[$(Get-Date)] Error removing registry key: $_"
          exit 1
          }
    try {
      # Delete validation file
      Remove-Item -Path $AppValidationFile -Force -ErrorAction Stop
      Write-Log -Path $AppLog -Value "[$(Get-Date)] Validation file was deleted successfully."
      } catch {
          Write-Log -Path $AppLog -Value "[$(Get-Date)] Error deleting validation file: $_"
          exit 1
          }
            
        }

  default {
    Write-Log -Path $AppLog -Value "[$(Get-Date)] Invalid argument. Please specify 'install' or 'uninstall'."
    exit 1
    }
}