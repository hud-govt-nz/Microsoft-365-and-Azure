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
    - Version: 3.1
    - Date: 11.11.24
    - NOTES: 
        - Replaced PowerPoint template at request of Comms


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

# functions
function Remove-HudTemplates {

    [CmdletBinding()]
    param ()

    $registryPaths = @(
        "Registry::HKEY_USERS\$SID\SOFTWARE\Microsoft\Office\16.0\Common\General",
        "Registry::HKEY_USERS\$SID\SOFTWARE\Microsoft\Office\16.0\Word\Options",
        "Registry::HKEY_USERS\$SID\SOFTWARE\Microsoft\Office\16.0\PowerPoint\Options"
        )

    foreach ($path in $registryPaths) {
        Remove-ItemProperty -Path $path -Name "SharedTemplates", "officestartdefaulttab", "HUD Office Templates" -Force -Confirm:$false -ErrorAction SilentlyContinue
        }

    $paths = @(
        "$env:ProgramData\HUD Templates",
        "$AppData\Microsoft\Document Building Blocks\1033\16\*",
        "C:\HUD\02_Validation\HUD - Document Templates.txt"
        )

    Remove-Item -Path $paths -Recurse -Force -Confirm:$false -ErrorAction Continue
}

# Initialize Directories
$folderpaths = Initialize-Directories -HomeFolder C:\HUD\

# Assign the returned values to individual variables
$stagingFolderVar = $folderPaths.StagingFolder
$logsFolderVar = $folderPaths.LogsFolder
$validationFolderVar = $folderPaths.ValidationFolder
$Date = Get-Date -Format "MM-dd-yyyy"
$AppName = "HUD - Office 365 Document Templates"
$AppValidationFile = "$validationFolderVar\$AppName.txt"
$AppVersion = "3.1"
$LogFileName = "$($AppName)_${Mode}_$Date.log"

# Begin Setup
Write-LogEntry -Value "Initiating setup process" -Severity 1

# Get current user information
$SID = Get-CurrentUserSID
$User = (Get-Process -IncludeUserName -Name explorer | Select-Object -First 1 | Select-Object -ExpandProperty UserName).Split("\")[1]
Write-LogEntry -Value "Current user SID is $SID and username is $User" -Severity 1
$AppData = "c:\users\$user\Appdata\Roaming"

# Install/Uninstall
if ($Mode -eq "Install") {
    
    # Remove existing templates
    if (Test-Path "$env:ProgramData\HUD Templates") {
        Remove-HudTemplates
        Write-LogEntry -Value "Removed Template Files, Document Building Blocks files, Registry keys, and App Validation File" -Severity 1
    }

    $TemplatesPath = "$env:ProgramData\HUD Templates"
    $DocumentBuildingBlocksPath = "$AppData\Microsoft\Document Building Blocks\1033\16"

    # Create directories
    New-Item -ItemType Directory -Path $TemplatesPath -Force -Confirm:$false
    New-Item -ItemType Directory -Path $DocumentBuildingBlocksPath -Force -Confirm:$false

    # Copy template files
    Copy-Item -Path "$PSScriptRoot\Templates\*" -Destination $TemplatesPath -Recurse -Force
    Copy-Item -Path "$PSScriptRoot\CoverPages\*" -Destination $DocumentBuildingBlocksPath -Recurse -Force

    # Set registry properties
    $RegistryPath = "Registry::HKEY_USERS\$SID\SOFTWARE\Microsoft\Office\16.0"
    New-ItemProperty -Path "$RegistryPath\Common\General" -Name "SharedTemplates" -PropertyType String -Value $TemplatesPath -Force -Confirm:$false
    New-ItemProperty -Path "$RegistryPath\Word\Options" -Name "officestartdefaulttab" -PropertyType Dword -Value 1 -Force -Confirm:$false
    New-ItemProperty -Path "$RegistryPath\PowerPoint\Options" -Name "officestartdefaulttab" -PropertyType Dword -Value 1 -Force -Confirm:$false
    New-ItemProperty -Path "$RegistryPath\Word\Options" -Name "HUDOfficeTemplates" -PropertyType ExpandString -Value $TemplatesPath -Force -Confirm:$false
    New-ItemProperty -Path "$RegistryPath\PowerPoint\Options" -Name "HUDOfficeTemplates" -PropertyType ExpandString -Value $TemplatesPath -Force -Confirm:$false

	# Add Validation File
	New-Item -ItemType File -Path $AppValidationFile -Force -Value $AppVersion
	Write-LogEntry -Value "Validation file has been created at $AppValidationFile" -Severity 1
	Write-LogEntry -Value "Install of $AppName is complete" -Severity 1

}

elseif ($Mode -eq "Uninstall") {
    Remove-HudTemplates
    Write-LogEntry -Value "Removed Template Files, Document Building Blocks files, Registry keys, and App Validation File" -Severity 1
}
