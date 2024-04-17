<#
.SYNOPSIS
    HUD Template Documents.

.DESCRIPTION
    Deploys a set of template files to M365 Apps for use by HUD Staff

.PARAMETER Mode
    Sets the mode of operation. Supported modes are Install or Uninstall.

.EXAMPLE 
    powershell.exe -executionpolicy bypass -file .\AppInstall.ps1 -Mode Install
    powershell.exe -executionpolicy bypass -file .\AppInstall.ps1 -Mode Uninstall

.NOTES
    - AUTHOR: Ashley Forde
    - Version: 2.8
#>

# Parameters
[CmdletBinding()]
param(
	[Parameter(Mandatory = $false)]
	[ValidateNotNullOrEmpty()]
	[ValidateSet("Install","Uninstall")]
	[string]$Mode
)

# Functions
function Write-LogEntry {
	param(
		[Parameter(Mandatory = $true,HelpMessage = "Value added to the log file.")]
		[ValidateNotNullOrEmpty()]
		[string]$Value,
		[Parameter(Mandatory = $true,HelpMessage = "Severity for the log entry. 1 for Informational, 2 for Warning and 3 for Error.")]
		[ValidateNotNullOrEmpty()]
		[ValidateSet("1","2","3")]
		[string]$Severity,
		[Parameter(Mandatory = $false,HelpMessage = "Name of the log file that the entry will written to.")]
		[ValidateNotNullOrEmpty()]
		[string]$FileName = $LogFileName
	)
	# Determine log file location
	$LogFilePath = Join-Path -Path $logsFolderVar -ChildPath $FileName

	# Construct time stamp for log entry
	$Time = -join @((Get-Date -Format "HH:mm:ss.fff")," ",(Get-WmiObject -Class Win32_TimeZone | Select-Object -ExpandProperty Bias))

	# Construct date for log entry
	$Date = (Get-Date -Format "MM-dd-yyyy")

	# Construct context for log entry
	$Context = $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)

	# Construct final log entry
	$LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""$($LogFileName)"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"

	# Add value to log file
	try {
		Out-File -InputObject $LogText -Append -NoClobber -Encoding Default -FilePath $LogFilePath -ErrorAction Stop
		if ($Severity -eq 1) {
			Write-Verbose -Message $Value
		}
		elseif ($Severity -eq 3) {
			Write-Warning -Message $Value
		}
	}
	catch [System.Exception]{
		Write-Warning -Message "Unable to append log entry to $LogFileName.log file. Error message at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
	}
}
function Initialize-Directories {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$HomeFolder
	)

	# Check if the path exists
	if (Test-Path -Path $HomeFolder) {
		Write-Verbose "Home folder exists..."
		# Force creating 00_Staging folder at a minimum if it is missing
		New-Item -Path "$HomeFolder" -Name "00_Staging" -ItemType "directory" -Force -Confirm:$false | Out-Null

	}
	else {
		Write-Verbose "Creating root folder..."
		New-Item -Path $HomeFolder -ItemType "directory" -Force -Confirm:$false
		if (-not $?) {
			Write-Verbose "Failed to create $HomeFolder"
		}

		# Create subfolders
		foreach ($subFolder in "00_Staging","01_Logs","02_Validation") {
			New-Item -Path "$HomeFolder\" -Name $subFolder -ItemType "directory" -Force -Confirm:$false
			if (-not $?) {
				Write-Verbose -Message "Failed to create sub-folder $subFolder under $HomeFolder"
			}
		}
	}

	# Calculate subfolder paths
	$StagingFolder = Join-Path -Path $HomeFolder -ChildPath "00_Staging"
	$LogsFolder = Join-Path -Path $HomeFolder -ChildPath "01_Logs"
	$ValidationFolder = Join-Path -Path $HomeFolder -ChildPath "02_Validation"


	# Return the folder paths as a custom object
	return @{
		HomeFolder = $HomeFolder
		StagingFolder = $StagingFolder
		LogsFolder = $LogsFolder
		ValidationFolder = $ValidationFolder
	}
}
function Get-CurrentUserSID {
    $currentUser = (Get-Process -IncludeUserName -Name explorer | Select-Object -First 1 | Select-Object -ExpandProperty UserName).Split("\")[1]
    $profileListKeys = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" -Recurse

    foreach ($profileListKey in $profileListKeys) {
        if (($profileListKey.GetValueNames() | ForEach-Object { $profileListKey.GetValue($_) }) -match $currentUser) {
            $sid = $profileListKey.PSChildName
            break
        }
    }

    return $sid
}
function Remove-HudTemplates {

    [CmdletBinding()]
    param ()

    $registryPaths = @(
        "Registry::HKEY_USERS\$UserSID\SOFTWARE\Microsoft\Office\16.0\Common\General",
        "Registry::HKEY_USERS\$UserSID\SOFTWARE\Microsoft\Office\16.0\Word\Options",
        "Registry::HKEY_USERS\$UserSID\SOFTWARE\Microsoft\Office\16.0\PowerPoint\Options"
    )

    foreach ($path in $registryPaths) {
        Remove-ItemProperty -Path $path -Name "SharedTemplates", "officestartdefaulttab", "HUD Office Templates" -Force -Confirm:$false -ErrorAction SilentlyContinue
    }

    $paths = @(
        "$env:ProgramData\HUD Templates",
        "C:\Users\$CurrentUserName\AppData\Roaming\Microsoft\Document Building Blocks\1033\16\*",
        $AppValidationFile
    )
    Remove-Item -Path $paths -Recurse -Force -Confirm:$false -ErrorAction Stop
}

# Initialisations
$HomeFolder = "C:\HUD"
$folderPaths = Initialize-Directories -HomeFolder $HomeFolder
$stagingFolderVar = $folderPaths.StagingFolder
$logsFolderVar = $folderPaths.LogsFolder
$validationFolderVar = $folderPaths.ValidationFolder
$Date = Get-Date -Format "MM-dd-yyyy"
$AppName = "HUD - Document Templates"
$AppValidationFile = "$validationFolderVar\$AppName.txt"
$AppVersion = "2.9"
$LogFileName = "$($AppName)_${Mode}_$Date.log"
$UserSID = Get-CurrentUserSID
$CurrentUserName = (Get-Process -IncludeUserName -Name explorer | Select-Object -First 1 | Select-Object -ExpandProperty UserName).Split("\")[1]

# Staging folder setup
Write-LogEntry -Value "Initiating setup process" -Severity 1

if (Test-Path "$stagingFolderVar") {
	Remove-Item -Path "$($stagingFolderVar)\*" -Recurse -Force -ErrorAction SilentlyContinue
}

$SetupFolder = (New-Item -ItemType "directory" -Path $stagingFolderVar -Name $AppName -Force).FullName
Write-LogEntry -Value "Setup folder has been created at: $Setupfolder." -Severity 1

# Install new templates
if ($Mode -eq "Install") {

    # Remove existing templates
    if (Test-Path "$env:ProgramData\HUD Templates") {
        Remove-HudTemplates
        Write-LogEntry -Value "Removed Template Files, Document Building Blocks files, Registry keys, and App Validation File" -Severity 1
    }

    $TemplatesPath = "$env:ProgramData\HUD Templates"
    $DocumentBuildingBlocksPath = "C:\Users\$CurrentUserName\AppData\Roaming\Microsoft\DocumentBuildingBlocks\1033\16"

    # Create directories
    New-Item -ItemType Directory -Path $TemplatesPath -Force -Confirm:$false
    New-Item -ItemType Directory -Path $DocumentBuildingBlocksPath -Force -Confirm:$false

    # Copy template files
    Copy-Item -Path "$PSScriptRoot\Installer\*" -Destination $TemplatesPath -Recurse -Force
    Copy-Item -Path "$PSScriptRoot\CoverPage\*" -Destination $DocumentBuildingBlocksPath -Recurse -Force

    # Set registry properties
    $RegistryPath = "Registry::HKEY_USERS\$UserSID\SOFTWARE\Microsoft\Office\16.0"
    New-ItemProperty -Path "$RegistryPath\Common\General" -Name "SharedTemplates" -PropertyType String -Value $TemplatesPath -Force -Confirm:$false
    New-ItemProperty -Path "$RegistryPath\Word\Options" -Name "officestartdefaulttab" -PropertyType Dword -Value 1 -Force -Confirm:$false
    New-ItemProperty -Path "$RegistryPath\PowerPoint\Options" -Name "officestartdefaulttab" -PropertyType Dword -Value 1 -Force -Confirm:$false
    New-ItemProperty -Path "$RegistryPath\Word\Options" -Name "HUDOfficeTemplates" -PropertyType ExpandString -Value $TemplatesPath -Force -Confirm:$false
    New-ItemProperty -Path "$RegistryPath\PowerPoint\Options" -Name "HUDOfficeTemplates" -PropertyType ExpandString -Value $TemplatesPath -Force -Confirm:$false
    
    # Create validation file
    New-Item -ItemType File -Path $AppValidationFile -Force -Value $AppVersion
    Write-LogEntry -Value "Validation file has been created at $AppValidationFile" -Severity 1
    
    # Finished installing new templates
    Write-LogEntry -Value "Finished installing new templates version $($version)" -Severity 1

} elseif ($mode -eq "Uninstall") {
    Remove-HudTemplates
    Write-LogEntry -Value "Removed Template Files, Document Building Blocks files, Registry keys, and App Validation File" -Severity 1
}

# Cleanup 
if (Test-Path "$SetupFolder") {
    Remove-Item -Path "$SetupFolder" -Recurse -Force -ErrorAction Continue
    Write-LogEntry -Value "Cleanup completed successfully" -Severity 1
}