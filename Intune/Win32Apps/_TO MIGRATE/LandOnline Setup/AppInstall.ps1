param(
    [Parameter(Mandatory=$true)]
    [ValidateSet("Install","Uninstall")]
    [string]$Mode
)

# Root Folder
$Directory = 'HUD'
$HomeFolder = "$($env:homedrive)\$Directory"
$path = "$HomeFolder\00_Staging"
$logs = "$HomeFolder\01_Logs"
$validation = "$HomeFolder\02_Validation"
$AppName=[string]'Land Online Set'
$AppVersion="1.0"
$AppValidationFile="$validation\$AppName.txt"
$AppLog="$logs\$AppName.log"

# Define Log function
function Write-Log {
    param(
        [string]$Path,
        [string]$Value
    )
    Add-Content -Path $Path -Value $Value
}

# Centralized Error Handling Function
Function Handle-Error {
    param(
        [string]$Message,
        [int]$ExitCode = 1
    )
    Write-Log -Path $AppLog -Value "[$(Get-Date)] ERROR: $Message"
    exit $ExitCode
}

# Check if App is Installed
function Check-InstalledApps {
    param (
        [Parameter(Mandatory=$true)]
        [array]$AppName
    )

    # Check if Application Already Exists
    $Installed = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, UninstallString
    $Installed += Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, UninstallString

    $Result = @()
    foreach ($item in $AppName) {
        $tempResult = $Installed | Where-Object { $_.DisplayName -ne $null } | Where-Object { $_.DisplayName -match $item }
        $Result += @($tempResult)
    }

    return $Result
}

# Manage App
Function Manage-App {
    param (
        [string]$appName,
        [string]$appVersion,
        [string]$appPath,
        [string]$appArgs,
        [int]$timeout,
        [string]$uninstallCommand,
        [string]$uninstallArgs,
        [switch]$SkipIfExists = $false
    )
    
    $dateStamp = Get-Date
    # Check if the app is already installed
    $existingApp = Check-InstalledApps -AppName @($appName)

    if ($existingApp.Count -gt 0 -and $SkipIfExists) {
        Write-Log -Path $AppLog -Value "[$dateStamp] $appName is already installed. Skipping..."
        return
    }

    try {
        if ($Mode -eq 'Install') {
            Write-Log -Path $AppLog -Value "[$dateStamp] Installing App: $appName $appVersion"
            $process = Start-Process -FilePath $appPath -ArgumentList $appArgs -PassThru -Wait -ErrorAction Stop
          
        } elseif ($Mode -eq 'Uninstall') {
            Write-Log -Path $AppLog -Value "[$dateStamp] Uninstalling App: $appName $appVersion"
            $process = Start-Process -FilePath $uninstallCommand -ArgumentList $uninstallArgs -PassThru -Wait -ErrorAction Stop
        }
        
        $exitCode = $process.ExitCode
        if ($exitCode -eq 0) {
            Write-Log -Path $AppLog -Value "[$dateStamp] $appName version $appVersion was ${Mode}ed successfully with exit code $exitCode"
        } else {
            Write-Log -Path $AppLog -Value "[$dateStamp] $appName version $appVersion was not ${Mode}ed successfully with exit code $exitCode"
            exit $exitCode
        }
    } catch {
        Write-Log -Path $AppLog -Value "[$dateStamp] Error ${Mode}ing App: $_"
        exit 1
    }
}

# Create Directories
if (Test-Path -Path $HomeFolder) { 
    "Path exists!"
} else { 
    "Creating root folder..."
    New-Item -Path $HomeFolder -ItemType "directory" -Force -Confirm:$false | Out-Null
    if (-not $?) {
        Handle-Error -Message "Failed to create $HomeFolder"
    }

    foreach ($subFolder in "00_Staging", "01_Logs", "02_Validation") {
        New-Item -Path "$HomeFolder\" -Name $subFolder -ItemType "directory" -Force -Confirm:$false | Out-Null
        if (-not $?) {
            Handle-Error -Message "Failed to create sub-folder $subFolder under $HomeFolder"
        }
    }
}

#$appsToCheck = @("DC Loader*", "LandOnline*", "Remote*")
#$result = Check-InstalledApps -AppName $appsToCheck

if ($Mode -eq 'Install') {
    # Copy installer files (only during install)
    try {
        Copy-Item -Path "$PSScriptRoot\Installer\*" -Destination $path -Recurse -Force
        Write-Log -Path $AppLog -Value "[$(Get-Date)] Installer files were copied successfully."
    } catch {
        Handle-Error -Message "Error copying installer Files: $_"
    }

    # Install Applications
    # Note: Added -SkipIfExists to each call to skip installation if already installed
    Manage-App -appName "DC Loader" -appVersion "3.1.13" -appPath "$Path\dcloader3.1.13.exe" -appArgs "/qb" -SkipIfExists
    Manage-App -appName "LandOnline Print-to-Tiff Driver" -appVersion "3.03" -appPath "$Path\LandOnlinePrintToTiff_x64.msi" -appArgs "/qb" -SkipIfExists
    Manage-App -appName "Remote Access Tool - Toitu Te Whenua" -appVersion "7.11.760" -appPath "$Path\Remote Access Installer - ToituTeWhenua.msi" -appArgs "/qb" -SkipIfExists
      
    # Create validation file
    try {
        New-Item -ItemType File -Path $AppValidationFile -Force -Value $AppVersion
        Write-Log -Path $AppLog -Value "[$(Get-Date)] Validation file was created successfully."
    } catch {
        Handle-Error -Message "Error creating validation file: $_"
    }
  
    # Delete installer files
    try {
        Remove-Item -Path "$path\*" -Recurse -Force
        Write-Log -Path $AppLog -Value "[$(Get-Date)] Installer files were deleted successfully."
    } catch {
        Handle-Error -Message "Error deleting installer files: $_"
    }

} elseif ($Mode -eq 'Uninstall') {
    # Uninstall Applications
    Manage-App -appName "DC Loader" -appVersion "3.1.13" -uninstallCommand "C:\ProgramData\Caphyon\Advanced Installer\{D617A396-009B-44D3-B436-8773A73E5704}\DC Loader 3.1.exe" -uninstallArgs "/x {D617A396-009B-44D3-B436-8773A73E5704} /qb AI_UNINSTALLER_CTP=1"
    Manage-App -appName "LandOnline Print to TIFF Driver" -appVersion "3.03" -uninstallCommand "MsiExec.exe" -uninstallArgs "/X{23D62982-92C7-4219-B53F-53E058226854} /qb"
    Manage-App -appName "Remote Access Tool - Toitu Te Whenua" -appVersion "7.11.760" -uninstallCommand "MsiExec.exe" -uninstallArgs "/X{BCE7D01B-D863-9EF6-EA14-7A6741FA6CD5} /qb"

    # Delete validation file
    try {
        Remove-Item -Path $AppValidationFile -Force
        Write-Log -Path $AppLog -Value "[$(Get-Date)] Validation file was deleted successfully."
    } catch {
        Handle-Error -Message "Error deleting validation file: $_"
    }
}

