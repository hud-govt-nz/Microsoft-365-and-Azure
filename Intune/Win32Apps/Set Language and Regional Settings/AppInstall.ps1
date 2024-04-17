<#
.APP: HUD - Language and Region Settings
.AUTHOR: Ashley Forde
.DATE: 23 May 2023
#>

# Root Folder
$Directory = 'HUD'

# Define Log function
function Write-Log {
    param(
        [string]$Path,
        [string]$Value
        )
    Add-Content -Path $Path -Value $Value
    }

# Create Directories
$HomeFolder = "$($env:homedrive)\$Directory"
New-Item -Path $HomeFolder -Name 00_Staging -ItemType "directory" -Force -Confirm:$false | Out-Null
    if (Test-Path -Path $HomeFolder) { 
        "Path exists!"
        } else { 
            "Creating root folder..."
            New-Item -Path $HomeFolder -ItemType "directory" -Force -Confirm:$false | Out-Null
            
            foreach($subFolder in "00_Staging", "01_Logs", "02_Validation") {
                New-Item -Path "$HomeFolder\" -Name $subFolder -ItemType "directory" -Force -Confirm:$false | Out-Null
                }
            }

#Set application details
$path = "$HomeFolder\00_Staging"
$logs = "$HomeFolder\01_Logs"
$validation = "$HomeFolder\02_Validation"
$AppName=[string]'HUD - Language and Region Settings'
$AppVersion="2.0"
$AppValidationFile="$validation\$AppName.txt"
$AppLog="$logs\$AppName.log"

# 1. Install Languages on Device
$Languages =@('en-US','en-NZ','mi-latn', 'en-GB')
$Languages | ForEach-Object { 
    if ($_ -eq 'en-US'){
        Write-Log -Path $AppLog -Value "[$(Get-Date)] Installing en-US language and copying to default settings"
        Install-Language $_ -CopyToSettings}
        else {
            Write-Log -Path $AppLog -Value "Installing $_ language"
            Install-Language $_}
    }

# 2. Reset Language Bar Settings to Default
$languageBarOption = Get-WinLanguageBarOption
    if ($languageBarOption.IsLegacyLanguageBar -eq 'true') {
        Set-WinLanguageBarOption
        Write-Log -Path $AppLog -Value "[$(Get-Date)] Language bar settings reset"
    }

# 3. Set Windows User Lanauge List
$userLanguageList = Get-WinUserLanguageList
$tags =@()
    foreach ($item in $userLanguageList) {
        $tags += @($item.LanguageTag)
    }
    Write-Log -Path $AppLog -Value "[$(Get-Date)] Current Listed Languages: $tags"
    Write-Log -Path $AppLog -Value "[$(Get-Date)] Setting Order"
    Set-WinUserLanguageList -LanguageList en-NZ, en-US, en-GB, mi-latn -Force -Confirm:$false

$userLanguageList = Get-WinUserLanguageList
Write-Log -Path $AppLog -Value "[$(Get-Date)] Language Lists updated, pending restart."
$userLanguageList | ForEach-Object {
    Write-Log -Path $AppLog -Value "Language Tag: $($_.LanguageTag)"
    Write-Log -Path $AppLog -Value "Autonym: $($_.Autonym)"
    Write-Log -Path $AppLog -Value "English Name: $($_.EnglishName)"
    }

# 4. Set Country or Region to New Zealand
$systemLocale = Get-WinSystemLocale
if ($systemLocale.Name -ne 'en-NZ') {
    Set-WinSystemLocale -SystemLocale "en-NZ"
    Write-Log -Path $AppLog -Value "[$(Get-Date)] System locale set to New Zealand"
    }

# 5. Set Home Location
$homeLocation = Get-WinHomeLocation
if ($homeLocation.HomeLocation -ne 'New Zealand') {
    Set-WinHomeLocation -GeoId 183
    Write-Log -Path $AppLog -Value "[$(Get-Date)] Home location set to New Zealand"
    }

# 6. Set Windows UI Language Override
$uiLanguageOverride = Get-WinUILanguageOverride
if ($uiLanguageOverride.Name -ne 'en-NZ') {
    Set-WinUILanguageOverride -Language en-NZ
    Write-Log -Path $AppLog -Value "[$(Get-Date)] Override language set to New Zealand"
    }

# 7. Set Culture
$currentCulture = Get-Culture
if ($currentCulture.Name -ne 'en-NZ') {
    Set-Culture -CultureInfo en-NZ
    Write-Log -Path $AppLog -Value "[$(Get-Date)] Culture set to New Zealand"
    }

# 8. Set Default Input Override to New Zealand
$defaultInputMethodOverride = Get-WinDefaultInputMethodOverride
if ($defaultInputMethodOverride.InputMethodTip -ne '1409:00001409') {
    Set-WinDefaultInputMethodOverride -InputTip "1409:00001409"
    Write-Log -Path $AppLog -Value "[$(Get-Date)] Default keyboard layout set to en-US"
    }

# 9. Copy User International Settings to System
Copy-UserInternationalSettingsToSystem -WelcomeScreen $True -NewUser $True
Write-Log -Path $AppLog -Value "[$(Get-Date)] Settings copied to system defaults"

# 10. Disable Beta: Use Unicode UTF-8 for Worldwide Language Support. Setting breaks calendar appoints in Outlook Desktop and causes test to appear as random characters. 

$regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Nls\CodePage"
$originalValues = @{
    "ACP" = "1252"
    "OEMCP" = "437"
    "MACCP" = "10000"
    }

$registryValues = Get-ItemProperty -Path $regPath

foreach ($entry in $originalValues.Keys) {
    $currentValue = $registryValues.$entry
    $originalValue = $originalValues[$entry]

    if ($currentValue -eq "65001") {
        Set-ItemProperty -Path $regPath -Name $entry -Value $originalValue
        Write-Log -Path $AppLog -Value "[$(Get-Date)] Value '$entry' changed from '65001' to '$originalValue'."
        Write-Log -Path $AppLog -Value "[$(Get-Date)] Beta: Use Unicode UTF-8 for worldwide language support has been disabled"
    } elseif ($currentValue -eq $originalValue) {
        Write-Log -Path $AppLog -Value "[$(Get-Date)] Value '$entry' is already set to '$originalValue'. No changes made."
    } else {
        Write-Log -Path $AppLog -Value "[$(Get-Date)] Value '$entry' is set to '$currentValue', which differs from the original value of '$originalValue'. No changes made."
        }
    }

try {
    # Create validation file
    New-Item -ItemType File -Path $AppValidationFile -Force -Value $AppVersion | Out-Null
    Write-Log -Path $AppLog -Value "[$(Get-Date)] Validation file was created successfully."
    } catch {
        Write-Log -Path $AppLog -Value "[$(Get-Date)] Error creating validation file: $_"
        exit 1
        }


# Restart required for settings to take effect.
#Shutdown /r /t 30























# Check if the install or uninstall switch is used
switch ($args[0]) {
    'install' {
        if ($Result[0]) {
            Write-Log -Path $AppLog -Value "[$(Get-Date)] $AppName is currently installed with version $AppVersion."
            } else {
                try {
                    #  Copy Installer to local device
                    Copy-Item -Path ".\Installer\*" -Destination $path -Recurse -Force
                    $outfile = "$Path\$AppInstallFile"
                    } catch {
                        Write-Log -Path $AppLog -Value "[$(Get-Date)] Error copying installer: $_"
                        exit 1
                        }                
                try {
                    # Install App
                    $installProcess = Start-Process $outfile -ArgumentList $AppInstallArguments -PassThru -Wait -ErrorAction Stop
                    $exitCode = $installProcess.ExitCode

                    # Log the result of the installation
                    if ($exitCode -eq 0) {
                        Write-Log -Path $AppLog -Value "[$(Get-Date)] $AppName version $AppVersion was installed successfully with exit code $($exitCode)"
                        } else {
                            Write-Log -Path $AppLog -Value "[$(Get-Date)] $AppName version $AppVersion was not installed successfully with exit code $($exitCode)" 
                            exit $exitCode
                            }
                    } catch {
                        Write-Log -Path $AppLog -Value "[$(Get-Date)] Error installing App: $_"
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
                try {
                    # Delete installer files
                    Remove-Item -Path "$path\*" -Recurse -Force -ErrorAction Stop
                    Write-Log -Path $AppLog -Value "[$(Get-Date)] Installer files were deleted successfully."
                    } catch {
                        Write-Log -Path $AppLog -Value "[$(Get-Date)] Error deleting installer files: $_"
                        exit 1
                        }
                }
        }

    'uninstall' {
        if ($Result) {
            try {
                # Uninstall App
                if ($Result.uninstallstring -like "msiexec*"){
                    $Result = (($Result.UninstallString -split ' ')[1] -replace '/I','/X ') + ' /q'
                    $uninstallProcess = Start-Process $uninstall_command -ArgumentList $uninstall_args -PassThru -Wait -ErrorAction Stop
                    $exitCode = $uninstallProcess.ExitCode
                    Write-Log -Path $AppLog -Value "[$(Get-Date)] $AppName version $AppVersion was uninstalled successfully with exit code $($exitCode)"
                    } else {
                        $uninstall_command = (($Result.UninstallString -split '\"')[1])
                        $uninstall_args = (($Result.UninstallString -split '\"')[2]) + '/S'
                        $uninstallProcess = Start-Process $uninstall_command -ArgumentList $uninstall_args -PassThru -Wait -ErrorAction Stop
                        $exitCode = $uninstallProcess.ExitCode
                        Write-Log -Path $AppLog -Value "[$(Get-Date)] $AppName version $AppVersion was uninstalled successfully with exit code $($exitCode)"
                        }
                } catch {
                    Write-Log -Path $AppLog -Value "[$(Get-Date)] Error uninstalling App: $_"
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
        }

    default {
        Write-Log -Path $AppLog -Value "[$(Get-Date)] Invalid argument. Please specify 'install' or 'uninstall'."
        exit 1
    }
}
