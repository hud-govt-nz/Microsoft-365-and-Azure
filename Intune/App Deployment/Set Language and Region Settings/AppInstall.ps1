<#
.SYNOPSIS
    Language and Region Settings

.DESCRIPTION
    Script to install Language and Region Settings

.PARAMETER Mode
    Sets the mode of operation. Supported modes are Install or Uninstall

.EXAMPLE 
    powershell.exe -executionpolicy bypass -file .\appinstall.ps1 -Mode Install
    powershell.exe -executionpolicy bypass -file .\appinstall.ps1 -Mode Uninstall

.NOTES
    - AUTHOR: 
    - Version: 
    - Date: 
#>
# Region Parameters
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("Install")]
    [string]$Mode = "Install"
)

# Reference functions.ps1
. "$PSScriptRoot\functions.ps1"
# Application Variables
$AppName = "Language and Region Settings"
$AppVersion = "2.0"

# Initialize Directories
$folderpaths = Initialize-Directories -HomeFolder C:\HUD\

# Template Variables
$Date = Get-Date -Format "MM-dd-yyyy"
$stagingFolderVar = $folderPaths.StagingFolder
$logsFolderVar = $folderPaths.LogsFolder
$LogFileName = "$($AppName)_${Mode}_$Date.log"
$validationFolderVar = $folderPaths.ValidationFolder
$AppValidationFile = "$validationFolderVar\$AppName.txt"

# Begin Setup
Write-LogEntry -Value "Initiating script" -Severity 1

# Install
try {
    # Define the languages to install
    $Languages = @('en-US', 'en-NZ', 'mi-latn', 'en-GB')

    # Function to install language and log the installation
    function InstallLanguage($language) {
        Write-LogEntry -Value "Installing $language language" -Severity 1
        Install-Language $language
    }

    # 1. Install Languages on Device
    $Languages | ForEach-Object {
        if ($_ -eq 'en-US') {
            Write-LogEntry -Value "Installing en-US language and copying to default settings" -Severity 1
            Install-Language $_ -CopyToSettings
        }
        else {
            InstallLanguage $_
        }
    }

    # 2. Reset Language Bar Settings to Default
    $languageBarOption = Get-WinLanguageBarOption
    if ($languageBarOption.IsLegacyLanguageBar -eq 'true') {
        Set-WinLanguageBarOption
        Write-LogEntry -Value "Language bar settings reset" -Severity 1
    }

    # 3. Set Windows User Language List
    $userLanguageList = Get-WinUserLanguageList
    $tags = $userLanguageList.LanguageTag
    Write-LogEntry -Value "Current Listed Languages: $tags" -Severity 1
    Write-LogEntry -Value "Setting Order" -Severity 1
    Set-WinUserLanguageList -LanguageList 'en-NZ', 'en-US', 'en-GB', 'mi-latn' -Force -Confirm:$false

    $userLanguageList = Get-WinUserLanguageList
    $userLanguageList | ForEach-Object {
        Write-LogEntry -Value "Language Tag: $($_.LanguageTag)" -Severity 1
        Write-LogEntry -Value "Autonym: $($_.Autonym)" -Severity 1
        Write-LogEntry -Value "English Name: $($_.EnglishName)" -Severity 1
    }

    # 4. Set Country or Region to New Zealand
    $systemLocale = Get-WinSystemLocale
    if ($systemLocale.Name -ne 'en-NZ') {
        Set-WinSystemLocale -SystemLocale "en-NZ"
        Write-LogEntry -Value "System locale set to New Zealand" -Severity 1
    }

    # 5. Set Home Location
    $homeLocation = Get-WinHomeLocation
    if ($homeLocation.HomeLocation -ne 'New Zealand') {
        Set-WinHomeLocation -GeoId 183
        Write-LogEntry -Value "Home location set to New Zealand" -Severity 1
    }

    # 6. Set Windows UI Language Override
    $uiLanguageOverride = Get-WinUILanguageOverride
    if ($uiLanguageOverride.Name -ne 'en-NZ') {
        Set-WinUILanguageOverride -Language en-NZ
        Write-LogEntry -Value "Override language set to New Zealand" -Severity 1
    }

    # 7. Set Culture
    $currentCulture = Get-Culture
    if ($currentCulture.Name -ne 'en-NZ') {
        Set-Culture -CultureInfo en-NZ
        Write-LogEntry -Value "Culture set to New Zealand" -Severity 1
    }

    # 8. Set Default Input Override to New Zealand
    $defaultInputMethodOverride = Get-WinDefaultInputMethodOverride
    if ($defaultInputMethodOverride.InputMethodTip -ne '1409:00001409') {
        Set-WinDefaultInputMethodOverride -InputTip "1409:00001409"
        Write-LogEntry -Value "Default keyboard layout set to en-US" -Severity 1
    }

    # 9. Copy User International Settings to System
    Copy-UserInternationalSettingsToSystem -WelcomeScreen $True -NewUser $True
    Write-LogEntry -Value "Settings copied to system defaults" -Severity 1

    # 10. Disable Beta: Use Unicode UTF-8 for Worldwide Language Support
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
            Write-LogEntry -Value "Value '$entry' changed from '65001' to '$originalValue'." -Severity 1
            Write-LogEntry -Value "Beta: Use Unicode UTF-8 for worldwide language support has been disabled" -Severity 1
        }
        elseif ($currentValue -eq $originalValue) {
            Write-LogEntry -Value "Value '$entry' is already set to '$originalValue'. No changes made." -Severity 1
        }
        else {
            Write-LogEntry -Value "Value '$entry' is set to '$currentValue', which differs from the original value of '$originalValue'. No changes made." -Severity 1
        }
    }

    # Create validation file
    New-Item -ItemType File -Path $AppValidationFile -Force -Value $AppVersion
    Write-LogEntry -Value "Validation file has been created at $AppValidationFile" -Severity 1
    Write-LogEntry -Value "Install of $AppName is complete" -Severity 1
}
catch [System.Exception] {
    Write-LogEntry -Value "Error preparing installation $FileName $($mode). Errormessage: $($_.Exception.Message)" -Severity 3
}
