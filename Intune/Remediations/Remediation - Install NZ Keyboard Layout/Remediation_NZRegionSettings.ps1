try {
    
    #Set PS Drive for HKEY_Users and Obtain Current User System Identifier
    New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS -Scope Global | Out-Null
    $currentUser = (Get-Process -IncludeUserName -Name explorer | Select-Object -First 1 | Select-Object -ExpandProperty UserName).Split("\")[1] 
    $Keys = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" -Recurse
    Foreach($Key in $Keys) {
        if(($key.GetValueNames() | ForEach-Object{$key.GetValue($_)}) -match $CurrentUser ){
            $sid = $key
            }
        }

    #SID for current user
    $UserSID = $sid.pschildname 

    Write-Host "Script is running in system context against user $currentUser " -ForegroundColor Green

    #Update Date & time formatting
    $ReigonKeyPath = "HKU:\$UserSID\Control Panel\International"
        
    #Set values
    Set-ItemProperty -Path $ReigonKeyPath -Name sCountry -Value "New Zealand" -Force
    Set-ItemProperty -Path $ReigonKeyPath -Name sShortDate -Value "M/d/yyyy" -Force
    Set-ItemProperty -Path $ReigonKeyPath -Name sLongDate -Value "dddd, MMMM d, yyyy" -Force
    Set-ItemProperty -Path $ReigonKeyPath -Name sShortTime -Value "h:mm tt" -Force
    Set-ItemProperty -Path $ReigonKeyPath -Name sTimeFormat -Value "h:mm:ss tt" -Force
    Set-ItemProperty -Path $ReigonKeyPath -Name iFirstDayOfWeek -Value "0" -Force
        
    #Display Values in terminal
    $sCountry = (Get-ItemProperty -Path $ReigonKeyPath -Name sCountry).sCountry
    $sShortDate = (Get-ItemProperty -Path $ReigonKeyPath -Name sShortDate).sShortDate
    $sLongDate = (Get-ItemProperty -Path $ReigonKeyPath -Name sLongDate).sLongDate
    $sShortTime = (Get-ItemProperty -Path $ReigonKeyPath -Name sShortTime).sShortTime
    $sTimeFormat = (Get-ItemProperty -Path $ReigonKeyPath -Name sTimeFormat).sTimeFormat
    $iFirstDayOfWeek = (Get-ItemProperty -Path $ReigonKeyPath -Name iFirstDayOfWeek).iFirstDayOfWeek
        
    $Obj = New-Object -TypeName PSObject -Property @{
        "Country" = $sCountry
        "Short date" = $sShortDate
        "Long date" = $sLongDate
        "Short time" = $sShortTime
        "Long time" = $sTimeFormat
        "First day of week" = $iFirstDayOfWeek
        }
        
    Write-Host "The current date and time formats:"
    $Obj
        

    #Manually update keyboard preload layout
    $array =@()
    $KeyLayoutPath = "HKU:\$UserSID\Keyboard Layout\Preload" 
    $Registrykeys = Get-Item -Path $KeyLayoutPath
      
    #Force remove en-GB keyboard layout from Registry
    $Registrykeys | Select-Object -ExpandProperty Property | ForEach-Object {
        $name = $_
        $data = "" | Select-Object -Property Name, Path, Type, Data
        $data.Name = $name
        $data.Type = $Registrykeys.GetValueKind($name)
        $data.Data = $Registrykeys.GetValue($name)
        $data.Path = "Registry::$Registrykeys"
        $array += $data
        }

    Foreach ($item in $array) { 
        if ($item.data -match "00000809") {
            Write-Host "UK Keyboard layout detected, removing registry key" -ForegroundColor Yellow
            Remove-ItemProperty -Path $KeyLayoutPath -Name $item.Name -Force -Confirm:$false -Verbose
            } 
        }

    $array = $null | Out-Null

    #Set the Default System UI Language
    $sysUIPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Nls\Language"
    $SysUIValue = (Get-ItemProperty -Path $sysUIPath -Name InstallLanguage).installlanguage

    
    if ($SysUIValue -match "0409") {
        Write-Host "en-US is already set as the system language" -ForegroundColor Yellow
    } else {
        Write-Host "en-US is not set as default system language, updating..." -ForegroundColor Green
        Set-ItemProperty -Path $sysUIPath -Name Default -Value 0409 -Force -Confirm:$false -Verbose
        Set-ItemProperty -Path $sysUIPath -Name InstallLanguage -Value 0409 -Force -Confirm:$false -Verbose
        Set-ItemProperty -Path $sysUIPath -Name InstallLanguageFallback -Value en-US -Force -Confirm:$false -Verbose

    }

    #Remove EnGB Language Pack from Device
    Write-Host "Checking for en-gb language pack..." -ForegroundColor Yellow
    $ENGB = Get-AppxPackage -allusers *LanguageExperiencePacken-GB* -Verbose
    
    if ($ENGB -ne $null) {
        Write-Host "En-GB Language Pack found, removing" -ForegroundColor Yellow
        Remove-AppxPackage -AllUsers -Package $Pack.PackageFullName -Confirm:$false -Verbose
        Write-Host "en-GB $($Pack) language pack removed" -ForegroundColor Green
        } else {
            Write-Host "en-GB is not installed on this device" -ForegroundColor Yellow  
            }
    #Set User and System Locale to New Zealand and Include US Keyboard Layout
    Set-WinUILanguageOverride -Language en-US -Verbose 
    Set-WinUserLanguageList -LanguageList en-US, en-NZ, mi-latn -Force -Confirm:$false -Verbose 
    Set-Culture -CultureInfo en-NZ -Verbose
    Set-WinHomeLocation -GeoId 183 #nz

    #Double check and Remove English GB Language Pack off Device
    $LangList = Get-WinUserLanguageList
    $MarkedLang = $LangList | Where-Object LanguageTag -eq "en-GB"
    $LangList.Remove($MarkedLang)
    Set-WinUserLanguageList $LangList -Force -Verbose

    #Timezone for Computer:
    Set-TimeZone -name "New Zealand Standard Time" -Verbose

    #Finish
    Write-Output "en-GB user languages removed and en-NZ and Maori updated..."
    Remove-PSDrive HKU

    Exit 0
    } catch {
    $errMsg = $_.exeption.essage
    Write-Output $errMsg
    }