Param(
    [Switch]$Remove
)

$SP = $PSScriptRoot
$ScriptVersion = "1.0.0.04"
$Client = "HUD"
$PackageName = "HUD Office Templates $ScriptVersion"
$ScriptCommand = '"' + "$env:ProgramData\HUD Templates\Install.bat" + '"'
$ActiveSetupVersion = $ScriptVersion -replace "\.",","
$ActveSetupKey = "OfficeTemplates_d47876f3-b83a-475d-9ce7-72309352bff7"

Function Get-ProcessUser {
    Param([array]$ProcessName="explorer.exe")
    $OUT = @()
    $LogOnCheck = $(Get-CimInstance -ClassName Win32_ComputerSystem -Property username).username
    foreach ($P in $ProcessName){
        if($P -notmatch '\.exe$'){$P="$P.exe"}
        try{
            [array]$UserObject = Get-WmiObject Win32_Process | ?{$_.name -eq $P} | %{$_.GetOwner() | Add-Member -NotePropertyName "SessionId" -NotePropertyValue $_.SessionId -PassThru} | ?{$_.user}
            
            $UserObject | %{
                $objUser = New-Object System.Security.Principal.NTAccount($_.domain, $_.user) 
                $strSID = $objUser.Translate([System.Security.Principal.SecurityIdentifier]) 
                $ExistingUser = $OUT | ?{$strSID -eq $_.SID}
                if($LogOnCheck -eq "$($_.domain)\$($_.user)"){
                    $ActiveUser = $true
                }
                else{
                    $ActiveUser = $false
                }
                if($ExistingUser){
                    $ExistingUser.Process = @($($ExistingUser.Process,$P) | Sort-Object -Unique)
                }
                else{
                    $OUT += New-Object -TypeName psobject -Property @{Username=$_.user;Domain=$_.domain;SessionId=$_.SessionId;SID=$strSID;Process=$P;ActiveUser=$ActiveUser}
                }

            }
        }
        catch{}
    }
    $OUT
}
Function Add-Brand {
    param(
        $Client,
        $PackageName,
        [switch]$remove
    )
    if($remove){
        if(Test-Path "HKLM:\Software\Wow6432Node\$Client\Install\Script"){
            Get-ChildItem -Path "HKLM:\Software\Wow6432Node\$Client\Install\Script" -ErrorAction SilentlyContinue | ?{$_.PSChildName -like $PackageName} | Remove-Item -Recurse -Force
        }
    }
    else{
        New-Item -Path "HKLM:\Software\Wow6432Node\$Client\Install\Script\$PackageName" -Force | Out-Null
        $currDate = Get-Date
        New-ItemProperty -Path "HKLM:\Software\Wow6432Node\$Client\Install\Script\$PackageName" -Name 'Installed' -Value '1' -PropertyType String -Force
        New-ItemProperty -Path "HKLM:\Software\Wow6432Node\$Client\Install\Script\$PackageName" -Name 'Date' -Value $currDate.ToShortDateString() -PropertyType String -Force
        New-ItemProperty -Path "HKLM:\Software\Wow6432Node\$Client\Install\Script\$PackageName" -Name 'Time' -Value $currDate.ToLongTimeString() -PropertyType String -Force
    }
}
Function Is-Admin {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

$Apppath = Get-ItemProperty HKLM:\SOFTWARE\Classes\Applications\winword.exe\shell\edit\command
if ($Apppath.'(default)' -match "\\([^\\]*)\\Winword.exe"){
    $OfficeVersion = $Matches[1]
}
else{
    Write-Host "can't find the office version" -ForegroundColor Yellow
    $OfficeVersion = "NotFound"
}

if(!(Get-PSDrive -Name HKU -ErrorAction SilentlyContinue)){New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS}

if($Remove){
    if(Test-Path "$env:ProgramData\HUD Templates"){
        Remove-Item -Path "$env:ProgramData\HUD Templates" -Recurse -Force
    }
    if(Test-Path "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\$ActveSetupKey"){
        Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\$ActveSetupKey" -Recurse -Force
    }
    $UsersReg = Get-ChildItem -Path HKU:\ -ErrorAction SilentlyContinue |?{$_.PSChildName -match "^S-\d-\d+-(\d+-){1,14}\d+$"} | %{if(Test-Path "HKU:\$($_.PSChildName)\SOFTWARE\Microsoft\Active Setup\Installed Components\$ActveSetupKey"){"HKU:\$($_.PSChildName)\SOFTWARE\Microsoft\Active Setup\Installed Components\$ActveSetupKey"}}
    $UsersReg | ForEach-Object{Remove-Item -Path $_ -Recurse -Force}
    Add-Brand -Client $Client -PackageName $PackageName -remove
}
else{
    if(Is-Admin){
        #SetActiveSetup
        New-Item -Path "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components" -Name $ActveSetupKey -Force | Out-Null
        New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\$ActveSetupKey" -Name "Version" -Value $ActiveSetupVersion -PropertyType string -Force | Out-Null
        New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\$ActveSetupKey" -Name "StubPath" -Value $ScriptCommand -PropertyType string -Force | Out-Null

        #robocopy "$StartPath\HUD Templates\Shared Office Templates" "$env:ProgramData\HUD Templates\Shared Office Templates" /e /purge
        $CurrentUsers = Get-ProcessUser 
        Add-Brand -Client $Client -PackageName $PackageName
    }
    else{
        $CurrentUsers = Get-ProcessUser |Where-Object{$_.Username -eq $env:USERNAME}
    }

    #userSetup
    foreach($U in $CurrentUsers){
        #RegChange
        switch ($OfficeVersion){
            'Office11' {$OficeReg = "HKU:\$($U.SID)\SOFTWARE\Software\Microsoft\Office\11.0"}
            'Office12' {$OficeReg = "HKU:\$($U.SID)\SOFTWARE\Microsoft\Office\12.0"}
            'Office14' {$OficeReg = "HKU:\$($U.SID)\SOFTWARE\Microsoft\Office\14.0"}
            'Office15' {$OficeReg = "HKU:\$($U.SID)\SOFTWARE\Microsoft\Office\15.0"}
            'Office16' {$OficeReg = "HKU:\$($U.SID)\SOFTWARE\Microsoft\Office\16.0"}
            'Office17' {$OficeReg = "HKU:\$($U.SID)\SOFTWARE\Microsoft\Office\17.0"}
            'Office19' {$OficeReg = "HKU:\$($U.SID)\SOFTWARE\Microsoft\Office\19.0"}
            default    {
                if($OfficeVersion -match "Office(.*)"){
                    $OficeReg = "HKU:\$($U.SID)\SOFTWARE\Microsoft\Office\$($Matches[1]).0"
                }
                else{
                    $OficeReg = "HKCU:\SOFTWARE\Microsoft\Office\16.0"
                }
            }
        }
        $WordReg = "$OficeReg\Word\Options"
        $CommonReg = "$OficeReg\Common\General"
        if(!(Test-Path $WordReg)){New-Item -Path $WordReg -Force | Out-Null}
        if(!(Test-Path $CommonReg)){New-Item -Path $CommonReg -Force | Out-Null}

        $WordOptions = Get-ItemProperty $WordReg
        $CommonOptions = Get-ItemProperty $CommonReg

        if($WordOptions.officestartdefaulttab -ne 1){
            
            New-ItemProperty -Name officestartdefaulttab -Value 1 -PropertyType dword -Path $WordReg -Force | Out-Null}

        New-ItemProperty -Name SharedTemplates -Value "$env:ProgramData\HUD Templates\Shared Office Templates" -PropertyType string -Path $CommonReg -Force | Out-Null
    
        New-Item -Path "HKU:\$($U.SID)\SOFTWARE\Microsoft\Active Setup\Installed Components" -Name $ActveSetupKey -Force | Out-Null
        New-ItemProperty -Path "HKU:\$($U.SID)\SOFTWARE\Microsoft\Active Setup\Installed Components\$ActveSetupKey" -Name "Version" -Value $ActiveSetupVersion -PropertyType string -Force | Out-Null
    }
}
