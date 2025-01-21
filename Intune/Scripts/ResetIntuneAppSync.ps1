$UserAccountStatusPath = "HKLM:SOFTWARE\Microsoft\IntuneManagementExtension\SideCarPolicies\StatusServiceReports"
$UserObjectIDs = Get-ChildItem -Path $UserAccountStatusPath

$UserObjectIDs | ForEach-Object {

    # Obtain User Object ID
    $User = $_.PSChildName


    # User Paths
    $GRSPath = "HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps\$User\GRS"
    $Win32AppPath = "HKLM:SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps\$User"
    $SideCarPath = "HKLM:SOFTWARE\Microsoft\IntuneManagementExtension\SideCarPolicies\StatusServiceReports\$User"
    $OpStatePath = "HKLM:SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps\OperationalState\$User"
    $ReportingPath = "HKLM:SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps\Reporting\$User"

    # Remove User Keys
    Get-item  -Path $GRSPath | Remove-Item -Recurse -Force -Verbose
    Get-item  -Path $Win32AppPath | Remove-Item -Recurse -Force -Verbose
    Get-item  -Path $SideCarPath | Remove-Item -Recurse -Force -Verbose
    Get-item  -Path $OpStatePath | Remove-Item -Recurse -Force -Verbose
    Get-item  -Path $ReportingPath | Remove-Item -Recurse -Force -Verbose

    
    # Clean System keys
    $SysOpStatePath = "HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps\OperationalState\00000000-0000-0000-0000-000000000000\"
    $SysReportingPath = "HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps\Reporting\00000000-0000-0000-0000-000000000000\"

    Get-item  -Path $SysOpStatePath | Remove-Item -Recurse -Force -Verbose
    Get-item  -Path $SysReportingPath | Remove-Item -Recurse -Force -Verbose


}

# Empty IME Incoming folder
Remove-Item -Path 'C:\Program Files (x86)\Microsoft Intune Management Extension\Content\Incoming\*' -Filter * -Recurse -Verbose

# Restart the IME Service
Get-Service -DisplayName "Microsoft Intune Management Extension" | restart-Service -Verbose

Start-Sleep 10 -Verbose

# Run Sync
$Shell = New-Object -ComObject Shell.Application
$Shell.open("intunemanagementextension://syncapp")

