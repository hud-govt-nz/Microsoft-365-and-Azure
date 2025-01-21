$regkey = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\General"
$regname = "HideNewOutlookToggle"
$regvalue = "1"

try
{
    
    IF(!(Test-Path $Regkey))
    {
    New-Item -Path $Regkey -Force -ErrorAction SilentlyContinue | Out-Null
    New-ItemProperty -Path $Regkey -Name $regname -Value $regvalue -PropertyType DWORD -Force -ErrorAction SilentlyContinue | Out-Null}
    ELSE{
    New-ItemProperty -Path $Regkey -Name $regname -Value $regvalue -PropertyType DWORD -Force -ErrorAction SilentlyContinue | Out-Null}
    Write-Host "Setting value to hide the try new Outlook app preview button"
    exit 0 
}
catch{
    $errMsg = $_.Exception.Message
    Write-Error $errMsg
    exit 1
}