#Fileext
$regkey="HKLM:\SOFTWARE\Policies\Adobe\Adobe Acrobat\DC\FeatureLockDown"
$name="bEnableGentech"
$value=0

#Registry Template
If (!(Test-Path $regkey))
{
    New-Item -Path $regkey -ErrorAction stop
}

if (!(Get-ItemProperty -Path $regkey -Name $name -ErrorAction SilentlyContinue))
{
    New-ItemProperty -Path $regkey -Name $name -Value $value -PropertyType DWORD -ErrorAction stop
    write-output "remediation complete"
    exit 0
}

set-ItemProperty -Path $regkey -Name $name -Value $value -ErrorAction stop
write-output "remediation complete"
exit 0