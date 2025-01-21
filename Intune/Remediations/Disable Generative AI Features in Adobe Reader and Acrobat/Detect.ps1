$path = "HKLM:\SOFTWARE\Policies\Adobe\Adobe Acrobat\DC\FeatureLockDown"
$name = "bEnableGentech"
$type = "DWORD"
$value = 0

Try {
    $Registry = Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop | Select-Object -ExpandProperty $Name
    If ($Registry -eq $Value){
        Write-Output "Compliant"
        Exit 0
    } 
    else {
    Write-Warning "Not Compliant"
    Exit 1
    }
} 
Catch {
    Write-Warning "Not Compliant"
    Exit 1
}