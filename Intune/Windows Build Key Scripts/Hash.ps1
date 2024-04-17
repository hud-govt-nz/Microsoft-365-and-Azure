function Export-DeviceHash {
    <#
    .SYNOPSIS
    Exports the device HASH to a CSV file with the device model and serial number as part of the file name.

    .RETURN
    Boolean value indicating whether the export was successful.
    #>

    Set-ExecutionPolicy Bypass -Scope Process -Force -Confirm:$false
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.208 -Force -Confirm:$False | Out-Null
    Install-Module -Name PSWindowsUpdate -Scope AllUsers -Force -Confirm:$False | Out-Null
    if (!(Get-Module -Name "Get-WindowsAutoPilotInfo")) {
        Install-Script -Name "Get-WindowsAutoPilotInfo" -Force
    }

    $autoPilotInfo = Get-WindowsAutoPilotInfo
    $serialNumber = [string](Get-WmiObject -Class Win32_BIOS).SerialNumber
    $model = [string](Get-WmiObject -Class Win32_ComputerSystem).Model
    $deviceID = (Get-WmiObject Win32_LogicalDisk | Where-Object {$_.DeviceID -like "W11_*"}).DeviceID

    $csvFileName = "$deviceID\HashFiles\Individual\$model - $serialNumber.csv"
    $mergedHash = "$deviceID\HashFiles\Merged\Merged.csv"

    if (-not (Test-Path -Path (Split-Path -Path $csvFileName) -PathType Container)) {
        New-Item -ItemType Directory -Path (Split-Path -Path $csvFileName)
    }

    if (-not (Test-Path -Path (Split-Path -Path $mergedHash) -PathType Container)) {
        New-Item -ItemType Directory -Path (Split-Path -Path $mergedHash)
    }

    $autoPilotInfo | Export-Csv -Path $csvFileName -NoTypeInformation -Force

    $path = "$deviceID\HashFiles\Individual\*.csv"
    Import-CSV -Path (Get-ChildItem $path) | ConvertTo-Csv -NoTypeInformation | ForEach-Object {$_ -replace '"', ''} | Out-File $mergedHash

    return $true
}