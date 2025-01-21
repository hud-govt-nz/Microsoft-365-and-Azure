function Get-InstalledApps {
    <#
    .SYNOPSIS
    Returns information about installed applications.

    .PARAMETER App
    An array of strings representing the names of the applications to search for.

    .RETURN
    An array of objects containing information about the installed applications, with properties DisplayName, DisplayVersion, and UninstallString.
    #>
    param (
        [string[]]$App
    )

    [array]$Installed = Get-ItemProperty -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -and $_.DisplayName -ne '' } | Select-Object DisplayName, DisplayVersion, UninstallString
    $Installed += Get-ItemProperty -Path HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -and $_.DisplayName -ne '' } | Select-Object DisplayName, DisplayVersion, UninstallString

    [array]$SelectedApp = @()
    foreach ($item in $App) {
        [array]$tempResult = $Installed | Where-Object { $_.DisplayName -match $item }
        $SelectedApp += @($tempResult)
    }

    return $SelectedApp
}