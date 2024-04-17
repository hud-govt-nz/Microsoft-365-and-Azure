function Get-CurrentUserSID {
    <#
    .SYNOPSIS
    Retrieves the SID of the current user.

    .DESCRIPTION
    Retrieves the SID of the current user by querying the profile list in the registry.

    .RETURN
    [string] The SID of the current user.
    #>
    $currentUser = (Get-Process -IncludeUserName -Name explorer | Select-Object -First 1 | Select-Object -ExpandProperty UserName).Split("\")[1]

    $profileListKeys = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" -Recurse

    foreach ($profileListKey in $profileListKeys) {
        if (($profileListKey.GetValueNames() | ForEach-Object { $profileListKey.GetValue($_) }) -match $currentUser) {
            $sid = $profileListKey.PSChildName
            break
        }
    }

    # Return the SID of the current user
    return $sid
}