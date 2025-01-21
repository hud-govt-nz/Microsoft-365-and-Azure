function New-LocalAdminUser {
    <#
    .SYNOPSIS
    Creates and updates a local administrator user.
    .PARAMETER userName
    The name of the user to be created or updated. (string)
    .PARAMETER password
    The password for the user. (string)
    .PARAMETER description
    The description of the user. (string)
    .RETURN
    The updated local user account. (LocalUser)
    #>
    param(
        [string]$userName,
        [string]$password,
        [string]$description
    )

    $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
    New-LocalUser -Name $userName -Password $securePassword -Description $description
    Add-LocalGroupMember -Group "Administrators" -Member $userName

    $account = Get-LocalUser -Name $userName
    $account | Set-LocalUser -Password $securePassword
}