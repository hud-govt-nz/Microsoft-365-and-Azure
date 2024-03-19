<#
    .NAME: HUD Onboarding Script
    .AUTHOR: Ashley Forde
    .VERSION: 1.0
    .PURPOSE: Inital attempt at creating an onboarding script for HUD. 

    .PROCESS: This script will leverage the Microsoft Graph API through powershell to create a new user directly within Azure AD.
        - test if microsoft graph powershell module is installed on machine, installs if it is not present.
        - create a user login
        - assign an administrator role - this can be removed for std. user creation.
        - assign any group memberships 
#>

#Import/Install MS Graph Module
Write-Verbose "Loading Microsoft.Graph module"
$connectCommand = Get-Command Connect-MgGraph -ErrorAction SilentlyContinue #tests if module is installed
    if(!$connectCommand) {
        Write-Warning "Could not find Microsoft.Graph module, trying to install..."
        Install-Module Microsoft.Graph -Scope CurrentUser -ErrorAction Stop #installs module
    } else {
        $latest = Find-Module Microsoft.Graph 
        if($latest.Version -gt $connectCommand.Module.Version) {
            Write-Warning "Microsoft.Graph module version $($connectCommand.Module.Version) installed, trying to update to version $($latest.Version)"
            Update-Module Microsoft.Graph -ErrorAction Stop
        }
    }

#Connect to Azure AD using the microsoft graph module
Connect-MgGraph -Scopes "User.ReadWrite.All","Application.ReadWrite.All","Domain.Read.All","RoleManagement.ReadWrite.Directory" -ErrorAction Stop

#Gather New User Details
$GivenName = Read-Host "Please enter the users first name"
$Surname = Read-Host "Please enter the users last name"

function New-AADUser {

#generates initial account password
$PasswordProfile = @{}
$PasswordProfile["Password"]= ("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!?.;,._".ToCharArray() | Get-random -Count 50) -join ""
$PasswordProfile["ForceChangePasswordNextSignIn"] = $false

#creates user object - other attributes can be added as required.
#More info on cmdlet here: https://docs.microsoft.com/en-us/powershell/module/microsoft.graph.users/new-mguser?view=graph-powershell-1.0
$user = New-MgUser `
    -GivenName $GivenName `
    -Surname $Surname  `
    -AccountEnabled:$false `
    -DisplayName "$GivenName $Surname" `
    -MailNickname "$GivenName$Surname" `
    -UserPrincipalName "$GivenName.$Surname@forde.co.nz" `
    -PasswordPolicies DisablePasswordExpiration `
    -PasswordProfile $PasswordProfile

#Add group memberships

$DefaultGroups =@("b1e2bdcf-4533-47ae-b9e5-d19d8cf4cf74","73a584ba-b65b-476a-8a28-2b5d6f28484d","71bbf70c-6d72-408d-8c5f-f1a2f491a04b" )

foreach ($group in $DefaultGroups) {
    New-MgGroupMember -GroupId $group -DirectoryObjectId $user.Id
}

#Role addition
# Global reader, Security reader, SharePoint Administrator and Authentication Policy Administrator
$requiredRoleTemplates = Get-MgDirectoryRoleTemplate | Where-Object Id -in "f2ef992c-3afb-46b9-b7cf-a126ee74c451" #, #"5d6b6bb7-de71-4623-b4af-96380a352509","f28a1f50-f6e7-4571-818b-6a12f2af6b6c","0526716b-113d-4c15-b2c8-68e3c22b9f80"
 
# Get all roles
$aadRoles = Get-MgDirectoryRole
 
# Onboard any unused role that we need
$requiredRoleTemplates |
    Where-Object {!($aadRoles | Where-Object RoleTemplateId -eq $_.Id)} |
    ForEach-Object {
        Write-Verbose "Role template '$($_.DisplayName)' with template id $($_.Id) has never been used in this tenant. Onboarding the role before use." -Verbose
        New-MgDirectoryRole -RoleTemplateId $_.Id
    } |
    Out-Null

# Get all roles again
$aadRoles = Get-MgDirectoryRole
 
$requiredRoleTemplates |
    ForEach-Object {
        $aadRoles | Where-Object RoleTemplateId -eq $_.Id
    } | ForEach-Object {
        $Members = Get-MgDirectoryRoleMember -DirectoryRoleId $_.Id
        if($User.Id -notin $Members.Id) {
            Write-Verbose "Adding user to Azure AD role '$($_.DisplayName)'" -Verbose
            New-MgDirectoryRoleMemberByRef -DirectoryRoleId $_.Id -BodyParameter @{"@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($user.Id)"}
        }
    }

}
New-AADUser