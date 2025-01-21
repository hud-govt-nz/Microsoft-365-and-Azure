<# PRODUCTION: Prepare Technical Pilot Teams users
    .PURPOSE:
        The process below will be implemented in on premise AD to prepare AD user accounts for Teams Adoption. 
        
        .This includes setting a Proxy SIP Address of *@mfat.govt.nz on each users AD user object so that MS Teams recognises the 
         account as being distinctly different from Skype for business on premise. 

        .A value of "Teams_User" will be added to ExtensionAttribute10 on each users AD user object. This flag will be used to dynamically
         add users to the AAD-Teams-Restricted-User security group in AAD. This group links to MS teams to apply all the correct application,
         service policies within teams policies configuration spreadsheet:
         http://o-wln-gdm/Functions/InformationManagement/EnterpriseArchitecture/ArchitecturePublic/Cloud/Restricted/Teams/Teams%20Policies.xlsx?web=1

        .A second script will be used in conjunction with this one to add all the AD users to the appropiate licence and security groups within Azure AD and
         will be run from the production EOE box that allows access to Azure Powershell. 

    .PROCESS:
        1. Import a set list of usernames for the pilot group via .csv
        2. Generate a SIP address value by copying the users Primary SMTP address (email). - RFC874
        3. Apply an additional proxyAddress value "SIP:<email address>" to each account in the list. - RFC874
        4. Apply the value "Teams_User" to the extensionAttribute10 field within each users AD object. - RFC874
        5. Force AD replication across OP1/OP2/OP4
        6. Force AAD Sync to ******.onmicrosoft.com (Prod Azure)

    .Created by: Ashley Forde
    .Version: 1
    .Created date: 7.4.22
#>

#Import Module
#Clear
Clear-Host
Write-Host "Prepare Technical Pilot Teams users"

#Active Directory Module
Import-module ActiveDirectory -DisableNameChecking -Cmdlet Get-ADDomain,Get-ADDomainController,Get-ADUser,Get-OrganizationalUnit,get-adforest,Add-ADGroupMember,Set-ADUser
$Domain =[string] (Get-ADDomain).Forest
$DC =[string] (Get-ADDomainController -DomainName $Domain -Discover -NextClosestSite).HostName
$OU = (Get-OrganizationalUnit -Identity "Orange Users").DistinguishedName
$DisabledUsers = (Get-OrganizationalUnit -Identity "Disabled Users").DistinguishedName
$UPNSuffix = [string] (get-adforest).UPNSuffixes[0]

#Functions
function Get-FileName ($InitialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv) | *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $openFileDialog.Filename
}
function Save-File([string] $initialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.initialDirectory = $initialDirectory
    $SaveFileDialog.filter = "TXT (*.txt) | *.txt"
    $SaveFileDialog.ShowDialog() |  Out-Null
    return $SaveFileDialog.Filename
}    
function Get-SAMFromName ($NameToConvert) {
    return (get-aduser -filter {name -like $NameToConvert}).SamAccountName
}

#Export List of Orange Users
Write-Host "Exporting List of Orange Users for validation. . ." -ForegroundColor Yellow
$tempFolder = "C:\Temp\"
#Orange Users
Get-AdUser -Filter * -Properties * -SearchBase $OU `
    | Select-Object DisplayName, GivenName, Surname, Name, UserPrincipalName, SamAccountName, EmailAddress, adminDescription, enabled `
    | Export-Csv "$tempFolder\Orange_Users.csv" -NoTypeInformation
#Disabled Users
Get-AdUser -Filter * -Properties * -SearchBase $DisabledUsers `
    | Select-Object DisplayName, GivenName, Surname, Name, UserPrincipalName, SamAccountName, EmailAddress, adminDescription, enabled `
    | Export-Csv "$tempFolder\Orange_Users.csv" -NoTypeInformation -Append

Write-Host "Exporting results. . ." -ForegroundColor Yellow
explorer.exe $tempFolder

#pause script until csv is checked and edited
Write-Host "Please validate the list before continuing, once completed continue." -ForegroundColor Yellow

#Resume
read-host “Press ENTER to continue...”

#Import CSV, Set Results Array
$Import = Get-FileName
$Users = Import-CSV $Import
$Failed = @()

#Loop through and apply attribute to each user account
$Users | ForEach-Object {
    
    #Checks name if given as SAMAccountName or as Full Name, resolves to SamAccountName
    if ($UserName = Get-SAMFromName($_.Name)) {}
    Else {$UserName = $_.Name}
    
    #Grabs user as variable
    $User = get-aduser $UserName -Properties *
    $UPN = (Get-ADUser $UserName -Properties UserPrincipalName).UserPrincipalName 
    #alternative Variable
    #$Email = (Get-ADUser $UserName -Properties EmailAddress).EmailAddress

    #Sets attribute against account
    Set-ADUser -Identity $User -Server $DC -Add @{ProxyAddresses="SIP:$UPN"} #Alternative - Set-ADUser -Identity $User -Server $DC -Add @{ProxyAddresses="SIP:$UPN"}
    Set-ADUser -Identity $User -Server $DC -add @{extensionAttribute10 = "Teams_User"}

    #Remove attribute against account
    #Set-ADUser -Identity $User -Server $DC -remove @{ProxyAddresses="SIP:$Email"} 
    #Set-ADUser -Identity $User -Server $DC -clear "extensionAttribute10"     

}
#Captures any accounts that are missed.
Write-Host 'Script failed for the following users:'
$Failed 
