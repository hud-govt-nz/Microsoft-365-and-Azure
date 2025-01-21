<# TEAMS User Adoption - Script to add Teams_User Attribute to accounts
    PURPOSE: The script below will add a custom attribute field to the list of Teams Technical Pilot users.
    
    From On Premise
    PROCESS: 
        - Import list of users via CSV
        - Add a value to Custom Attribute 10 "Teams_User"

    From Azure
    PROCESS:
        - Once the attribute change syncs via AADConnet to Azure then a dynamic group will flag all the applicable users and add their accounts to the relevant Teams Policies. 
#>

#Clear
Clear-Host
Write-Host "TEAMS User Adoption - add Teams_User Attribute to on premise AD user accounts"

#Active Directory Module
Import-module ActiveDirectory -DisableNameChecking -Cmdlet Get-ADDomain,Get-ADDomainController,Get-ADUser,Get-OrganizationalUnit,get-adforest,Add-ADGroupMember,Set-ADUser
$Domain =[string] (Get-ADDomain).Forest
$DC =[string] (Get-ADDomainController -DomainName $Domain -Discover -NextClosestSite).HostName
$OU = (Get-OrganizationalUnit -Identity "Orange Users").DistinguishedName
$UPNSuffix = [string] (get-adforest).UPNSuffixes[0]

#Secondary Functions
function Get-FileName ($InitialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv) | *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $openFileDialog.Filename}
function Save-File([string] $initialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.initialDirectory = $initialDirectory
    $SaveFileDialog.filter = "TXT (*.txt) | *.txt"
    $SaveFileDialog.ShowDialog() |  Out-Null
    return $SaveFileDialog.Filename}    
function Get-SAMFromName ($NameToConvert) {
    return (get-aduser -filter {name -like $NameToConvert}).SamAccountName}

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

    #Sets attribute against account
    #Set-ADUser -Identity $User -Server $DC -add @{extensionAttribute10 = "Teams_User"}

    #Remove attribute against account
    Set-ADUser -Identity $User -Server $DC -clear "extensionAttribute10" 

    }


#Captures any accounts that are missed.
Write-Host 'Script failed for the following users:'
$Failed 

