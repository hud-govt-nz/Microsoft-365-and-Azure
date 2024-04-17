#===========================================================================================================================================================================#
# Script Name:     Bulk_Add_Alternative_SIPAddress
# Author:          Ashley Forde
# Version:         1.0
# Description:     This script exports all users in the orange users OU then loops through and assigns an alternative proxy address to their ad object. No other services are impacted. 
# 
# Version 1.0 - initial script 30.3.22
#===========================================================================================================================================================================

#Clear
Clear-Host
Write-Host "Bulk_Add_Alternative_SIPAddress"

#Active Directory Module
Import-module ActiveDirectory -DisableNameChecking
$Domain =[string] (Get-ADDomain).Forest
$DC =[string] (Get-ADDomainController -DomainName $Domain -Discover -NextClosestSite).HostName

#Exchange Module
$uri = [string](Get-Exchangeserver).fqdn[1]
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$uri/Powershell/ -Authentication Kerberos 
Import-PSSession -Session $session -AllowClobber -DisableNameChecking


#Variables
$Date = get-Date -format dd.MM.yyyy

#Export List of all users from Orange Users OU
$OU = (Get-OrganizationalUnit -Identity "Orange Users").DistinguishedName
Get-AdUser `
    -Filter * `
    -Properties *
    -SearchBase $OU `
    | Select-Object GivenName, Surname, Name, UserPrincipalName, SamAccountName, EmailAddress `
    | Export-Csv "C:\Temp\Users.csv" -NoTypeInformation

#Raw Export
$file = "C:\Temp\Users.csv"

#Bulk add proxy address to users - Import the Raw Export CSV (assuming you have checked the export and confirmed its correct before re-importing)

Import-CSV $file | ForEach-Object {     
    if ($User = Get-ADUser -Identity $_.SamAccountName -properties GivenName,Surname,SamAccountName -ErrorAction SilentlyContinue) { #Check account is valid
        $SIP = [string] (Get-ADUser -Identity $_.SamAccountName -Properties EmailAddress).EmailAddress #String output users primary email address, this will avoid any issues with preferred naming.
        Set-ADUser $_.SamAccountName -Add @{ProxyAddresses="SIP:$SIP"} #add Alternative SIP address to ProxyAddresses Attribute
        #Set-ADUser $_.SamAccountName -Add @{msRTCSIP-PrimaryUserAddress="SIP:$SIP"} #add Alternative SIP address to ProxyAddresses Attribute

        }
    #catches any names that fail to update and logs it in the SIP_failed.csv below
       else {$failed += $_.User}}
    
$failed | Out-File "C:\temp\SIP_failed_$Date.txt" -Append

"Execution failed for:"
$failed 
