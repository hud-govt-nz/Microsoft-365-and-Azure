#===========================================================================================================================================================================#
# Script Name:     Onboard Cloud Only Accounts
# Author:          Ashley Forde
# Version:         1.0
# Description:     This script is for bulk creating user accounts in Azure AD. 
# Notes:           Needs to be run from a box that has the Azure AD Module and access to the Internet
#
# Version 1.0 - inital script 2/3/22
#===========================================================================================================================================================================

#Module
Clear-Host
Write-Host "Connecting to Azure AD"

#Connect to Azure AD
Connect-AzureAD

#Core Functions
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
function Get-RandomPassword {
    param (
        [Parameter(Mandatory)]
        [ValidateRange(4,[int]::MaxValue)]
        [int] $length,
        [int] $upper = 1,
        [int] $lower = 1,
        [int] $numeric = 1,
        [int] $special = 1
        )
    if($upper + $lower + $numeric + $special -gt $length) {
        throw "number of upper/lower/numeric/special char must be lower or equal to length"
        }

    $uCharSet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    $lCharSet = "abcdefghijklmnopqrstuvwxyz"
    $nCharSet = "0123456789"
    $sCharSet = "#!?=@"
    $charSet = ""

    if($upper -gt 0) { $charSet += $uCharSet }
    if($lower -gt 0) { $charSet += $lCharSet }
    if($numeric -gt 0) { $charSet += $nCharSet }
    if($special -gt 0) { $charSet += $sCharSet }

    $charSet = $charSet.ToCharArray()
    $rng = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
    $bytes = New-Object byte[]($length)
    $rng.GetBytes($bytes)

    $result = New-Object char[]($length)
    for ($i = 0 ; $i -lt $length ; $i++) {
        $result[$i] = $charSet[$bytes[$i] % $charSet.Length]
        }

    $password = (-join $result)
    $valid = $true

    if($upper   -gt ($password.ToCharArray() | Where-Object {$_ -cin $uCharSet.ToCharArray() }).Count) { $valid = $false }
    if($lower   -gt ($password.ToCharArray() | Where-Object {$_ -cin $lCharSet.ToCharArray() }).Count) { $valid = $false }
    if($numeric -gt ($password.ToCharArray() | Where-Object {$_ -cin $nCharSet.ToCharArray() }).Count) { $valid = $false }
    if($special -gt ($password.ToCharArray() | Where-Object {$_ -cin $sCharSet.ToCharArray() }).Count) { $valid = $false }

    if(!$valid) {   
        $password = Get-RandomPassword $length $upper $lower $numeric $special
        }
    return $password
    }
function Transpose-Data {
    Param(
        [Parameter(Mandatory=$True)]
        [string[]]$ArrayNames,
        [switch]$NoWarnings = $False
        )
    $ValidArrays,$ItemCounts = @(),@()
    $VariableLookup = @{}
        ForEach ($Array in $ArrayNames) {
        Try {
            $VariableData = Get-Variable -Name $Array -ErrorAction Stop
            $VariableLookup[$Array] = $VariableData.Value
            $ValidArrays += $Array
            $ItemCounts += ($VariableData.Value | Measure).Count
            }
        Catch {
            If (!$NoWarnings) {Write-Warning -Message "No variable found for [$Array]"}
            }
        }
    $MaxItemCount = ($ItemCounts | Measure -Maximum).Maximum
    $FinalArray = @()
        For ($Inc = 0; $Inc -lt $MaxItemCount; $Inc++) {
            $FinalObj = New-Object PsObject
                ForEach ($Item in $ValidArrays) {
                    $FinalObj | Add-Member -MemberType NoteProperty -Name $Item -Value $VariableLookup[$Item][$Inc]
                    }
            $FinalArray += $FinalObj
            }
        $FinalArray
    }  

#Import CSV, Set Results Array
$Import = Get-FileName
$Users = Import-CSV $Import
$Result = @()

#Account Generation
Function Create-AzureNewUser ($UserDetails) {

#Fixed Variables - Set the UPN suffix depending on the environment you are using this script.
#DEV
#$UPNSuffix = 'develop.mfat.govt.nz'
#ACPT
#$UPNSuffix = 'acceptance.mfat.govt.nz'
#PROD
$UPNSuffix = 'mfat.govt.nz'

$Company = "MFAT"

#Output Arrays
$Username = @()
$Password = @()
$Duplicate = @()

#Variables from CSV
$Validation = @()
$FirstName = $UserDetails.FirstName
$LastName = $UserDetails.LastName
$Department = $UserDetails.Department
$DepartmentAcronym = $UserDetails.DepartmentAcronym
$JobTitle = $UserDetails.JobTitle
$Office = $UserDetails.Office
$StreetAddress = $UserDetails.StreetAddress
$City = $UserDetails.City
$PostalCode = $UserDetails.PostalCode
$Country = $UserDetails.Country
$MobilePhone = $UserDetails.MobilePhone

#Name Formatting     
$Name = "$FirstName $LastName"

#Include row 154 for standard user accounts, For Admin accounts # out.
$DisplayName = ($LastName.ToUpper())+", $FirstName ($DepartmentAcronym)"
$DisplayName = "$FirstName $LastName ($DepartmentAcronym)"

    #Truncates spaces in first name
    if($FirstName.Contains(' ')) {$FirstName = $FirstName.Replace(' ','')}
    #Truncates spaces in last name
    if($LastName.Contains(' ')) {$LastName = $LastName.Replace(' ','')}

$FirstDotLast = "$FirstName.$LastName"
$UPN = "$FirstDotLast@$UPNSuffix"
    
    If (Get-AzureAdUser -Filter "userPrincipalName eq '$UPN'") {   
        Write-host "A user account named ($UPN) already exists in Azure AD." -ForegroundColor Yellow
        $Duplicate += "$UPN already exists, please check"} 
        Else {
            Write-Host "Creating user $Name" # User does not exist then proceed to create the new user account          
                    
            Add-Type -AssemblyName System.Web
            $CreatePwd = Get-RandomPassword 12 2 7 2 1
            $PasswordProfile=New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
            $PasswordProfile.Password=$CreatePwd
                    
            New-AzureADUser `
                -DisplayName $DisplayName `
                -GivenName $FirstName `
                -Surname $LastName `
                -UserPrincipalName $UPN `
                -UsageLocation 'NZ' `
                -MailNickName $FirstDotLast `
                -PasswordProfile $PasswordProfile `
                -AccountEnabled $false `
                -JobTitle $JobTitle `
                -CompanyName $Company `
                -Country $Country `
                -City $City `
                -Department $Department `
                -physicalDeliveryOfficeName $Office `
                -PostalCode $PostalCode `
                -Mobile $MobilePhone 
  
    If (Get-AzureAdUser -Filter "UserPrincipalName eq'$UPN'") {
        $String = $StreetAddress
        $StAddress = Out-String -InputObject $String.Split(";") -Width 100
        Set-AzureADUser -ObjectID $UPN -StreetAddress $StAddress} 
        
        Write-Host "Account $UPN has been created" -ForegroundColor Green
        }

    $Username += $UPN
    $Password += $CreatePwd
    $Validation += Transpose-Data -ArrayNames "Username","Password","Duplicate" | ft -AutoSize -Wrap
    return $Validation
    }
  
#Create Multiple User Accounts
    ForEach ($User in $Users){
        If($User.StreetAddress -eq "#N/A") {
            Continue
            }
            Else {
                $Result += Create-AzureNewUser ($User)
                }
        }   

#Display Output on Console Window
Write-Host "Usernames and Passwords are listed below, if a user is a duplicate please check the user name and/or create manually" -ForegroundColor Yellow
$Result
Write-Host "Please wait while save window appears" -ForegroundColor Yellow
start-sleep 5

#Save TXT Output      
$ResultTXT = Save-File
Write-Output "--------------------------------------------------------------------------------------------------------------------"| Out-File $ResultTXT
Write-Output "The following Users have been onboarded into Azure" | Out-File $ResultTXT -Append
Write-Output "--------------------------------------------------------------------------------------------------------------------"| Out-File $ResultTXT -Append
Write-Output "Usernames and Passwords are listed below, if a user is a duplicate please check the user name and/or create manually"| Out-File $ResultTXT -Append
Write-Output "--------------------------------------------------------------------------------------------------------------------"| Out-File $ResultTXT -Append
$result | Out-File $ResultTXT -Append
start-sleep 2
Invoke-Item $ResultTXT   
