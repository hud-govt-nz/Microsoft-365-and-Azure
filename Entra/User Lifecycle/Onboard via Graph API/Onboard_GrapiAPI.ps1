<#
.NAME: Onboard User Script
.PURPOSE: This script is intended to onboard users into Azure AD using the Microsoft Graph API native commands.
.DATE: 17 June 2022
.AUTHOR: Ashley Forde

.NOTE: This script leverages the MS Graph API via the invoke-webRequest cmd which is native to PowerShell. This removes the requirement to 
       install the MS Graph module so that it can be run from any machine/automation account.

.PROCESS: 
    - Load functions, Get-JSONFile: simple filedialog used to facilitate New-AADUser function. 
    - Supply Credentials via JSON payload
    - Obtain AuthToken 
    - Supply JSON payload for account creation
    - Create user via Invoke-WebRequest call
        -User set is segmented from JSON payload so that it can be isolated and set into its own body for the account creation command.
    - Assign Manager via Invoke-WebRequest call
        -ManagerData set is segmented from JSON payload so that it can be isolated and set into its own body for the assign manager command.
    -Results shown - note the Assign Manager command returns no result if successful. 

#>

#Open File Dialog
function Get-JSONFile {
    [cmdletBinding()]
    param(
        [Parameter()]
        [ValidateScript({Test-Path $_})]
        [String]
        $InitialDirectory
    )

    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog 
    if($InitialDirectory){
    $FileBrowser.InitialDirectory = $InitialDirectory
    }
    else{
    $fileBrowser.InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
    }   
    $FileBrowser.Filter = 'JSON (*.json)|*.json|All Files (*.*)|*.*'

    [void]$FileBrowser.ShowDialog()
    $FileBrowser.FileName
}

#Auth file import (JSON) - future state will be to incorporate this with LastPass over Powershell.
$File = Get-JSONFile "C:\Users\Ashley.Forde\OneDrive - Ministry of Housing and Urban Development\Documents\_Code\Microsoft-365-and-Azure\Entra\Auth\Provision New Starters App (Aho).json"
#$File = "C:\Support\Code_Repository\.New_Code\.Public Cloud\.Capability\AAD\M365_DSC_Application_Auth.JSON"
$Content = Get-Content $File | ConvertFrom-Json

#Credentials as a string
$ApplicationId = $Content.ApplicationId
$TenantId = $Content.TenantId
$ClientSecret = $Content.Secret

#Mandatory URI
$Uri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
$Scope = "https://graph.microsoft.com/.default"

#BodyParameters - for Invoke Web Request
$RequestBody = @{client_id = $ApplicationId;client_Secret = $ClientSecret;grant_type = "client_credentials";scope = $Scope}

#Method & ContentType
$Method1 = "POST"
$ContentType = "application/json"

#Auth Request via Web Request
$OAuthResponse = Invoke-RestMethod -Method $Method1 -Uri $Uri -Body $RequestBody
$AccessToken = $OAuthResponse.access_token
    
#Graph User URL
$UserURL = "https://graph.microsoft.com/v1.0/users"

#New user details import (JSON)
$Import = Get-JSONFile C:\
$UserJSON = Get-Content $Import | ConvertFrom-Json

#User Field as String
$Userbody =@()
$UserBody = $UserJSON.User | ConvertTo-Json
    
#User Creation Web Request
$Headers = @{'Content-Type'="application\json";'Authorization'="Bearer $AccessToken"}
$Usr = Invoke-RestMethod -Uri $UserURL -Method $Method1 -Headers $Headers -ContentType $ContentType -Body $UserBody -UseBasicParsing

#Assign Manager
$MgrBodyUpdate = @{
    "@odata.id"= "$($UserURL)/$($UserJSON.ManagerData.Manager)"
}

$Method2 = "PUT"
$ref = [string]'$ref'
$BodyUpdate = $MgrBodyUpdate | ConvertTo-Json
$URI2 = [string]"$($UserURL)/$($UserJSON.User.userPrincipalName)/manager/$ref"

$Mgr = Invoke-RestMethod -Uri $URI2 -Headers $Headers -ContentType $ContentType -Method $Method2 -Body $BodyUpdate -UseBasicParsing

#Assign Custom Attributes
$Attributes =@{
    "extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmploymentCategory" = "$($UserJSON.CustomAttributes.EmploymentCategory)"
    "extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserStartDate"  = "$($UserJSON.CustomAttributes.UserStartDate)"
    "extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmployeeType" = "$($UserJSON.CustomAttributes.EmploymentType)"
    "extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserOrganisationalGroup" = "$($UserJSON.CustomAttributes.Group)"
}

$Method3 = "PATCH"
$AttributesUpdate = $Attributes | ConvertTo-Json
$URI3 = [string]"$($UserURL)/$($UserJSON.User.userPrincipalName)"

$Attributes = Invoke-RestMethod -Uri $URI3 -Headers $Headers -ContentType $ContentType -Method $Method3 -Body $AttributesUpdate -UseBasicParsing

#Results
$Usr, $Mgr, $Attributes
