<#
.NAME: Azure AD User Export 
.DESCRIPTION: Custom User Export Script which outputs the current state of users in Azure AD. This has been compiled as part of HUDs ITGC Audit. This script leverages the PowerShell Graph SDK/Cmdlets.
.AUTHOR: Ashley Forde
.DATE: 12 June 2023
#>

# Connect to Azure via Microsoft Graph SDK
Connect-MgGraph -Scopes "User.Read.All","Directory.Read.All","User.ReadBasic.All","Reports.Read.All","AuditLog.Read.All","Organization.Read.All"

# Select Graph Beta Profile
Select-MgProfile -Name beta

# Value Arrays
$1 = @('id', 'SignInActivity', 'lastPasswordChangeDateTime')
$2 = @( 'id',
        'GivenName',
        'Surname',
        'DisplayName',
        'AccountEnabled',
        'UserType',
        'UserPrincipalName',
        'JobTitle',
        @{Name='CreatedDate';Expression={([datetime]$_.CreatedDateTime).ToString("dd/MMM/yyyy")}},
         @{Name='AHOStartDate';Expression={$_.AdditionalProperties['extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserStartDate']}},
         @{Name='EmployeeCategory';Expression={$_.AdditionalProperties['extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmploymentCategory']}},
         @{Name='M365E3';Expression={if ($_.assignedLicenses.skuid -eq "05e9a617-0261-4cee-bb44-138d3ef5d965"){$true}else{$false}}},
         @{Name='M365E5';Expression={if ($_.assignedLicenses.skuid -eq "06ebc4ee-1bb5-47dd-8120-11324bc54e06"){$true}else{$false}}},
         @{Name='NoLicense';Expression={($_.assignedLicenses.count -eq 0)}},
         @{Name='RoomMailbox';Expression={$_.AdditionalProperties['extension_56a473fa1d5b476484f306f7b06ee688_RoomMailbox']}},
         @{Name='SharedMailbox';Expression={$_.AdditionalProperties['extension_56a473fa1d5b476484f306f7b06ee688_SharedMailbox']}}
         )

# Array 1: Obtain user values not natively selected under the Graph SDK and place in an array object.
$1s =@()
$1s += Get-MgUser -All -Property $1 | Select-Object -Property id,
            @{Name='LastSignInDateTime';Expression={([datetime]$_.SignInActivity.LastSignInDateTime).ToString("dd/MMM/yyyy")}},
            @{Name='LastPasswordChange'; Expression={([datetime]$_.lastPasswordChangeDateTime).ToString("dd/MMM/yyyy")}}

# Array 2: Obtain general user values as well as custom attributes
$2s =@()
$2s += Get-MgUser -All | Select-Object $2

# Merge the arrays
for ($i = 0; $i -lt $1s.Count; $i++) {
    $1 = $1s[$i]
    $2 = $2s[$i]
            
    $1 | Add-Member -MemberType NoteProperty -Name 'GivenName' -Value $2.GivenName
    $1 | Add-Member -MemberType NoteProperty -Name 'Surname' -Value $2.Surname
    $1 | Add-Member -MemberType NoteProperty -Name 'DisplayName' -Value $2.DisplayName
    $1 | Add-Member -MemberType NoteProperty -Name 'CreatedDate' -Value $2.CreatedDate
    $1 | Add-Member -MemberType NoteProperty -Name 'AccountEnabled' -Value $2.AccountEnabled
    $1 | Add-Member -MemberType NoteProperty -Name 'UserType' -Value $2.UserType
    $1 | Add-Member -MemberType NoteProperty -Name 'AHOStartDate' -Value $2.AHOStartDate
    $1 | Add-Member -MemberType NoteProperty -Name 'EmployeeCategory' -Value $2.EmployeeCategory
    $1 | Add-Member -MemberType NoteProperty -Name 'UserPrincipalName' -Value $2.UserPrincipalName
    $1 | Add-Member -MemberType NoteProperty -Name 'JobTitle' -Value $2.JobTitle
    $1 | Add-Member -MemberType NoteProperty -Name 'M365E3' -Value $2.M365E3
    $1 | Add-Member -MemberType NoteProperty -Name 'M365E5' -Value $2.M365E5
    $1 | Add-Member -MemberType NoteProperty -Name 'NoLicense' -Value $2.NoLicense
    $1 | Add-Member -MemberType NoteProperty -Name 'RoomMailbox' -Value $2.RoomMailbox
    $1 | Add-Member -MemberType NoteProperty -Name 'SharedMailbox' -Value $2.SharedMailbox
            
}

# Output the results to the console
$Folder = "$($env:homedrive)\HUD\06_Reporting"
            
if(Test-Path -Path $Folder) {
    "06_Reporting Folder exists..."
} else {
    New-Item -Path C:\HUD\ -Name 06_Reporting -ItemType Directory -Force -Confirm:$false
}
            
$Date = Get-Date -f yyyyMMddhhmm
$FileName = "AuditNZ_AADUserReport_$Date.CSV"
$1s | Export-CSV "$Folder\$FileName" -NoTypeInformation -Encoding UTF8
            
Write-Host "The report $FileName has been saved in C:\HUD\06_Reports\" -ForegroundColor Green