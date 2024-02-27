<#
History.
- V1 created by Malcolm Jeffrey @ Fujitsu Consulting.
- V2 Updates by Ashley Forde @HUD. 
    - Updated graph connection scopes and profile set to Beta to give cmds the ability to view user additional properties which sync from Aho
    - Changed Get-MGUser cmd to amend filtering to leverage employee categories sync'd from Aho
#>

#Connect to Microsoft Graph
. .\GraphLogin.ps1

# Collect User Information
$Values = Get-MgBetaUser -All | Select-Object GivenName,Surname,JobTitle,MobilePhone,UserPrincipalName,accountEnabled,userType,Officelocation,CompanyName,AdditionalProperties

# Gather up all fields and filter
$FilteredValues = $Values | Where-Object {
	($_.CompanyName -eq 'Ministry of Housing and Urban Development') `
 		-and ($_.UserType -eq 'Member') `
 		-and ($_.OfficeLocation -like 'Wellington*') `
 		-and ($_.GivenName -ne 'Mobile') `
 		-and ($_.GivenName -ne 'Teleconference') `
 		-and ($_.Jobtitle -ne 'Oracle Support Consultant') `
 		-and ($_.Jobtitle -ne 'Lead Oracle Support Consultant') `
 		-and ($_.AdditionalProperties -ne $null -and $_.AdditionalProperties['extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmploymentCategory'] -ne $null) `
 		-and ($_.AdditionalProperties -ne $null -and $_.AdditionalProperties['extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmploymentCategory'] -ne 'ORA_HRX_CONSULTANT') `
 		-and ($_.AdditionalProperties -ne $null -and $_.AdditionalProperties['extension_56a473fa1d5b476484f306f7b06ee688_RoomMailbox'] -ne 'Room Mailbox') `
 		-and ($_.AdditionalProperties -ne $null -and $_.AdditionalProperties['extension_56a473fa1d5b476484f306f7b06ee688_SharedMailbox'] -ne 'Shared Mailbox')

} | Select-Object GivenName,Surname,Jobtitle,MobilePhone,UserPrincipalName,@{ n = 'Status'; e = { if ($_.accountenabled -eq $true) { "1" } else { "0" } } }

# Generate Filename
$filename = (Get-Date -Format "dd MMMM yy").ToString() + ' - SFTP Upload.csv'

# Export Filtered Results to CSV
$FilteredValues | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 | Set-Content -Path ".\$filename"

#Set variables for use during SFTP upload
$Dir = "./"
$SFTPath = "/staff_sync_files"

#Set SFTP server credentials
$passwordTest = "99a$#4af!@2as@"
$user = "hud_staff"
$securePasswordTest = ConvertTo-SecureString $passwordTest -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential ($user,$securePasswordTest)

#Connect to SFTP Session
New-SFTPSession -ComputerName 54.252.245.176 -Credential $credential -AcceptKey

#Create new SFTP file from gathered information
Set-SFTPItem -Destination $SFTPath -Path $dir\$filename -SessionId 0 -Force

#Show list of files to confirm upload. Result only visible during testing
Get-SFTPChildItem -SessionId 0 -File -Path /staff_sync_files
