#Sign into Exchange Online
Set-ExecutionPolicy Unrestricted -Scope CurrentUser
$usrname = Read-host 'Enter sign in name'
Connect-ExchangeOnline -UserPrincipalName $usrname  -ShowProgress $true

#confirm mailbox exists if not create new
remove-Mailbox res-surfacehub01@develop.mfat.govt.nz
Get-Mailbox RES-MEETRM01@acceptance.mfat.govt.nz

#Create new meeting room mailbox
New-Mailbox `
-MicrosoftOnlineServicesID 'res-surfacehub01@develop.mfat.govt.nz' `
-Alias res-surfacehub01 `
-Name "..WLN HSBC-L16 Surface Hub (Teams)" `
-Room -EnableRoomMailboxAccount $true `
-RoomMailboxPassword (ConvertTo-SecureString -String W@rning10 -AsPlainText -Force)
New-Mailbox `
-MicrosoftOnlineServicesID 'RES-MEETRM01@acceptance.mfat.govt.nz' `
-Alias res-meetrm01 `
-Name "..WLN Meeting Room 01 (Teams)" `
-Room -EnableRoomMailboxAccount $true `
-RoomMailboxPassword (ConvertTo-SecureString -String W@rning10 -AsPlainText -Force)

#Manually Set Mailbox Type
Set-Mailbox -Identity res-devteamshub -Type room

#Configure the meeting 
Set-Mailbox -Identity res-devteamshub@develop.mfat.govt.nz -EnableRoomMailboxAccount $true -RoomMailboxPassword (ConvertTo-SecureString -String W@rning10 -AsPlainText -Force)
Set-CalendarProcessing -Identity res-devteamshub -AutomateProcessing AutoAccept -AddOrganizerToSubject $false -DeleteComments $false -DeleteSubject $false -RemovePrivateProperty $false -ProcessExternalMeetingMessages $true
Set-CalendarProcessing -Identity res-devteamshub -AddAdditionalResponse $true -AdditionalResponse "Hi this is a Teams meeting room."
Set-MailboxCalendarConfiguration -Identity res-devteamshub -AddOnlineMeetingToAllEvents $true

Set-Mailbox -Identity RES-MEETRM01@acceptance.mfat.govt.nz -EnableRoomMailboxAccount $true -RoomMailboxPassword (ConvertTo-SecureString -String W@rning10 -AsPlainText -Force)
Set-CalendarProcessing -Identity RES-MEETRM01@acceptance.mfat.govt.nz -AutomateProcessing AutoAccept -AddOrganizerToSubject $false -DeleteComments $false -DeleteSubject $false -RemovePrivateProperty $false -ProcessExternalMeetingMessages $true
Set-CalendarProcessing -Identity RES-MEETRM01@acceptance.mfat.govt.nz -AddAdditionalResponse $true -AdditionalResponse "Hi this is a Teams meeting room."
Set-MailboxCalendarConfiguration -Identity RES-MEETRM01@acceptance.mfat.govt.nz -AddOnlineMeetingToAllEvents $true

#Set the room's password not to expire and apply license
Connect-MsolService -Credential $cred
Set-MsolUser -UserPrincipalName res-devteamshub -PasswordNeverExpires $true -UsageLocation "NZ"
Set-MsolUser -UserPrincipalName RES-MEETRM01@acceptance.mfat.govt.nz -PasswordNeverExpires $true -UsageLocation "NZ"

#Get-MsolAccountSku
Set-MsolUserLicense -UserPrincipalName res-surfacehub01@develop.mfat.govt.nz -AddLicenses $strLicense
Set-MsolUserLicense -UserPrincipalName RES-MEETRM01@acceptance.mfat.govt.nz -AddLicenses $strLicense

#Setup Teams Room List
#Set office location
Set-Mailbox -Identity res-devteamshub -Office "WLN HSBC - Level 16"

#Create teams distribution group (room list)
New-DistributionGroup -Name "WLN HSBC Meeting Rooms" –PrimarySmtpAddress "wln-hsbc-meetrms@develop.mfat.govt.nz" –RoomList
New-DistributionGroup -Name "...WLN Meeting Rooms" –PrimarySmtpAddress "wln-meetrms@acceptance.mfat.govt.nz" –RoomList

#add members to room list based on location/office
Add-DistributionGroupMember -Identity "WLN HSBC Meeting Rooms" -Member res-devteamshub@develop.mfat.govt.nz
Add-DistributionGroupMember -Identity "Wellington Meeting Rooms" -Member res-meetrm19.02
Add-DistributionGroupMember -Identity "...WLN Meeting Rooms" -Member RES-MEETRM01


# BULK: add all meeting rooms in the Wellington office
$office = "Wellington - Level 19"
$resmeetrm = Get-Mailbox -RecipientTypeDetails RoomMailbox -Filter {Office -eq '$office'} | select -ExpandProperty Alias
$resmeetrm | Add-DistributionGroupMember -Identity "Wellington Meeting Rooms"


Connect-MicrosoftTeams
Get-CsTeamsMeetingPolicy

Get-CsUserPolicyAssignment -Identity cf069567-7b33-4b99-b2f1-442b8470d553 #MFAT-BookingLaptop01
Get-CsUserPolicyAssignment -Identity 6f32ac4c-f90d-418e-9523-542fffca0c6b #MFAT-BookingLaptop02
Get-CsUserPolicyAssignment -Identity 3fa84d18-baf1-4ef3-bb87-c955198309b0 #WLN Meeting Room 19.12
Get-CsUserPolicyAssignment -Identity 15804af0-839e-44cb-8efe-ea7efb4b6935 #WLN Meeting Room Hub

Get-CsGroupPolicyAssignment -PolicyType TeamsMeetingPolicy
Get-CsGroupPolicyAssignment -PolicyType TeamsMessagingPolicy

Get-CsGroupPolicyAssignment -GroupId 20ce916f-71d1-4189-9115-e8d0043e9a37 #TeamsMessagingPolicy

Set-CsUserPolicyAssignment -Identity cf069567-7b33-4b99-b2f1-442b8470d553 -PolicyType TeamsMeetingPolicy
