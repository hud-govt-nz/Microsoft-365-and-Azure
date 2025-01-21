Connect-ExchangeOnline

#1 align transport rule 5 in prod "Client Rules To External Block"

New-TransportRule -Name 'Deny' -Comments '' -Mode Enforce -ExceptIfFromMemberOf 'AAD-Exchange-AllowSendEMail@mfatprod.onmicrosoft.com' -DeleteMessage $true
New-TransportRule -Name 'Block email to OnMicrosoft Domain' -Comments '' -Mode Enforce -RecipientAddressContainsWords '*.onmicrosoft.com' -ExceptIfSenderIpRanges 20.36.42.148 -DeleteMessage $true
New-TransportRule -Name 'Limit message flow to Calendar items only' -Comments '' -Mode Enforce -FromMemberOf 'MVPPilotUsers@mfatprod.onmicrosoft.com' -ExceptIfMessageTypeMatches Calendaring -DeleteMessage $true
New-TransportRule -Name 'Auto Accept Meeting Invites' -Comments '' -Mode Enforce -SentToMemberOf 'AAD-Exchange-AutoAcceptInvites@mfatprod.onmicrosoft.com' -FromScope InOrganization -SetHeaderName 'X-MS-Exchange-Organization-CalendarBooking-Response' -SetHeaderValue 'Accept'
New-TransportRule -Name 'Client Rules To External Block' -Comments 'CIS Benchmark Requirement 4.9' -Mode Enforce -SentToScope NotInOrganization -MessageTypeMatches AutoForward -FromScope InOrganization -RejectMessageReasonText "To improve security, auto-forwarding rules to external addresses have been disabled. Please contact your Microsoft Partner if you'd like to set up an exception." -RejectMessageEnhancedStatusCode '5.7.1'


# Fix up AntiPhysiPolicies
Get-AntiPhishPolicy | Format-Table Name,Enabled,IsDefault
Get-AntiPhishRule | Format-Table Name,Enabled,IsDefault

Set-AntiPhishPolicy -Identity "Strict Preset Security Policy1620941915492" -Enabled $false
Set-AntiPhishPolicy -Identity "MFAT Custom Strict Policy" -Enabled $true

Remove-AntiPhishPolicy -Identity "Strict Preset Security Policy1620941915492" -Force