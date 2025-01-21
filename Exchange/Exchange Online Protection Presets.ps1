Connect-ExchangeOnline

Get-HostedContentFilterPolicy -Identity * | ft
Get-HostedContentFilterRule -Identity * |ft

#Anti-Spam Policies
#Inbound - MFAT Custom Strict Policy (Inbound)
Get-HostedContentFilterPolicy -Identity * | ft

#ensure the rule is enabled.
Get-HostedContentFilterRule -Identity * |ft

#Set custom attributes to policy.
#DOESNT WORK IN PROD, ERROR
Set-HostedContentFilterPolicy -Identity "MFAT Custom Strict Policy (Inbound)" -BulkThreshold 4 -MarkAsSpamBulkMail on -EnableLanguageBlockList $false -LanguageBlockList $null -EnableRegionBlockList $false -RegionBlockList $null -SpamAction Quarantine -HighConfidenceSpamAction Quarantine -PhishSpamAction quarantine -HighConfidencePhishAction quarantine -BulkSpamAction quarantine -QuarantineRetentionPeriod 30 -InlineSafetyTipsEnabled $true -PhishZapEnabled $true -SpamZapEnabled $true -EnableEndUserSpamNotifications $true -EndUserSpamNotificationFrequency 1

#Outbound
Get-HostedOutboundSpamFilterPolicy -Identity * | ft

#Create policy if required
New-HostedOutboundSpamFilterPolicy -Name "MFAT Custom Strict Policy (Outbound)" -RecipientLimitExternalPerHour 400 -RecipientLimitInternalPerHour 800 -RecipientLimitPerDay 800 -ActionWhenThresholdReached blockuser -AutoForwardingMode automatic -BccSuspiciousOutboundMail $false -NotifyOutboundSpam $false

# or amend policy if it exists already.
Set-HostedOutboundSpamFilterPolicy -Identity "MFAT Custom Strict Policy (Outbound)" -RecipientLimitExternalPerHour 400 -RecipientLimitInternalPerHour 800 -RecipientLimitPerDay 800 -ActionWhenThresholdReached blockuser -AutoForwardingMode automatic -BccSuspiciousOutboundMail $false -NotifyOutboundSpam $false

# Anti-Malware Policies
Get-MalwareFilterPolicy -Identity * | ft
Get-MalwareFilterRule -Identity * | ft

Get-MalwareFilterPolicy -Identity "MFAT Custom Strict Policy" | fl

#Set policy to match DD
Set-MalwareFilterPolicy -Identity "MFAT Custom Strict Policy" -EnableFileFilter $true -ZapEnabled $true -Action deletemessage -EnableInternalSenderNotifications $true -EnableExternalSenderNotifications $false -EnableInternalSenderAdminNotifications $true -InternalSenderAdminAddress "imdcloudopsalerts@mfat.govt.nz" -EnableExternalSenderAdminNotifications $false -CustomNotifications $false # not required as customnotifications set to false -CustomFromName $null -CustomFromAddress $true -CustomInternalSubject $null -CustomInternalBody $null -CustomExternalSubject $null -CustomExternalBody $null

#Anti-Phishing Policies 
Get-AntiPhishPolicy -Identity * | select name, enabled | ft
Get-AntiPhishRule  -Identity * | select name, enabled | ft

#set policy
Set-AntiPhishPolicy -Identity "MFAT Custom Strict Policy" -EnableSimilarUsersSafetyTips $true -TargetedUserProtectionAction quarantine -EnableSpoofIntelligence $true -AuthenticationFailAction quarantine -EnableFirstContactSafetyTips $true -EnableUnauthenticatedSender $true -EnableViaTag $true

#Set Defender for O365 Antiphish policies. apply these settings to same policy.
#DOESNT WORK IN PROD, ERROR 
Set-AntiPhishPolicy -Identity "MFAT Custom Strict Policy" -PhishThresholdLevel 3 -EnableTargetedUserProtection $true -EnableOrganizationDomainsProtection $true -EnableTargetedDomainsProtection $true -EnableMailboxIntelligence $true -EnableMailboxIntelligenceProtection $true -TargetedDomainProtectionAction quarantine -TargetedUserProtectionAction quarantine -MailboxIntelligenceProtectionAction quarantine -EnableSimilarUsersSafetyTips $true -EnableSimilarDomainsSafetyTips $true -EnableUnusualCharactersSafetyTips $true

#Safe Attachment Policy
Connect-IPPSSession
Get-SafeAttachmentPolicy -Identity * | ft

#set policy
Set-AtpPolicyForO365 -EnableATPForSPOTeamsODB $true -EnableSafeDocs $true -AllowSafeDocsOpen $false -EnableSafeLinksForO365Clients $true -TrackClicks $true -AllowClickThrough $false

New-ActivityAlert -Name "Malicious Files in Libraries" -Description "Notifies admins when malicious files are detected in SharePoint Online, OneDrive, or Microsoft Teams" -Category ThreatManagement -Operation FileMalwareDetected -NotifyUser 'imdcloudopsalerts@mfat.govt.nz'


#Safe Links
Set-SafeLinksPolicy -Identity "MFAT Custom Strict Policy" -TrackClicks $true -AllowClickThrough $false


Disconnect-ExchangeOnline




