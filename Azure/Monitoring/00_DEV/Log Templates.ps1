#Look up sign in logs of user
SigninLogs
| where Identity contains "Jenny"

SigninLogs
| where UserPrincipalName == "jenny@frickelsoft.net"

SigninLogs
| where UserPrincipalName == "jenny@frickelsoft.net"
| where TimeGenerated > ago(60d)

SigninLogs
| where AppDisplayName == "Office 365 Exchange Online"
| where ClientAppUsed == "Browser"
| summarize requests = count() by OriginalRequestId, UserPrincipalName

SigninLogs
| where UserPrincipalName == "admin@testtenant.onmicrosoft.com"
| where TimeGenerated > ago(7d)
| summarize count() by TimeGenerated
| render timechart

AuditLogs
| where TargetResources contains "a0fdc91a-a1b2-4ec5-b352-03bda610be0e"
| where TimeGenerated > ago(7d)
| summarize a = count() by ActivityDisplayName
| render piechart

#Finding changes to objects

AuditLogs
| where Category == "GroupManagement"
| where ActivityDisplayName == "Add member to group"
| project Identity, TargetResources[0].modifiedProperties[1].newValue, TargetResources[0].userPrincipalName

AuditLogs
| where TargetResources contains "a0fdc91a-a1b2-4ec5-b352-03bda610be0e"
| where TimeGenerated > ago(7d)

AuditLogs
| where Category == "UserManagement"
| where OperationName == "Invite external user" or OperationName == "Redeem external user invite"
| project ActivityDateTime, OperationName, TargetResources[0].displayName
| order by ActivityDateTime