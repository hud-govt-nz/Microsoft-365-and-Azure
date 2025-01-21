$Policies = Get-RetentionCompliancePolicy | select GUID, Name, Workload, Enabled, Mode
$Policies | Export-Csv -Path C:\HUD\CompliancePolicies.csv -NoTypeInformation -Encoding UTF8 -Force