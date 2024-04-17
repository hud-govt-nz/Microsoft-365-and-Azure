AzureADPreview\Connect-AzureAD
$Setting = Get-AzureADDirectorySetting -Id (Get-AzureADDirectorySetting | where -Property DisplayName -Value "Group.Unified" -EQ).id
$Setting.Values
$Setting["EnableMIPLabels"] = "True"
Set-AzureADDirectorySetting -Id $Setting.Id -DirectorySetting $Setting


Connect-IPPSSession
Execute-AzureAdLabelSync

