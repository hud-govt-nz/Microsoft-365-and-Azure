# DLM Labels in Purview > DLM > Retention Labels

# Get all retention policies
Get-RetentionCompliancePolicy | Sort-Object Name | Export-csv

# Get all retention labels
Get-ComplianceTag | Sort-Object Name | Format-Table -Auto Name,Priority,RetentionAction,RetentionDuration,Workload


$itemsPendingDisposition = Get-ReviewItems -TargetLabelId 10740682-e941-4d9c-9cd5-df2c777b857a -IncludeHeaders $true -Disposed $true

$formattedExportItems = $itemsPendingDisposition.ExportItems | ConvertFrom-Csv -Header $itemsPendingDisposition.Headers

$formattedExportItems | Select Subject,Location,ReviewAction,Comment,DeletedBy,DeletedDate



