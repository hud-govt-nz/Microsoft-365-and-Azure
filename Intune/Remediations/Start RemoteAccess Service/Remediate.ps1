# Remediate
$RemoteAccess = Get-service -Name RemoteAccess
$RemoteAccess | Set-Service -StartupType Automatic  
$RemoteAccess | Start-Service