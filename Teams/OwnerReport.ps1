$Result = ""  
$Results = @() 
$Path = "./All Teams Members and Owner Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
Write-Host Exporting all Teams members and owners report...
$Count = 0
Get-Team | foreach {
    $TeamName = $_.DisplayName
    Write-Progress -Activity "`n     Processed Teams count: $Count "`n"  Currently Processing: $TeamName"
    $Count++
    $GroupId = $_.GroupId
    Get-TeamUser -GroupId $GroupId | foreach {
        $Name = $_.Name
        $MemberMail = $_.User
        $Role = $_.Role
        $Result = @{'Teams Name' = $TeamName; 'Member Name' = $Name; 'Member Mail' = $MemberMail; 'Role' = $Role }
        $Results = New-Object psobject -Property $Result
        $Results | select 'Teams Name', 'Member Name', 'Member Mail', 'Role' | Export-Csv $Path -NoTypeInformation -Append
    }
}
Write-Progress -Activity "`n     Processed Teams count: $Count "`n"  Currently Processing: $TeamName" -Completed
if ((Test-Path -Path $Path) -eq "True") {
    Write-Host `nReport available in $Path -ForegroundColor Green
}