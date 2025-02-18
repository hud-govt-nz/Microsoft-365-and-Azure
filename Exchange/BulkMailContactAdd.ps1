# CSV Import Path
$Path ="<CSV FILE PATH>"

# Import CSV to $Import Object
$Import = Import-csv -Path $Path
 
# Loop - Add contacts to Exchange Online & Update Distribution List Member
$Import| ForEach-Object {
 
    try {
        # Adding Contact
        New-MailContact -DisplayName $_.Person -ExternalEmailAddress $_.Email -Name $_.Person
        Write-Host "Successfully created contact for $($_.Email)" -ForegroundColor Green

        # Adding Contact to Distribution Group
        Add-DistributionGroupMember -Identity $_.DL -Members $_.Email
        Write-Host "Successfully added $($_.Email) to distribution list $($_.DL)" -ForegroundColor Green
    
    } catch {
        Write-Host "Error processing contact: $($_.Email)" -ForegroundColor Red
        Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Command in error: $($_.InvocationInfo.MyCommand)" -ForegroundColor Yellow
        Write-Host "----------------------------------------" -ForegroundColor Red
    }
}