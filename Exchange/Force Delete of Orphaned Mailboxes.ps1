
#Export Mailbox list
$path = ""
Get-mailbox -RecipientTypeDetails Unlimited -IncludeInactiveMailbox -SoftDeletedMailbox | Select-Object Name,Alias | export-csv -Path $path

#Import Sanitised CSV file
$mailboxes = Import-csv -Path C:\temp\Mailboxes.csv | select-object Name

#test import if required
#return $mailboxes | ft -AutoSize -Wrap

#For each loop, remove orgHolds and Permanently Delete

$mailboxes | ForEach-Object {
    Set-Mailbox -Identity $_.Name -ExcludeFromAllOrgHolds
    Get-mailbox -Identity $_.Name -SoftDeletedMailbox| Remove-Mailbox -PermanentlyDelete -confirm:$false -force #-whatif
    }
