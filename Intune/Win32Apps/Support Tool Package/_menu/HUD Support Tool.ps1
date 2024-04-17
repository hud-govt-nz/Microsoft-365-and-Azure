function Show-Menu {
    Write-Host "1. Elevate PIM roles"
    Write-Host "2. Create a new user in Azure AD"
    Write-Host "3. Create a new shared mailbox in Exchange Online"
    Write-Host "4. Check users mailbox access in Exchange Online"
    Write-Host "5. Add a DDI Phone Number to a user"
    Write-Host "6. Add/Change Aho 'Employee Category' to existing user"
    Write-Host "7. Change users email address in Exchange Online"
    Write-Host "8. Run Basic Reports"
    Write-Host "Q. Exit"
    Write-Host ""
    $option = Read-Host "Enter your choice (1-5 or Q to exit)"
    return $option
}
####################################################################################












####################################################################################

#Select task based on Show-Menu function
do
{
    Clear-Host
    Write-Host ""
    Write-Host '## Digital Support Common Tasks ##' -ForegroundColor Yellow
    $selection = Show-Menu
    switch ($selection) {
                 '1' {New-PIMSession}
                 '2' {New-HUDUser}
                 '3' {New-SharedMailbox}
                 '4' {Search-MailboxAccess}
                 '5' {Add-PhoneNumber}
                 '6' {Add-EmployeeCategory}
                 '7' {Update-UserName}
                 '8' {Export-Reports}
                 'q' {return}
                 }
        pause
        }
until ($selection -eq 'q')