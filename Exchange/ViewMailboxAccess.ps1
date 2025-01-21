#===========================================================================================================================================================================#
# Script Name:     View Mailbox Access (EXO)
# Author:          Ashley Forde
# Version:         1.0
# Description:     This Script is a easy way to run some basic exchange management functions. 
# Requires EXO powershell module installed on local device
#
# Version 1.0 - inital script 23.8.22
#===========================================================================================================================================================================

#Clear
Clear-Host
Write-Host "View Mailbox Access..."

#Connect to Exchange Online
Connect-ExchangeOnline

#Secondary Function(s)
function Get-MBXAddress ($MBXAlias) {
    if ($MBXAlias) {$SMTP = (Get-ADUser $MBXAlias -Properties EmailAddress | Select-Object EmailAddress).EmailAddress
        return $SMTP}}

#Primary Function
function Get-MailboxAccess {
    #Header
    Clear-Host
    Write-Host "========================== Mailbox Access Checker ==========================" -ForegroundColor Green
    Write-Host ""

    #Enter Email address or Alias
    $EmailAlias = Read-Host "Enter Alias of Mailbox or Full email Address i.e. jdoe or john.doe@mfat.govt.nz"
    
    #Try/Catch for Email address or Alias
    try {$email = Get-MBXAddress ($EmailAlias)}
        Catch {$email = $EmailAlias} 
    
    Write-host ""
    Write-Host "The following accounts have Full Permission to mailbox $email" -ForegroundColor Yellow
    Write-host ""

    #Get Mailbox Access - Full Access
    $Permission = (Get-MailboxPermission -Identity $email | Select-Object User,AccessRights).user
    
    #Name Formatting
    $PermissionOutput = foreach ($User in $Permission) {
        $Permission = [string]$User
        $Names = $Permission.Replace('ORANGE\','')
            foreach ($Name in $Names) {[string] $Name}}

    $PermissionOutput

    Write-host ""
    Write-host "The following users have send as permissions to mailbox $email" -ForegroundColor Yellow
    Write-host ""
    
    #Get Send As Access
    $SendAs = (Get-RecipientPermission -Identity $email | Select-Object -Property @{name="User";expression='Trustee'}, AccessRights).User
    
    #Name Formatting
    $SendAsOutput = foreach ($User in $SendAs) {
        $SendAs = [string]$User
        $Names = $SendAs.Replace('orange.mfat.net.nz/Orange Users/','')
        foreach ($Name in $Names) {[string] $Name}}

    $SendAsOutput
    
    #Export Results
    if ($email) {
        $title = "Export"
        $sendasmsg = "Would you like to export this to a .txt file?"
        $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", ` "Exports file"
        $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", ` "Closes script"
        $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
        $result = $host.ui.PromptForChoice($title, $sendasmsg, $options, 0) 
        switch ($result) {
            0 { #Yes
                If ($email) {
                    #Create log
                    $timestamp = Get-Date -Format "dd.MM.yy"
                    $saveto = "C:\Support\Common Tasks\04_View_Mailbox_Access\Output Log\Access-$email-$timestamp.txt"
                    Write-Output "The following users have 'Full Access' permissions to mailbox $email" | Out-File $saveto 
                    Write-Output $PermissionOutput | Out-File $saveto -Append
                    Write-Output "" | Out-File $saveto -Append
                    Write-Output "The following users have 'Send As' permissions to mailbox $email" | Out-File $saveto -Append
                    Write-Output $SendAsOutput | Out-File $saveto -Append

                    Write-Host "Log file will open shortly..."
                    start-sleep 3
                    Invoke-Item $saveto}}
            1 { #No
                exit}}     

        }
    [System.GC]::Collect()
    }
Get-MailboxAccess
