<# Default PS Profile on Cloud EOE boxes
    .PURPOSE
     This profile is setup to be loaded on any of the applocker managed servers so that powershell can be run by admins in a controlled context.

    .PROCESS
     1. To load this profile please ensure the server has been put into no inheritance applocker first, via AD.
     2. Once thats completed and the server is rebooted, open Powershell ISE and run the command psedit $profile.AllUsersAllHosts
     3. Copy and past the text below and then go file > save 
     4. close and re-open ISE and confirmed that powershell is working as expected.
     5. Move the server back into Applocker enforce mode and reboot.
#>

#Variables to add to profile
$env:PSModulePath = 'C:\Program Files\WindowsPowerShell\Modules'
$env:PSModulePath  += ';C:\Windows\system32\WindowsPowerShell\v1.0\Modules'
$env:PSModulePath  += ';C:\Program Files\Microsoft Monitoring Agent\Agent\PowerShell\'
$env:temp = "C:\Powershell\temp\"
$env:tmp = "C:\Powershell\temp\"

#Simple Function to Clear Variables and MEM in PS Session
Function ClearAll {
    Get-Variable -Exclude PWD,*Preference | Remove-Variable -EA 0;Remove-Module *; $error.Clear();
    [system.gc]::Collect()
    Clear-Host
    }

Remove-Item C:\PowerShell\4 -Confirm:$false -Force -Recurse -ErrorAction SilentlyContinue