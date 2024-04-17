$exitCode = 0
#NOTE: The Proactive Remediations portal only shows the LAST LINE of output in the summary
# So I'll use this to print a useful summary just before exiting.
$exitSummary = ''

$QuickAssistCapabilityName = $(Get-WindowsCapability -Online -Name *quickassist*).Name
if ($(Get-WindowsCapability -Online -Name $QuickAssistCapabilityName).State -ne 'Installed') {
    $exitSummary = "The old Quick Assist was not found, remediation should not have run!"
    Write-Host $exitSummary
    $exitCode = 1
} else {
    Write-Host "Removing the old Quick Assist Capability"
    Remove-WindowsCapability -Online -Name $(Get-WindowsCapability -Online -Name *quickassist*).Name
    if ($(Get-WindowsCapability -Online -Name $QuickAssistCapabilityName).State -eq 'Installed') {
        $exitSummary = "Failed to remove the old Quick Assist!"
        Write-Host $exitSummary
        $exitCode = 1
    } else {
        $exitSummary = "Sucessfully removed the old Quick Assist"
        Write-Host $exitSummary
        $exitCode = 0
    }
}

# While we're here let's create a batch script that'll run the new Quick Assist
# It'll help if you ever want to run it from command line
$ScriptOutFile = "${env:windir}\system32\quickassist.cmd"
$SCriptEncoded = "QHBvd2Vyc2hlbGwuZXhlIC1leGVjdXRpb25wb2xpY3kgcmVtb3Rlc2lnbmVkIC1jb21tYW5kICIld2luZGlyJVxzeXN0ZW0zMlxxdWlja2Fzc2lzdC5wczEi"
if (-not $(Test-Path $ScriptOutFile)) {
    Write-Host "Did not find $ScriptOutFile"
    Write-Host "Creating $ScriptOutFile"
    $SCriptDecoded = [System.Convert]::FromBase64String($SCriptEncoded)
    $SCriptDecoded = [System.Text.Encoding]::UTF8.GetString($SCriptDecoded)
    $ScriptBytes = [System.Text.Encoding]::UTF8.GetBytes($SCriptDecoded)
    Set-Content -Path $ScriptOutFile -Value $ScriptBytes -Encoding Byte

    if (Test-Path $ScriptOutFile) {
        $exitSummary = $exitSummary + " Created $ScriptOutFile"
    }
}

# This is the powershell script the cmd script runs
# It'll make sure Quick Assist is available to the user trying to run it
$ScriptOutFile = "${env:windir}\system32\quickassist.ps1"
$SCriptEncoded = "IyBRdWljayBBc3Npc3QgQXBwIExhdW5jaGVyIHYwLjAxDQoNCiMgVG8gZWFzaWx5IHJ1biBRdWljayBBc3Npc3QgZnJvbSBhIGNvbW1hbmQgcHJvbXB0LCBmaXJzdA0KIyBTYXZlIHRoaXMgc2NyaXB0IHRvICdjOlx3aW5kb3dzXHN5c3RlbTMyXHF1aWNrYXNzaXN0LnBzMScsIHRoZW4gDQojIENvcHkgdGhlIG5leHQgbGluZSB0byAnYzpcd2luZG93c1xzeXN0ZW0zMlxxdWlja2Fzc2lzdC5jbWQnIA0KIyBAcG93ZXJzaGVsbC5leGUgLWV4ZWN1dGlvbnBvbGljeSByZW1vdGVzaWduZWQgLWNvbW1hbmQgIiV3aW5kaXIlXHN5c3RlbTMyXHF1aWNrYXNzaXN0LnBzMSINCiMgZnJvbSBhIGNvbWFuZCBwcm9tcHQsIGp1c3QgdHlwZSAncXVpY2thc3Npc3QnIHRvIHJ1biB0aGUgY21kIHdoaWNoIHdpbGwgcnVuIHRoZSBwczENCg0KJHdob2lhbSA9IFtzeXN0ZW0uc2VjdXJpdHkucHJpbmNpcGFsLndpbmRvd3NpZGVudGl0eV06OmdldGN1cnJlbnQoKS5uYW1lDQoNCiRBcHBOYW1lID0gIlF1aWNrIEFzc2lzdCINCiRQa2dOYW1lID0gJ01pY3Jvc29mdENvcnBvcmF0aW9uSUkuUXVpY2tBc3Npc3QnDQojJFBrZ05hbWUgPSAiKnF1aWNrYXNzaXN0KiINCg0KIyBJcyB0aGUgYXBwIHBhY2thZ2UgYWRkZWQgZm9yIHRoZSB1c2VyIHJ1bm5pbmcgdGhpcyBzY3JpcHQ/DQokUGFja2FnZSA9IEdldC1BcHB4UGFja2FnZSAtTmFtZSAkUGtnTmFtZSAtRXJyb3JBY3Rpb24gU2lsZW50bHlDb250aW51ZQ0KaWYgKCRQYWNrYWdlLlZlcnNpb24pIHsNCiAgICBXcml0ZS1Ib3N0ICIkd2hvaWFtIGhhcyAkQXBwTmFtZSAkKCRQYWNrYWdlLlZlcnNpb24pIGF2YWlsYWJsZSINCn0gZWxzZSB7DQogICAgV3JpdGUtSG9zdCAiJHdob2lhbSBkb2VzIE5PVCBoYXZlICRBcHBOYW1lIGF2YWlsYWJsZSINCiAgICAjIElzIHRoZSBBcHAgb24gdGhpcyBkZXZpY2U/DQogICAgJFBhY2thZ2UgPSBHZXQtQXBweFBhY2thZ2UgLUFsbFVzZXJzIC1OYW1lICRQa2dOYW1lIC1FcnJvckFjdGlvbiBTaWxlbnRseUNvbnRpbnVlDQogICAgaWYgKCRQYWNrYWdlLlZlcnNpb24pIHsNCiAgICAgICAgV3JpdGUtSG9zdCAiVGhpcyBkZXZpY2UgaGFzICRBcHBOYW1lICQoJFBhY2thZ2UuVmVyc2lvbikgYXZhaWxhYmxlIg0KDQogICAgICAgICMgVHJ5IHRvIGFkZCB0aGUgYXBwIGZvciB0aGUgdXNlcg0KICAgICAgICAkTWFuaWZlc3RQYXRoID0gKEdldC1BcHB4UGFja2FnZSAtQWxsVXNlcnMgLU5hbWUgIiRQa2dOYW1lIikuSW5zdGFsbExvY2F0aW9uICsgIlxBcHB4bWFuaWZlc3QueG1sIg0KICAgICAgICBBZGQtQXBweFBhY2thZ2UgLVBhdGggJE1hbmlmZXN0UGF0aCAtUmVnaXN0ZXIgLURpc2FibGVEZXZlbG9wbWVudE1vZGUNCg0KICAgICAgICAjIGNoZWNrIG91ciB3b3JrLi4uDQogICAgICAgICRQYWNrYWdlID0gR2V0LUFwcHhQYWNrYWdlIC1OYW1lICRQa2dOYW1lIC1FcnJvckFjdGlvbiBTaWxlbnRseUNvbnRpbnVlDQogICAgICAgIGlmICgkUGFja2FnZS5WZXJzaW9uKSB7DQogICAgICAgICAgICBXcml0ZS1Ib3N0ICJTdWNlc3NmdWxseSBhZGRlZCAkQXBwTmFtZSAkKCRQYWNrYWdlLlZlcnNpb24pIGZvciAkd2hvaWFtIg0KICAgICAgICB9IGVsc2Ugew0KICAgICAgICAgICAgV3JpdGUtSG9zdCAiRmFpbGVkIHRvIGFkZCAkQXBwTmFtZSBmb3IgJHdob2lhbSINCiAgICAgICAgfQ0KICAgIH0gZWxzZSB7DQogICAgICAgIFdyaXRlLUhvc3QgIlRoaXMgZGV2aWNlIGRvZWQgTk9UIGhhdmUgJEFwcE5hbWUgYXZhaWxhYmxlIg0KICAgIH0NCn0NCg0KIyBUZXN0aW5nPw0KIyBSZW1vdmUtQXBweFBhY2thZ2UgLVBhY2thZ2UgJChHZXQtQXBweFBhY2thZ2UgLU5hbWUgKnF1aWNrYXNzaXN0KikNCg0KIyBXaG8gaXMgcnVubmluZyBleHBsb3Jlci5leGU/IChpdCBjb3VsZCBiZSBtYW55IHVzZXJzIGJ1dCB3ZSBkb24ndCByZWFsbHkgZGVhbCB3aXRoIHRoYXQpDQokZXhlT3duZXJzID0gQChHZXQtV21pT2JqZWN0IC1DbGFzcyBXaW4zMl9Qcm9jZXNzIC1GaWx0ZXIgJ05hbWU9ImV4cGxvcmVyLmV4ZSInIHwgRm9yRWFjaC1PYmplY3QgeyAkXy5HZXRPd25lcigpIH0pDQokbG9nZ2VkT25Vc2VycyA9IEAoKQ0KZm9yZWFjaCAoJG93bmVyIGluICRleGVPd25lcnMpIHsNCiAgICAkbG9nZ2VkT25Vc2VycyArPSAoJ3swfVx7MX0nIC1mICgkb3duZXIgfFNlbGVjdC1PYmplY3QgLVVuaXF1ZSAtRXhwYW5kUHJvcGVydHkgRG9tYWluKSwoJG93bmVyIHxTZWxlY3QtT2JqZWN0IC1VbmlxdWUgLUV4cGFuZFByb3BlcnR5IFVzZXIpKQ0KfQ0KDQppZiAoJFBhY2thZ2UuUGFja2FnZUZhbWlseU5hbWUpIHsNCiAgICAjIE5PVEU6IFRoZSBhcHAgc2VlbSB0byBydW4gaW4gdGhlIGNvbnRleHQgb2Ygd2hhdGV2ZXIgdXNlciBvd25zIHRoZSBhbHJlYWR5IHJ1bm5pbmcgZXhwbG9yZXIuZXhlIHByb2Nlc3MNCiAgICAjIEV2ZW4gd2hlbiBydW5uaW5nIHRoaXMgc2NyaXB0IGFzIGEgZGlmZmVyZW50IChhZG1pbikgdXNlciwgdGhlIHNwYXduZWQgcHJvY2VzcyB3aWxsIGJlIHRoZSBvdGhlciB1c2VyLg0KICAgICMgTGV0J3MgY2hlY2sgc28gd2UgY2FuIHdhcm4gaWYgaXQgbG9va3Mgb2RkLi4uDQogICAgaWYgKCR3aG9pYW0gLWVxICRsb2dnZWRPblVzZXJzWzBdKSB7DQogICAgICAgIFdyaXRlLUhvc3QgIlN0YXJ0aW5nICRBcHBOYW1lIGZvciAkd2hvaWFtIg0KICAgIH0gZWxzZSB7DQogICAgICAgICNOT1RFOiBJIHNob3VsZCBwcm9iYWJseSBtYWtlIHN1cmUgdGhhdCB1c2VyIGhhcyB0aGUgYXBwIGFkZGVkIGFzIHdlbGwuLi4NCiAgICAgICAgV3JpdGUtV2FybmluZyAiZXhwbG9yZXIuZXhlIGlzIHJ1bm5pbmcgYXMgYSBkaWZmZXJlbnQgdXNlciEgJEFwcE5hbWUgd2lsbCBydW4gYXMgdGhhdCB1c2VyLi4uIg0KICAgICAgICBXcml0ZS1Ib3N0ICJTdGFydGluZyAkQXBwTmFtZSBmb3IgJCgkbG9nZ2VkT25Vc2Vyc1swXSkiDQogICAgfQ0KICAgIFN0YXJ0LVByb2Nlc3MgJ2V4cGxvcmVyLmV4ZScgLUFyZ3VtZW50TGlzdCAic2hlbGw6QXBwc0ZvbGRlclwkKCRQYWNrYWdlLlBhY2thZ2VGYW1pbHlOYW1lKSFBcHAiDQp9IGVsc2Ugew0KICAgIFdyaXRlLUhvc3QgIlVuYWJsZSB0byBzdGFydCAkQXBwTmFtZSBmb3IgJHdob2lhbSINCn0="
if (-not $(Test-Path $ScriptOutFile)) {
    Write-Host "Did not find $ScriptOutFile"
    Write-Host "Creating $ScriptOutFile"
    $SCriptDecoded = [System.Convert]::FromBase64String($SCriptEncoded)
    $SCriptDecoded = [System.Text.Encoding]::UTF8.GetString($SCriptDecoded)
    $ScriptBytes = [System.Text.Encoding]::UTF8.GetBytes($SCriptDecoded)
    Set-Content -Path $ScriptOutFile -Value $ScriptBytes -Encoding Byte

    if (Test-Path $ScriptOutFile) {
        $exitSummary = $exitSummary + " Created $ScriptOutFile"
    }
}

Write-Host $exitSummary
exit $exitCode


<# Encoded Script in Line 45, encoded as Base64 string $ScriptEncoded

# Quick Assist App Launcher v0.02

# To easily run Quick Assist from a command prompt, first
# Save this script to 'c:\windows\system32\quickassist.ps1', then 
# Copy the next line to 'c:\windows\system32\quickassist.cmd' 
# @powershell.exe -executionpolicy remotesigned -command "%windir%\system32\quickassist.ps1"
# from a comand prompt, just type 'quickassist' to run the cmd which will run the ps1

$whoiam = [system.security.principal.windowsidentity]::getcurrent().name

$AppName = "Quick Assist"
$PkgName = 'MicrosoftCorporationII.QuickAssist'
#$PkgName = "*quickassist*"

# Is the app package added for the user running this script?
$Package = Get-AppxPackage –Name $PkgName –ErrorAction SilentlyContinue
if ($Package.Version) {
    Write-Host "$whoiam has $AppName $($Package.Version) available"
} else {
    Write-Host "$whoiam does NOT have $AppName available"
    # Is the App on this device?
    $Package = Get-AppxPackage –AllUsers –Name $PkgName –ErrorAction SilentlyContinue
    if ($Package.Version) {
        Write-Host "This device has $AppName $($Package.Version) available"

        # Try to add the app for the user
        $ManifestPath = (Get-AppxPackage –AllUsers –Name "$PkgName").InstallLocation + "\Appxmanifest.xml"
        Add-AppxPackage –Path $ManifestPath –Register –DisableDevelopmentMode

        # check our work…
        $Package = Get-AppxPackage –Name $PkgName –ErrorAction SilentlyContinue
        if ($Package.Version) {
            Write-Host "Sucessfully added $AppName $($Package.Version) for $whoiam"
        } else {
            Write-Host "Failed to add $AppName for $whoiam"
        }
    } else {
        Write-Host "This device does NOT have $AppName available"
    }
}

# Testing?
# Remove-AppxPackage -Package $(Get-AppxPackage -Name *quickassist*)

if ($Package.PackageFamilyName) {
    # NOTE: The app seem to run in the context of whatever user owns the already running explorer.exe process
    # Even when running this script as a different (admin) user, the spawned process will be the other user.
    # Let's check so we can warn if it looks odd…

    # Who is running explorer.exe? (it could be many users but we don't really deal with that)
    $exeOwners = @(Get-WmiObject –Class Win32_Process –Filter 'Name="explorer.exe"' | ForEach-Object { $_.GetOwner() })
    $loggedOnUsers = @()
    foreach ($owner in $exeOwners) {
        $loggedOnUsers += ('{0}\{1}' -f ($owner |Select-Object –Unique –ExpandProperty Domain),($owner |Select-Object –Unique –ExpandProperty User))
    }

    if ($whoiam -eq $loggedOnUsers[0]) {
        Write-Host "Starting $AppName for $whoiam"
    } else {
        #NOTE: I should probably make sure that user has the app added as well…
        Write-Warning "explorer.exe is running as a different user! $AppName will run as that user…"
        Write-Host "Starting $AppName for $($loggedOnUsers[0])"
    }
    Start-Process 'explorer.exe' –ArgumentList "shell:AppsFolder\$($Package.PackageFamilyName)!App"
} else {
    Write-Host "Unable to start $AppName for $whoiam"
}
#>