Clear-Host
Write-Host '## Exchange Online: Remove Calendar Events for User ##' -ForegroundColor Yellow

# Requirements
#Requires -Modules ExchangeOnlineManagement

# Connect to Graph and Exchange
try {
    Connect-ExchangeOnline `
        -AppId $env:DigitalSupportAppID `
        -Organization "mhud.onmicrosoft.com" `
        -CertificateThumbprint $env:DigitalSupportCertificateThumbprint `
        -ShowBanner: $false
    Write-Host "Connected" -ForegroundColor Green
        
    } catch {
        Write-Host "Error connecting. Please check your credentials and network connection." -ForegroundColor Red
        exit 1
        }

#Calendar Form
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object Windows.Forms.Form -Property @{
    StartPosition = [Windows.Forms.FormStartPosition]::CenterScreen
    Size          = New-Object Drawing.Size 243, 250  # Adjusted size
    Text          = 'Select a Date'
    Topmost       = $true
}

$calendar = New-Object Windows.Forms.MonthCalendar -Property @{
    ShowTodayCircle   = $false
    MaxSelectionCount = 1
}
$form.Controls.Add($calendar)

$okButton = New-Object Windows.Forms.Button -Property @{
    Location     = New-Object Drawing.Point 30, 180  # Adjusted position
    Size         = New-Object Drawing.Size 75, 23
    Text         = 'OK'
    DialogResult = [Windows.Forms.DialogResult]::OK
}
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object Windows.Forms.Button -Property @{
    Location     = New-Object Drawing.Point 125, 180  # Adjusted position
    Size         = New-Object Drawing.Size 75, 23
    Text         = 'Cancel'
    DialogResult = [Windows.Forms.DialogResult]::Cancel
}
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

$Date = Get-Date -Format "dd.MM.yyyy"
Start-Transcript -Path "C:\HUD\01_Logs\Remove-CalendarEvents_$($Date).txt"

#Disclaimer
Write-Warning "Before proceeding please ensure the user is still enabled in Entra."
Write-Host ""

#Variables
$OrganiserMailbox = Read-host "Please provide the meeting organisers email address"
Write-Host "Please enter a start date" -ForegroundColor Green

#Open Calendar Dialog Box
$result = $form.ShowDialog()



#Select Result and input as Start date variable in EXO cmd.
if ($result -eq [Windows.Forms.DialogResult]::OK) {
    $date = [string]$calendar.SelectionStart
    [datetime]$StartDate = $date
    $FStartDate = $StartDate.ToString("dd-MM-yyyy")
    Write-Host "Date selected: $($FStartDate)"
}

#Remove Meetings where the user is the organizer (PreviewOnly)
Write-Host "The following Events have $($OrganiserMailbox) as the meeting organiser..." -ForegroundColor Yellow
Write-Host ""
Remove-CalendarEvents `
    -Identity $OrganiserMailbox `
    -CancelOrganizedMeetings `
    -QueryStartDate $FStartDate `
    -QueryWindowInDays 730 `
    -PreviewOnly `
    -Confirm:$false `

#Prompt for user confirmation
Write-Host ""
$UserConfirmation = Read-Host "Are you sure you want to proceed with deleting these events? (yes/no)"
if ($UserConfirmation -eq "yes") {
    #Remove Meetings where the user is the organizer (Actual Deletion)
    Write-Host "Removing Events..."
    Remove-CalendarEvents `
        -Identity $OrganiserMailbox `
        -CancelOrganizedMeetings `
        -QueryStartDate $FStartDate `
        -QueryWindowInDays 730 `
        -Confirm:$false `
} else {
    Write-Host "Operation cancelled by user."
}

Stop-Transcript