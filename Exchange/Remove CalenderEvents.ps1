<#
.NAME: Remove Calendar Events for Offboarded Staff Member
#>

#Calendar Form
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object Windows.Forms.Form -Property @{
    StartPosition = [Windows.Forms.FormStartPosition]::CenterScreen
    Size          = New-Object Drawing.Size 243, 230
    Text          = 'Select a Date'
    Topmost       = $true
}

$calendar = New-Object Windows.Forms.MonthCalendar -Property @{
    ShowTodayCircle   = $false
    MaxSelectionCount = 1
}
$form.Controls.Add($calendar)

$okButton = New-Object Windows.Forms.Button -Property @{
    Location     = New-Object Drawing.Point 38, 165
    Size         = New-Object Drawing.Size 75, 23
    Text         = 'OK'
    DialogResult = [Windows.Forms.DialogResult]::OK
}
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object Windows.Forms.Button -Property @{
    Location     = New-Object Drawing.Point 113, 165
    Size         = New-Object Drawing.Size 75, 23
    Text         = 'Cancel'
    DialogResult = [Windows.Forms.DialogResult]::Cancel
}
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

#Connect to EXO
#Connect-ExchangeOnline

#Transcript
$Date = Get-date -Format "dd.MM.yyyy HH:mm"
Start-Transcript -Path C:\Support\RemoveCalendarEvents_$Date

#Disclaimer
Write-Host "Before proceeding please ensure the user is still enabled in Azure AD."
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

#Remove Meetings where the user is the organis
Write-Host "Removing Events where $($OrganiserMailbox) is the meeting organiser..." -ForegroundColor Yellow
Remove-CalendarEvents `
    -Identity $OrganiserMailbox `
    -CancelOrganizedMeetings `
    -QueryStartDate $FStartDate `
    -QueryWindowInDays 365 `
    -Confirm:$false 
