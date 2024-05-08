# Connect to Exchange Online ProtectioAdminn (EOP)
Connect-IPPSSession

$reviewTags = Get-ComplianceTag | Where-Object { $_.IsReviewTag -eq $true }

foreach ($tag in $reviewTags) {
    $reviewers = $tag.ReviewerEmail
    if (!$reviewers) { 
        Write-Host "Compliance tag $($tag.Name) has no reviewers, skipping..." -ForegroundColor Yellow
        continue }

    Set-ComplianceTag -Identity $tag.Name -ReviewerEmail "DIGITALDISPOSITIONADMIN@hud.govt.nz"
    Write-Host "Compliance tag $($tag.Name) has successfully had its reviewers updated" -ForegroundColor Green
}

