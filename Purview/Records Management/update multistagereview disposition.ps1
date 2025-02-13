Connect-IPPSSession

# Define the parameters
$TagName = "DA697-4.2.2 Data analysis working records"
$MultiStageReviewProperty = 




'{"MultiStageReviewSettings":[{"StageName":"IM Review","Reviewers":["Ashley.Forde@hud.govt.nz"]}]}'

'{"MultiStageReviewSettings":[{"StageName":"IM Review","Reviewers":["DIGITALDISPOSITIONADMIN@hud.govt.nz"]}]}'


'{"MultiStageReviewSettings":[{"StageName":"Stage1","Reviewers":[jie@contoso.onmicrosoft.com]},{"StageName":"Stage2","Reviewers":[bharath@contoso.onmicrosoft.com,helen@contoso.onmicrosoft.com]},]}'

# Update the multistage review properties
Set-ComplianceTag -Identity $TagName -MultiStageReviewProperty $MultiStageReviewProperty