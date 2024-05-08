Connect-IPPSSession

# Define the parameters
$TagName = "DA697-4.2.2 Data analysis working records"
$MultiStageReviewProperty = '{"MultiStageReviewSettings":[{"StageName":"IM Review","Reviewers":["Ashley.Forde@hud.govt.nz"]}]}'

# Update the multistage review properties
Set-ComplianceTag -Identity $TagName -MultiStageReviewProperty $MultiStageReviewProperty