#Add Dynamic AD Group - Uses Graph Module

$GroupUpdate = @{
    "DisplayName"="MHUD - M365 E3 - Test Dynamic License Group v2"
    "Description"="Testing Updating Group via MG Cmdlet"
    "MailEnabled"="$False"
    "SecurityEnabled"="$True"
    "GroupTypes"="DynamicMembership"
    "MembershipRule"="User.memberof -any (group.objectid -in 'bce3628a-23c1-48d8-b270-fbd399fa1998')"
    "MembershipRuleProcessingState"="On"
}

Update-MgGroup  -GroupId 'cbb8a172-26ec-416e-a0ba-1015b9b04181' -BodyParameter $GroupUpdate