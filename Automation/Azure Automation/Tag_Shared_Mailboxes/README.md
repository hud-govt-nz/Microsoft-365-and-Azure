# Automating Tagging Shared and Room Mailboxes

- The following checks daily whether or not any new shared or room mailbox resources have been created. Any new mailboxes are updated to reflect this in Entra.
- Currently Entra does not natively support shared or room mailbox attributes, this is handled by Exchange Online. 
- This is helpful in assigning these mailbox types to dynamic groups in Entra

For more information on setting up resources and accounts please read through the [Azure Automation/Teams_DDI_Manager/Resource_Creation.md](https://github.com/hud-govt-nz/Microsoft-365-and-Azure/blob/6a1b9c2930b5b31db47757a4e21b83ec2e2b88cd/Azure%20Automation/Teams_DDI_Manager/Resource_Creation.md)


## Outline 

### 1. Permissions
  - Graph and Exchange Online API Permission
  - Entra Roles Assignment to Managed Identity

### 2. Runbooks
  - Update-Shared-Mailbox-Job-Title

## 1 Permissions

### 1.1 Create Automation Account and add API permissions
```powershell  
# Create Azure Resource Group, Automation Account, and Managed Identity  
Import-Module Az  
Connect-AzAccount  
  
Set-AzContext -SubscriptionId "Enter HUD Azure Subscription ID"  
  
$resourceGroupName = "syd-rg-app-shared-mbx-monitor"  
$location = "Australia East"  
$automationAccountName = "SharedMBXMonitor" # Automation Account  
  
# Create the Resource Group  
New-AzResourceGroup -Name $resourceGroupName -Location $location  
  
# Create the Automation Account within the specified Resource Group  
New-AzAutomationAccount -ResourceGroupName $resourceGroupName -Name $automationAccountName -Location $location  
  
# Enable System Managed Identity on the Automation Account  
$automationAccount = Get-AzAutomationAccount -ResourceGroupName $resourceGroupName -Name $automationAccountName  
$automationAccount | Set-AzAutomationAccount -AssignSystemIdentity $true  
  
# Assign Graph permissions to Managed Identity  
Import-Module Microsoft.Graph   
Connect-MgGraph # can define scope here but since we are using this same session to also create the new user you can just connect generally.  
  
# Get GUID of Managed Identity  
$managedIdentityId = (Get-MgServicePrincipal -Filter "displayName eq 'SharedMBXMonitor'").id  
  
# Assign Graph Permissions to Managed Identity  
$graphApp = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'" #AppId of Microsoft Graph in all Enterprise Applications, always the same in each tenant.  
$graphScopes =@('User.ReadWrite.All','Grouo.ReadWrite.All','GroupMember.ReadWrite.All')  
  
ForEach($scope in $graphScopes){  
    $appRole = $graphApp.AppRoles | Where-Object {$_.Value -eq $scope}  
    New-MgServicePrincipalAppRoleAssignment -PrincipalId $managedIdentityId -ServicePrincipalId $managedIdentityId -ResourceId $graphApp.Id -AppRoleId $appRole.Id  
}  
  
# Assign Office 365 Exchange Online API Permissions to Managed Identity  
$ExchangeAPI =@('Exchange.ManageAsApp')  
$ExchangeAPI = Get-MgServicePrincipal -Filter "AppId eq '00000002-0000-0ff1-ce00-000000000000'" #AppId of Office 365 Exchange Online API in all Enterprise Applications, always the same in each tenant.  
  
ForEach($scope in $ExchangeAPI){  
    $appRole = $SkypeAPI.AppRoles | Where-Object {$_.Value -eq $scope}  
    New-MgServicePrincipalAppRoleAssignment -PrincipalId $managedIdentityId -ServicePrincipalId $managedIdentityId -ResourceId $SkypeAPI.Id -AppRoleId $appRole.Id  
}  
```

### 1.2 Assign Entra Role to Managed Identity
Assigning Enterprise Application roles to a managed identity via the Azure Portal involves a few steps. Here's how you can do it:

1. **Login to Azure Portal**:
   - Navigate to the [Azure Portal](https://portal.azure.com/) and log in with your credentials.

2. **Access Azure Active Directory**:
   - In the left-hand menu, click on "Microsoft Entra ID".

3. **Navigate to Enterprise Applications**:
   - In the Azure AD blade, find and click on "Enterprise applications".

4. **Select the Enterprise Application**:
   - Find and click on the enterprise application you want to assign roles from.

5. **Access Role Assignment**:
   - In the enterprise application blade, find and click on "Roles and administrators" or it might be labeled as "Roles and groups" depending on the app configuration.

6. **Add Role Assignment**:
   - Click on "+ Add Assignment" or "+ Add Role Assignment".
   - In the "Role" dropdown, select the role you want to assign.
   - In the "Assign access to" dropdown, select "User, group, or service principal".
   - In the "Select" field, type the name of the managed identity, select it from the list, and then click on the "Select" button at the bottom.
   - Click on the "Assign" button to finalize the role assignment.

7. **Verify Role Assignment**:
   - You can verify the role assignment by going back to the "Roles and administrators" or "Roles and groups" section and checking the list of role assignments.

## 2. Update-Shared-Mailbox-Job-Title Runbook

[Update-Shared-Mailbox-Job-Title](https://github.com/hud-govt-nz/Microsoft-365-and-Azure/blob/9e62a43fabb4ca6a104e7018bffec97742dc892c/Azure%20Automation/Tag_Shared_Mailboxes/Runbooks/Update-Shared-Mailbox-Job-Title.ps1)

This powershell runbook does 3 things.

1. Obtains a list of all Shared and Room mailboxes from Exchange
2. Takes the result and assigns an extension attribute with the value "Shared Mailbox" or "Room Mailbox"
3. Updates the Entra Objects Job Title

You can now use dynamic groups in Entra to group/manage these Mailbox owned Entra object IDs.




