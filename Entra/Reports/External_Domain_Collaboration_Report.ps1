# Connect to Graph and Teams and SharePoint Online
Connect-MgGraph -Scopes "IdentityProvider.Read.All policy.read.all CrossTenantInformation.ReadBasic.All SharePointTenantSettings.Read.All" | Out-Null
$Login = (Get-MgContext).Account
Connect-MicrosoftTeams -AccountId $Login | Out-Null

$env:PNPPOWERSHELL_UPDATECHECK="false"
$siteUrl = "https://mhud.sharepoint.com/sites/Hub"
Connect-PnPOnline -Url $SiteURL -Interactive


function GetAllowedGuestDomains {
    $B2BManagementPolicy = ((Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/legacy/policies").value.definition | ConvertFrom-Json).b2bmanagementpolicy
    return $B2BManagementPolicy.InvitationsAllowedAndBlockedDomainsPolicy.alloweddomains
}

function New-DomainSettingsObject {
    param(
        [string]$Domain,
        [string]$GuestInvitations
    )

    return [pscustomobject] @{
        Domain                       = $Domain
        GuestInvitations             = $GuestInvitations
        TeamsFederation              = ""
        SharePointSharing            = ""
        tenantID                     = ""

    }
}

function Update-SharePointSharing {
    param(
        [string]$sharingCapability,
        [string]$sharingDomainRestrictionMode,
        [array]$sharingAllowedDomainList,
        [array]$domainSettingsObjectArray
    )

    switch -wildcard ($sharingCapability, $sharingDomainRestrictionMode) {
        "Disabled" {
            $domainSettingsObjectArray | Where-Object { $_.domain -eq "Default" } | ForEach-Object { $_.SharePointSharing = "Blocked" }
        }
        "*none*" {
            $domainSettingsObjectArray | Where-Object { $_.domain -eq "Default" } | ForEach-Object { $_.SharePointSharing = "Allowed" }
        }
        "*AllowList" {
            $domainSettingsObjectArray | Where-Object { $_.domain -eq "Default" } | ForEach-Object { $_.SharePointSharing = "Blocked" }
            foreach ($domain in $sharingAllowedDomainList) {
                if ($domainSettingsObjectArray.Domain -contains $domain) {
                    $domainSettingsObjectArray | Where-Object { $_.domain -eq $domain } | ForEach-Object { $_.SharePointSharing = "Allowed" }
                } else {
                    $domainSettingsObject = [pscustomobject]@{
                        Domain                       = $domain
                        GuestInvitations             = ""
                        TeamsFederation              = ""
                        SharePointSharing            = "Allowed"
                        tenantID                     = ""

                    }
                    $domainSettingsObjectArray += $domainSettingsObject
                }
            }
        }
    }
}

function Update-DomainSettingsFromTeamsFederation {
    param(
        [object]$TeamsFederationSettings,
        [array]$domainSettingsObjectArray
    )

    if (($TeamsFederationSettings.alloweddomains -eq "AllowAllKnownDomains") -and ($TeamsFederationSettings.AllowFederatedUsers)) {
        $domainSettingsObjectArray | Where-Object { $_.domain -eq "Default" } | ForEach-Object { $_.TeamsFederation = "Allowed" }
        foreach ($domain in [array]$TeamsFederationSettings.BlockedDomains.domain) {
            Update-DomainSettings $domain $domainSettingsObjectArray "Blocked"
        }
    } else {
        $domainSettingsObjectArray | Where-Object { $_.domain -eq "Default" } | ForEach-Object { $_.TeamsFederation = "Blocked" }
        foreach ($domain in [array]$TeamsFederationSettings.AllowedDomains.AllowedDomain.domain) {
            Update-DomainSettings $domain $domainSettingsObjectArray "Allowed"
        }
    }
}

function Update-DomainSettings {
    param(
        [string]$domain,
        [array]$domainSettingsObjectArray,
        [string]$federationStatus
    )

    if ($domainSettingsObjectArray.Domain -contains $domain) {
        $domainSettingsObjectArray | Where-Object { $_.domain -eq $domain } | ForEach-Object { $_.TeamsFederation = $federationStatus }
    } else {
        $domainSettingsObject = [pscustomobject]@{
            Domain                       = $domain
            GuestInvitations             = "Org Default"
            TeamsFederation              = $federationStatus
            SharePointSharing            = ""
            tenantID                     = ""


        }
        $domainSettingsObjectArray += $domainSettingsObject
    }
}

function Update-TenantIDs {
    param(
        [array]$domainSettingsObjectArray
    )

    foreach ($domain in $domainSettingsObjectArray) {
        Try {
            $tenantID = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/tenantRelationships/findTenantInformationByDomainName(domainName='$($domain.Domain)')").tenantID
            $domain.tenantID = $tenantID
        }
        Catch {
            $tenantID = "N/A"
        }
        $domain.tenantID = $tenantID
    }
}


# ENTRA - External Collaboration Allowed Domains
$domainSettingsObjectArray = @()
$domainSettingsObjectArray += New-DomainSettingsObject -Domain "Default" -GuestInvitations "Blocked"

$allowedGuestDomains = GetAllowedGuestDomains

foreach ($domain in $allowedGuestDomains) {
    $domainSettingsObjectArray += New-DomainSettingsObject -Domain $domain -GuestInvitations "Allowed"
}


# SPO - SharePoint Sharing
$Uri = "https://graph.microsoft.com/beta/admin/sharepoint/settings"
$SPOSettings = Invoke-MgGraphRequest -Uri $Uri -Method Get
Update-SharePointSharing -sharingCapability $SPOSettings.sharingCapability -sharingDomainRestrictionMode $SPOSettings.sharingDomainRestrictionMode -sharingAllowedDomainList $SPOSettings.sharingAllowedDomainList -domainSettingsObjectArray $domainSettingsObjectArray

 
# TEAMS - Teams Federation
$TeamsFederationSettings = Get-CsTenantFederationConfiguration
Update-DomainSettingsFromTeamsFederation -TeamsFederationSettings $TeamsFederationSettings -domainSettingsObjectArray $domainSettingsObjectArray

Update-TenantIDs -domainSettingsObjectArray $domainSettingsObjectArray

# Output to screen and file
$ConditionalFormat =$(
    New-ConditionalText -Range '1:1' -BackgroundColor Blue -ConditionalTextColor White
)

$FilePath = ".\Allowed Domains.xlsx"

$domainSettingsObjectArray | Export-Excel $FilePath -AutoSize -AutoFilter -WorksheetName 'External Domains' -FreezeTopRow -BoldTopRow -ConditionalFormat $ConditionalFormat

Start-Sleep 3

#Upload the Excel file to SharePoint
$fileName = "Allowed Domains.xlsx"
$listName = "Shared Documents"
$list = Get-PnPList -Identity $listName

Add-PnPFile -Path $FilePath -Folder "Shared Documents" -NewFileName $fileName

