<#
.SYNOPSIS
This script performs a specific task or set of tasks.

.DESCRIPTION
This PowerShell script is designed to automate a particular process or set of processes. It includes various functions and commands to achieve the desired outcome efficiently. The script is intended to be used by system administrators or users with appropriate permissions.

.PARAMETER <ParameterName>
Description of the parameter and its purpose.

.EXAMPLE
An example of how to run the script with appropriate parameters.

.NOTES
Additional information about the script, such as author, version, and date.

#>

Clear-Host

# Variables
$domain = 'mhud'
$dateTime = (Get-Date).toString("dd-MM-yyyy-hh-ss")
$directorypath = "C:\HUD\06_Reporting\SPO"

$SiteCollectionUrl = Read-Host -Prompt "Enter site collection URL "

$u = [System.Uri]$SiteCollectionUrl
$siteName = Split-Path $u.AbsolutePath -Leaf

$exportFilePath = Join-Path -Path "$directorypath\reports" -ChildPath "$($domain)-$($siteName)-everyone_$($dateTime).csv"
$loggingPath = Join-Path -Path $directorypath -ChildPath "Logs\transcript_$($siteName)_$($dateTime).txt"

# Start logging
Start-Transcript -Path $loggingPath -Append

# Function to get list items with unique permissions
$properties = @{
    SiteUrl           = ''
    SiteTitle         = ''
    ListTitle         = ''
    SensitivityLabel  = ''
    Type              = ''
    RelativeUrl       = ''
    ParentGroup       = ''
    MemberType        = ''
    MemberName        = ''
    MemberLoginName   = ''
    Roles             = ''
}
$excludeLimitedAccess = $true
$includeListsItems   = $true



# Removed the $everyoneGroups array to capture all unique permissions
$global:siteTitle = ""

#Exclude certain libraries
$ExcludedLibraries = @(
    "Form Templates", "Preservation Hold Library", "Site Assets", "Images",
    "Pages", "Settings", "Videos", "Timesheet", "Site Collection Documents",
    "Site Collection Images", "Style Library", "AppPages", "Apps for SharePoint",
    "Apps for Office"
)

$global:permissions  = @()
$global:sharingLinks = @()

function Get-ListItems_WithUniquePermissions {
    param(
        [Parameter(Mandatory)]
        [Microsoft.SharePoint.Client.List]$List
    )
    $selectFields = "ID,HasUniqueRoleAssignments,FileRef,FileLeafRef,FileSystemObjectType"
    $Url = $siteUrl + '/_api/web/lists/getbytitle(''' + $($list.Title) + ''')/items?$select=' + $($selectFields)
    $nextLink = $Url
    $listItems = @()
    $Stoploop = $true

    while ($nextLink) {
        do {
            try {
                $response = Invoke-PnPSPRestMethod -Url $nextLink -Method Get
                $Stoploop = $true
            }
            catch {
                Write-Host "An error occured: $_  : Retrying" -ForegroundColor Red
                $Stoploop = $true
                Start-Sleep -Seconds 30
            }
        }
        while ($Stoploop -eq $false)

        $listItems += $response.value | Where-Object { $_.HasUniqueRoleAssignments -eq $true }
        if ($response.'odata.nextlink') {
            $nextLink = $response.'odata.nextlink'
        }
        else {
            $nextLink = $null
        }
    }
    return $listItems
}

Function PermissionObject (
    $_object, $_type, $_relativeUrl, $_siteUrl, $_siteTitle,
    $_listTitle, $_memberType, $_parentGroup, $_memberName,
    $_memberLoginName, $_roleDefinitionBindings, $_sensitivityLabel
) {
    $permission = New-Object -TypeName PSObject -Property $properties
    $permission.SiteUrl          = $_siteUrl
    $permission.SiteTitle        = $_siteTitle
    $permission.ListTitle        = $_listTitle
    $permission.SensitivityLabel = $_sensitivityLabel
    $permission.Type             = $_Type -eq 1 ? "Folder" : $_Type -eq 0 ? "File" : $_Type
    $permission.RelativeUrl      = $_relativeUrl
    $permission.MemberType       = $_memberType
    $permission.ParentGroup      = $_parentGroup
    $permission.MemberName       = $_memberName
    $permission.MemberLoginName  = $_memberLoginName
    $permission.Roles            = $_roleDefinitionBindings -join ","
    $permission | Select-Object SiteUrl,SiteTitle,Type,SensitivityLabel,RelativeUrl,ListTitle,MemberType,MemberName,MemberLoginName,ParentGroup,Roles |
        Export-Csv -Path $exportFilePath -NoTypeInformation -Append
}

Function Extract-Guid ($inputString) {
    $splitString = $inputString -split '\|'
    return $splitString[2].TrimEnd('_o')
}

Function QueryUniquePermissionsByObject (
    $_ctx, $_object, $_Type, $_RelativeUrl, $_siteUrl, $_siteTitle, $_listTitle
) {
    $roleAssignments = Get-PnPProperty -ClientObject $_object -Property RoleAssignments
    switch ($_Type) {
        0       { $sensitivityLabel = $_object.FieldValues["_DisplayName"] }
        1       { $sensitivityLabel = $_object.FieldValues["_DisplayName"] }
        "Site"  { $sensitivityLabel = (Get-PnPSiteSensitivityLabel).displayname }
        default { " " }
    }

    foreach ($roleAssign in $roleAssignments) {
        Get-PnPProperty -ClientObject $roleAssign -Property RoleDefinitionBindings,Member
        $PermissionLevels = $roleAssign.RoleDefinitionBindings | Select -ExpandProperty Name

        # Exclude Limited Access if needed
        if ($excludeLimitedAccess -eq $true) {
            $PermissionLevels = ($PermissionLevels | Where { $_ -ne "Limited Access" }) -join ","
        }

        $MemberType = $roleAssign.Member.GetType().Name
        $PermissionType = $roleAssign.Member.PrincipalType

        if ($PermissionLevels.Length -gt 0) {
            # Return all unique permissions (removed filtering for 'everyone except external users')
            if ($roleAssign.Member.Title -notlike "SharingLinks*" -and ($MemberType -eq "Group" -or $MemberType -eq "User")) {
                if ($MemberType -eq "User") {
                    $ParentGroup = "NA"
                }
                else {
                    $ParentGroup = $roleAssign.Member.Title
                }
                PermissionObject $_object $_Type $_RelativeUrl $_siteUrl $_siteTitle $_listTitle $MemberType $ParentGroup $roleAssign.Member.Title $roleAssign.Member.LoginName $PermissionLevels $sensitivityLabel
            }

            if ($_Type -eq "Site" -and $MemberType -eq "Group") {
                $sensitivityLabel = (Get-PnPSiteSensitivityLabel).DisplayName
                if ($PermissionType -eq "SharePointGroup") {
                    $groupUsers = Get-PnPGroupMember -Identity $roleAssign.Member.LoginName
                    $groupUsers | ForEach-Object {
                        PermissionObject $_object "Site" $_RelativeUrl $_siteUrl $_siteTitle "" "GroupMember" $roleAssign.Member.LoginName $_.Title $_.LoginName $PermissionLevels $sensitivityLabel
                    }
                }
            }
        }
    }
}

Function QueryUniquePermissions($_web) {
    Write-Host "Querying web $($_web.Title)"
    $siteUrl = $_web.Url

    Write-Host $siteUrl -Foregroundcolor "Red"
    $global:siteTitle = $_web.Title
    $ll = Get-PnPList -Includes BaseType, Hidden, Title, HasUniqueRoleAssignments, RootFolder -Connection $siteconn |
        Where-Object { $_.Hidden -eq $False -and $_.Title -notin $ExcludedLibraries }

    Write-Host "Number of lists $($ll.Count)"

    QueryUniquePermissionsByObject $_web $_web "Site" "" $siteUrl $siteTitle ""

    $totalLists = $ll.Count
    $currentList = 0

    foreach ($list in $ll) {
        $currentList++
        $listUrl = $list.RootFolder.ServerRelativeUrl

        if ($list.Hidden -ne $True) {
            Write-Host "Processing list $currentList of $($totalLists): $($list.Title)" -Foregroundcolor "Yellow"
            $listTitle = $list.Title
            if ($list.HasUniqueRoleAssignments -eq $True) {
                $Type = $list.BaseType.ToString()
                QueryUniquePermissionsByObject $_web $list $Type $listUrl $siteUrl $siteTitle $listTitle
            }

            if ($includeListsItems) {
                $collListItem = Get-ListItems_WithUniquePermissions -List $list
                $count = $collListItem.Count
                Write-Host "Number of items with unique permissions: $count within list $listTitle"
                foreach ($item in $collListItem) {
                    $Type = $item.FileSystemObjectType
                    $fileUrl = $item.FileRef
                    $i = Get-PnPListItem -List $list -Id $item.ID
                    #$exportFilePath = Join-Path -Path $directorypath -ChildPath ("{0}-everyone_{1}.csv" -f $domain, $dateTime)
                    QueryUniquePermissionsByObject $_web $i $Type $fileUrl $siteUrl $siteTitle $listTitle
                }
            }
        }
    }
}

if (Test-Path $directorypath) {
    $env:PNPPOWERSHELL_UPDATECHECK = "Off"
    Connect-PnPOnline -Url $SiteCollectionUrl -ClientId $env:DigitalSupportAppID -Tenant 'mhud.onmicrosoft.com' -Thumbprint $env:DigitalSupportCertificateThumbprint

    $web = Get-PnPWeb
    QueryUniquePermissions($web)
}
else {
    Write-Host "Invalid directory path:" $directorypath -ForegroundColor "Red"
}

Stop-Transcript
