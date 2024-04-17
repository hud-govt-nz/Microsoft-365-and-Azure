$AdminURL = "https://mhud-admin.sharepoint.com/"  
$env:PNPPOWERSHELL_UPDATECHECK="false"
Connect-PnPOnline -url $AdminURL -Interactive

$Sites = Get-PnPTenantSite | Where-Object {$_.URL -like "https://mhud.sharepoint.com/sites/*" -and $site.StorageQuota -ne 0}

$ResultsArray = @()
$i = 0
$Total = $Sites.Count
Write-Progress -Activity "Progress" -Status "Beginning site search" -PercentComplete 0
foreach ($site in $sites) {
    $i++
    $pctComplete = [Math]::Truncate(($i/$Total)*100)
    Write-Progress -Activity "Progress" -Status "$i of $Total sites processed" -PercentComplete $pctComplete

    try {
    $siteURL = $site.Url

    $SiteConn = Connect-PnPOnline -url $siteURL -Interactive -ReturnConnection 
    $error.clear()  

    $web = Get-PnPWeb -Connection $SiteConn –Includes AppTiles
    $appTiles = $web.AppTiles
    Invoke-PnPQuery

    foreach ($appTile in $appTiles)
    {   
        #Filter out to show only the Apps that are the AppType: "Instance"
        If($appTile.AppType -eq "Instance"){

            #Create a new custom object and add to the array
            $ResultsObject = New-Object -TypeName PSObject
            $ResultsObject | Add-Member -MemberType NoteProperty -Name SiteURL -Value $SiteUrl
            $ResultsObject | Add-Member -MemberType NoteProperty -Name Title -Value $appTile.Title
            $ResultsObject | Add-Member -MemberType NoteProperty -Name AppType -Value $appTile.AppType
            $ResultsObject | Add-Member -MemberType NoteProperty -Name AppStatus -Value $appTile.AppStatus
            $ResultsObject | Add-Member -MemberType NoteProperty -Name AppSource -Value $appTile.AppSource
            $ResultsObject | Add-Member -MemberType NoteProperty -Name IsCorporateCatalogSite -Value $appTile.IsCorporateCatalogSite

            $ResultsArray += $ResultsObject
        }
    }  
    } Catch {
        Write-Output "Unable to connect to PnP site, please check site permissions for: $SiteUrl "
        Continue
    }
}


$ResultsArray | Export-Excel C:\HUD\SPOApps.xlsx -AutoSize -AutoFilter -WorksheetName 'Apps' -FreezeTopRow -BoldTopRow