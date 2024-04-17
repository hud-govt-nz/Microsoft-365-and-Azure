#
# Copyright (C) Microsoft Corporation.  All rights reserved.
#

<#
.SYNOPSIS
    Script for fetching all Stream Classic videos and exporting to a CSV
.AADTENANTID
    Aad Tenant Id of the customer.
.INPUTFILE
    File Path to import the Stream token from. EX: "C:\Users\Username\Desktop\token.txt"
.OUTDIR
    Folder Path where CSV will be exported. EX: "C:\Users\Username\Desktop"
.RESUMELASTRUN
    True/False. Whether execution should be resumed from last run or scratch. Default value is true.
.PUBLISHEDDATELE
    yyyy-mm-dd. Optional Parameter. Fetches video entries for which PublishedDate less than value. Default filter not applied. EX: "2021-02-15"
.PUBLISHEDDATEGE
    yyyy-mm-dd. Optional Parameter. Fetches video entries for which PublishedDate greater than value. Default filter not applied. EX: "2021-02-15"
.CREATEDESTINATIONPATHMAPPINGFORM365GROUPCONTAINERS
    True/False. Optional Parameter. If set true, the script will create a destination path mapping for M365Group containers. Inventory report generation will not be done.
.MIGRATIONDESTINATIONCSVFILEPATH
    File path to import details of M365Group containers, for which destination path mapping is required. Optional Parameter.
.CWCCREATORDETAILMAPPING
    True/False. Optional Parameter. If set true, the script will create a mapping of CWC creator details.
.CWCCREATORCSVFILEPATH
    File path to import details of CWC containers, for which creator details mapping is required. Optional Parameter.
.GENERATEMASTERCONTAINERLIST
    True/False. Optional Parameter. If set true, the script will generate a master list of all containers in the tenant.

Example:
.\StreamClassicVideoReportGenerator.ps1 -AadTenantId "00000000-0000-0000-0000-000000000000" -InputFile "C:\Users\Username\Desktop\token.txt" -OutDir "C:\Users\Username\Desktop" -ResumeLastRun true
Example:
.\StreamClassicVideoReportGenerator.ps1 -AadTenantId "00000000-0000-0000-0000-000000000000" -InputFile "C:\Users\Username\Desktop\token.txt" -OutDir "C:\Users\Username\Desktop" -ResumeLastRun true -PublishedDateLe "2022-02-15"
Example:
.\StreamClassicVideoReportGenerator.ps1 -AadTenantId "00000000-0000-0000-0000-000000000000" -InputFile "C:\Users\Username\Desktop\token.txt" -OutDir "C:\Users\Username\Desktop" -ResumeLastRun true -PublishedDateGe "2022-02-15"
Example:
.\StreamClassicVideoReportGenerator.ps1 -AadTenantId "00000000-0000-0000-0000-000000000000" -InputFile "C:\Users\Username\Desktop\token.txt" -OutDir "C:\Users\Username\Desktop" -ResumeLastRun true -PublishedDateLe "2022-02-15" -PublishedDateGe "2021-02-15"
Example:
.\StreamClassicVideoReportGenerator.ps1 -AadTenantId "00000000-0000-0000-0000-000000000000" -InputFile "C:\Users\Username\Desktop\token.txt" -OutDir "C:\Users\Username\Desktop" -CreateDestinationPathMappingForM365GroupContainers true -MigrationDestinationCsvFilePath "C:\Users\Username\Desktop\MigrationDestinations.csv"
Example:
.\StreamClassicVideoReportGenerator.ps1 -AadTenantId "00000000-0000-0000-0000-000000000000" -InputFile "C:\Users\Username\Desktop\token.txt" -OutDir "C:\Users\Username\Desktop" -CWCCreatorDetailMapping true -CWCCreatorCsvFilePath C:\Users\Username\Desktop\CWC\CWCContainerIds.csv
Example:
.\StreamClassicVideoReportGenerator.ps1 -AadTenantId "00000000-0000-0000-0000-000000000000" -InputFile "C:\Users\Username\Desktop\token.txt" -OutDir "C:\Users\Username\Desktop" -GenerateMasterContainerList true
#>

[CmdletBinding()]
Param(
    [Parameter(Mandatory = $true)]
    [string]$AadTenantId,

    [Parameter(Mandatory = $true)]
    [string]$InputFile,

    [Parameter(Mandatory = $true)]
    [string]$OutDir,

    [Parameter(Mandatory = $false)]
    [string]$ResumeLastRun = 'true',

    [Parameter(Mandatory = $false)]
    [string]$PublishedDateLe,

    [Parameter(Mandatory = $false)]
    [string]$PublishedDateGe,

    [Parameter(Mandatory = $false)]
    [string]$CreateDestinationPathMappingForM365GroupContainers = 'false',

    [Parameter(Mandatory = $false)]
    [string]$MigrationDestinationCsvFilePath = '',

    [Parameter(Mandatory = $false)]
    [string]$CWCCreatorDetailMapping = 'true',

    [Parameter(Mandatory = $false)]
    [string]$CWCCreatorCsvFilePath,

    [Parameter(Mandatory = $false)]
    [string]$GenerateMasterContainerList = 'false'
)

Function GetBaseUrl {
    $tenantPatchUri = "https://api.microsoftstream.com/api/tenants/" + $AadTenantId + "?api-version=1.4-private"

    $headers = @{
        Authorization = "Bearer $token"
    }
    $body = "{}"

    ((Get-Date).tostring() + ' TenantPatch URI: ' + $tenantPatchUri + "`n") | Out-File $logFilePath -Append

    try {
        $response = Invoke-RestMethod -Uri $tenantPatchUri -Method Patch -Body $body -Headers $headers -ContentType "application/json"
    }
    catch {
        #Log error.
        ((Get-Date).tostring() + ' ' + $Error[0].ErrorDetails + "`n") | Out-File $logFilePath -Append
        ((Get-Date).tostring() + ' ' + $Error[0].Exception + "`n") | Out-File $logFilePath -Append

        #Stop execution if Unauthorized(401).
        if ($_.Exception.Response.StatusCode.value__ -eq 401) {
            Write-Host "========Enter new token and start the script again======="
        }

        Write-Host "`nSome error occurred. Check logs for more info.`n" -ForegroundColor Red
        exit
    }

    return $response.apiEndpoint
}

Function ReportOrchestration {
    Param(
        $baseUrl,
        $offsetId
    )

    $orchestrationUri = $baseUrl + 'migrationReports/orchestration?api-version=1.0-odsp-migration'

    if ($filter.Length -ne 0) {
        $orchestrationUri += '&$filter=' + $filter
    }

    if ($offsetId) {
        $orchestrationUri += '&$skiptoken=offsetId:' + $offsetId
    }

    $headers = @{
        Authorization = "Bearer $token"
    }

    ((Get-Date).tostring() + ' ReportOrchestration URI: ' + $orchestrationUri + "`n") | Out-File $logFilePath -Append

    try {
        $response = Invoke-RestMethod -Uri $orchestrationUri -Method Get -Headers $headers -ContentType "application/json"
    }
    catch {
        #Log error.
        ((Get-Date).tostring() + ' ' + $Error[0].ErrorDetails + "`n") | Out-File $logFilePath -Append
        ((Get-Date).tostring() + ' ' + $Error[0].Exception + "`n") | Out-File $logFilePath -Append

        #Stop execution if Unauthorized(401).
        if ($_.Exception.Response.StatusCode.value__ -eq 401) {
            Write-Host "========Enter new token and start the script again======="
        }

        #Stop execution if BadRequest(400)
        if ($_.Exception.Response.StatusCode.value__ -eq 400) {
            $errorMessage = $Error[0].ErrorDetails
            Write-Host "$errorMessage"
        }

        Write-Host "`nSome error occurred. Check logs for more info.`n" -ForegroundColor Red
        exit
    }

    return $response
}

Function CreateDirectory {
    Param(
        $category
    )

    $directoryPath = $OutDir + '\' + $category

    Write-Host "Checking if $category directory exists or not..."

    #Create directory if it does not exist.
    if (!(Test-Path $directoryPath)) {
        Write-Host "Not found. Creating directory..."

        New-Item -Path $directoryPath -ItemType Directory | Out-Null

        Write-Host "Created directory. Path: $directoryPath."
    }
    else {
        Write-Host "Directory found."
    }

    return $directoryPath
}

Function MergeAndDedupe {
    $finalReportName = $reportPath + '\Report_' + $timeStamp;

    #Fetch all the intermediate CSVs, merge, sorting based on a unique Id - VideoId, remove duplicate entries
    $csvfile = Get-ChildItem -Filter *.csv -Path $reportPath -Recurse | Select-Object -ExpandProperty FullName | Import-Csv | Sort-Object VideoId -Unique 
    $totalVideoCount = $csvfile.Count - 1;
    $csvfile | Export-CSV "$finalReportName.csv" -NoTypeInformation

    Write-Host "`n******************************************************************************"
    Write-Host "Reports are available at this location: $finalReportName.csv" -ForegroundColor Green
    Write-Host "******************************************************************************`n"

    Write-Host "Number of videos discovered: $totalVideoCount"
    ((Get-Date).tostring() + " Number of videos discovered: $totalVideoCount`n") | Out-File $logFilePath -Append
}

Function Dedupe {

    Param(
        $csvPath,
        $dateTimeNameSuffix
    )

    Write-Host "`nPreparing output directory..."

    $reportPath = CreateDirectory 'report'

    $reportFullName = $reportPath + '\StreamClassicVideoReport_' + $dateTimeNameSuffix + ".csv"

    #Fetch all the temporary file(temp CSVs), merge, sorting based on a unique Id - VideoId, remove duplicate entries and keep the last entry it found in the merged file.
    $csvfile = Get-ChildItem -Filter *.csv -Path $csvPath | Select-Object -ExpandProperty FullName | Import-Csv | Sort-Object VideoId -Unique

    $finalReportName = 'StreamClassicVideoReport_' + $dateTimeNameSuffix

    $reportsFolderPath = CreateDirectory ('report\' + $finalReportName)

    $fileName = $reportsFolderPath + '\' + $finalReportName + '_'

    #SPLIT the deduped file to multiple CSVs of 10k records each.

    # variable used to advance the number of the row from which the export starts.
    $startrow = 0

    # counter used in names of resulting CSV files
    $counter = 1

    while ($startrow -lt $csvfile.Count) {
        #pick 10k records starting from the $startrow position and export content to a new file.
        $csvfile | Select-Object -skip $startrow -first 10000 | Export-CSV "$fileName$($counter).csv" -NoTypeInformation

        # Increment the number of the rows from which the export starts.
        $startrow += 10000

        # incrementing the $counter variable.
        $counter++

    }

    Write-Host "`n******************************************************************************"
    Write-Host "Reports are available at this location: $reportsFolderPath"
    Write-Host "******************************************************************************`n"

    return $csvfile.Count
}

Function DedupeMasterList {

    Param(
        $csvPath,
        $dateTimeNameSuffix
    )

    Write-Host "`nPreparing output directory..."

    $reportPath = CreateDirectory 'report'

    $reportFullName = $reportPath + '\StreamClassicMasterContainerList_' + $dateTimeNameSuffix + ".csv"

    #Fetch all the temporary file(temp CSVs), merge, sorting based on a unique Id - VideoId, remove duplicate entries and keep the last entry it found in the merged file.
    $csvfile = Get-ChildItem -Filter *.csv -Path $csvPath | Select-Object -ExpandProperty FullName | Import-Csv | Sort-Object ContainerId -Unique

    $finalReportName = 'StreamClassicMasterContainerList_' + $dateTimeNameSuffix

    $reportsFolderPath = CreateDirectory ('report\' + $finalReportName)

    $fileName = $reportsFolderPath + '\' + $finalReportName + '_'

    #SPLIT the deduped file to multiple CSVs of 10k records each.

    # variable used to advance the number of the row from which the export starts.
    $startrow = 0

    # counter used in names of resulting CSV files
    $counter = 1

    while ($startrow -lt $csvfile.Count) {
        #pick 10k records starting from the $startrow position and export content to a new file.
        $csvfile | Select-Object -skip $startrow -first 10000 | Export-CSV "$fileName$($counter).csv" -NoTypeInformation

        # Increment the number of the rows from which the export starts.
        $startrow += 10000

        # incrementing the $counter variable.
        $counter++

    }

    Write-Host "`n******************************************************************************"
    Write-Host "Reports are available at this location: $reportsFolderPath"
    Write-Host "******************************************************************************`n"

    return $csvfile.Count
}

Function WriteToCsvMasterReport {
    Param(
        $reportData,
        $csvPath,
        $csvName, 
        [ref]$log,
        $isChannel
    )

    $ReportFilePath = $csvPath + '\' + $csvName
    $csvHeaders = 'ContainerId', 'ContainerName', 'Email', 'VideoCount', 'ContainerType';

    # Create file if doesn't exist
    if (!(Test-Path $ReportFilePath -PathType leaf)) {
        New-Item -Path $ReportFilePath -ItemType File | Out-Null;
    }
    $lengthOfHeadersRow = (Get-Content $ReportFilePath | Select-Object -First 1).Length;
    if ($lengthOfHeadersRow -eq 0) {
        #It will be 0, in case of newly created file
        #Put the headers in the file first
        Add-Content -Path $ReportFilePath -Value (($csvHeaders | ForEach-Object { return $_ }) -join ',');
    }
    #Writing content to temp file
    try {
        #all channels with null aad group are company wide channels
        if ($isChannel -eq $true) {
            $response.value | Where-Object { $null -eq $_.aadGroup } | Select-Object @{Name = "ContainerId"; Expression = { $_.id } }, @{Name = "ContainerName"; Expression = { $_.name } }, @{Name = "Email"; Expression = { $_.aadGroup.mail } }, @{Name = "VideoCount"; Expression = { $_.metrics.videos } }, @{Name = "ContainerType"; Expression = { "CompanyWideChannel" } } | Export-Csv -Path $ReportFilePath -encoding utf8 -Append -NoTypeInformation 
        }
        #all groups with non null aad group are m365 groups
        else {
            $response.value | Where-Object { $null -ne $_.aadGroup } | Select-Object @{Name = "ContainerId"; Expression = { $_.id } }, @{Name = "ContainerName"; Expression = { $_.name } }, @{Name = "Email"; Expression = { $_.aadGroup.mail } }, @{Name = "VideoCount"; Expression = { $_.metrics.videos } }, @{Name = "ContainerType"; Expression = { "M365Group" } } | Export-Csv -Path $ReportFilePath -encoding utf8 -Append -NoTypeInformation 
        }
        
    }
    catch {
        $log.Value += "error occurred while writing response to csv file`n"
        if ($Error[0] -and $Error[0].ErrorDetails) {
            $log.Value += "$Error[0].ErrorDetails`n"
        }
        return $null
    }
}

$functions = {

    function Get-ContainerVideoCount {
        param(
            [string]$baseUrl,
            [String]$containerId,
            [String]$containerType,
            [hashtable]$Headers = @{},
            [ref]$log

        )
        try {
            if ($containerType -eq "CompanywideChannel") {
                $url = $baseUrl + 'channels/' + $containerId + "?adminmode=true&api-version=1.4-private"
            }
            elseif ($containerType -eq "M365Group" -or $containerType -eq "StreamOnlyGroup") {
                $url = $baseUrl + 'groups/' + $containerId + "?adminmode=true&api-version=1.4-private"
            } 
            else {
                # User container case  
                return "NA";
            }  
            $response = Invoke-RestMethod -Uri $url -Method Get -Headers $Headers
            return $response.metrics.videos
        }
        catch {
            $log.Value += ((Get-Date).tostring() + " Error Occured for Container Id :  $containerId , ContainerType $containerType  $_`n")
            return $null
        }
    }  
    
    Function WriteToCsv {
        Param(
            $reportData,
            $csvPath,
            $csvName, 
            $cwcFilePath,
            $cwcCsvName,
            [ref]$log
        )

        $ReportFilePath = $csvPath + '\' + $csvName
        $cwcContainerFile = $cwcFilePath + '\' + $cwcCsvName
        $csvHeaders = 'VideoId', 'Name', 'State', 'Description', 'PublishedDate', 'LastViewDate', 'Size (in Bytes)', 'Views', 'Likes', 'ContentType', 'PrivacyMode', 'Creator', 'Owners', 'ContainerId', 'ContainerName', 'ContainerType', 'ContainerEmailId', 'ContainerAadId', 'MigratedDestination', 'ContainerVideosInClassicUI', 'IsEligibleForMigration', 'IsRemigrationNeeded';

        $ActionDelegate = {
            param($video)
            $row = '';
            $row += '"' + $video.id + '",';
            $row += '"' + $video.name.Replace('"', '""') + '",';
            $row += '"' + $video.state + '",';
            if ($video.description) {
                $row += '"' + $video.description.Replace('"', '""') + '",';
            }
            else {
                $row += '"",';
            }
            if ($video.publishedDate -eq "9999-12-31T23:59:59.9999999Z") {
                $row += '"",'
            }
            else {
                if ($video.publishedDate) {
                    $video.publishedDate = Get-Date -Date $video.publishedDate -format "MM/dd/yyyy HH:mm:ss";
                }
                $row += '"' + $video.publishedDate + '",';
            }
            if ($video.lastViewDate) {
                $video.lastViewDate = Get-Date -Date $video.lastViewDate -format "MM/dd/yyyy HH:mm:ss";
            }
            $row += '"' + $video.lastViewDate + '",';
            if ($video.size) {
                $row += '"' + $video.size.tostring() + '",';
            }
            else {
                $row += '"",'
            }
            if ($video.container.containerType -eq 'CompanywideChannel') {
                if (!(Test-Path $cwcContainerFile -PathType leaf)) {
                    New-Item -Path $cwcContainerFile -ItemType File | Out-Null
                }
                $cwcContainerobject = [PSCustomObject]@{
                    'ContainerId' = $video.container.id
                }
                # Adding the CWC containerId to the CSV file
                $cwcContainerobject | Export-Csv -Path $cwcContainerFile -Append -NoTypeInformation
            }
            $row += '"' + $video.viewCount.tostring() + '",';
            $row += '"' + $video.likeCount.tostring() + '",';
            $row += '"' + $video.contentType + '",';
            $row += '"' + $video.privacyMode + '",';
            $row += '"' + $video.creator + '",';
            $row += '"' + $video.owners + '",';
            $row += '"' + $video.container.id + '",';
            $row += '"' + $video.container.name + '",';
            $row += '"' + $video.container.containerType + '",';
            $row += '"' + $video.container.emailId + '",';
            $row += '"' + $video.container.containerAadId + '",';
            $row += '"' + $video.destinationUrl + '",';
            if ($null -ne $video.container.id -and $containerVideoCountMap.ContainsKey($video.container.id)) {
                $row += '"' + $containerVideoCountMap[$video.container.id] + '",'
            }
            else {
                $row += '"NA",'
            }
            if ($video.state -in @("Processing", "Completed") -and $video.publishedDate -ne "9999-12-31T23:59:59.9999999Z" -and ($video.container.id -ne "" -and $video.container.id -ne "00000000-0000-0000-0000-000000000000")) {
                $isEligible = "Yes"
            }
            else {
                $isEligible = "No"
            }
            $row += '"' + $isEligible + '",';
            if ($video.remigrate -eq $true) {
                $remigrationNeeded = "Yes"
            }
            else {
                $remigrationNeeded = "No"
            }
            $row += '"' + $remigrationNeeded + '"'
            return $row
        } 
        # Create file if doesn't exist
        if (!(Test-Path $ReportFilePath -PathType leaf)) {
            New-Item -Path $ReportFilePath -ItemType File | Out-Null;
        }
        $lengthOfHeadersRow = (Get-Content $ReportFilePath | Select-Object -First 1).Length;
        if ($lengthOfHeadersRow -eq 0) {
            #It will be 0, in case of newly created file
            #Put the headers in the file first
            Add-Content -Path $ReportFilePath -Value (($csvHeaders | ForEach-Object { return $_ }) -join ',');
        }
        #Writing content to temp file
        foreach ($video in $reportData) {
            $row = Invoke-Command $ActionDelegate -ArgumentList $video
            $row | Out-File $ReportFilePath -encoding utf8 -Append
        }
    }

    Function ReportDetailsAndWriteToCsv {
        Param(
            $baseUrl,
            $offsetId,
            $csvPath,
            $csvName,
            $token,
            $filter,
            $cwcFilePath,
            $cwcCsvName
        )

        $log = ''
        $status = ''
        $statusCode = ''

        $reportUri = $baseUrl + 'migrationReports?api-version=1.0-odsp-migration'

        if ($filter.Length -ne 0) {
            $reportUri += '&$filter=' + $filter
        }

        if ($offsetId) {
            $reportUri += '&$skiptoken=offsetId:' + $offsetId
        }

        $headers = @{
            Authorization = "Bearer $token"
        }

        $log += ((Get-Date).tostring() + ' ReportDetails URI: ' + $reportUri + "`n")

        try {
            $global:containerVideoCountMap = @{}
            #ReportDetails API call
            $response = Invoke-RestMethod -Uri $reportUri -Method Get -Headers $headers -ContentType "application/json"

            foreach ($video in $response.value) {
                if ($null -ne $video.container.id -and !$containerVideoCountMap.ContainsKey($video.container.id)) {
                    $videoCount = Get-ContainerVideoCount -baseUrl $baseUrl -containerId $video.container.id -containerType $video.container.containerType -Headers $headers -log ([ref]$log)
                    if ($null -ne $videoCount) {
                        $containerVideoCountMap.Add($video.container.id, $videoCount)  
                    }   
                }
                
            }
            
            #Write content returned from ReportDetails API call to a temp file(CSV).
            WriteToCsv $response.value $csvPath $csvName $cwcFilePath $cwcCsvName ([ref]$log)

            $status = "Success"
        }
        catch {
            $log += ((Get-Date).tostring() + ' ' + $_.Exception + "`n")
            
            $log += ((Get-Date).tostring() + ' Error: ' + $Error + "`n")

            #If API or write to CSV fails then status should be written as FAILED in State.csv for this OffsetId.
            $status = "Failed"

            if ($_.Exception.Response.StatusCode.value__ -eq 401) {
                $statusCode = $_.Exception.Response.StatusCode.value__

                Write-Host "========Enter new token and start the script again======="
            }
                
            if ($_.Exception.Response.StatusCode.value__ -eq 400) {
                $statusCode = $_.Exception.Response.StatusCode.value__
                $errorMessage = $Error[0].ErrorDetails
                Write-Host "$errorMessage"
            }

            Write-Host "`nSome error occurred. Check logs for more info.`n"
        }

        return @($status, $log, $statusCode)
    }
}

Function GetGroupDetailsAndWriteToCsv {
    Param(
        $baseUrl,
        $token,
        $skip,
        $top,
        $csvPath,
        $csvName
    )

    $log = ''
    $status = ''
    $statusCode = ''
    $flag = $true

    $url = -join ("$baseUrl", "groups?$", "skip=$skip", '&$', "top=$top", '&$filter=isDefault%20eq%20false&$orderby=metrics%2Fvideos%20desc&', "api-version=1.4-private")
    "`ntrying $url" | Out-File $masterListLogFilePath -Append
    $headers = @{
        Authorization = "Bearer $token"
    }

    $log += ((Get-Date).tostring() + ' Fetch Groups URI: ' + $reportUri + "`n")

    try {
        $response = Invoke-RestMethod -Uri $url -Method Get -Headers $Headers -ContentType "application/json"
        $numberOfEntitiesFetched = $response.value.Length
        if ($numberOfEntitiesFetched -lt $top) { 
            $flag = $false 
        }

        $status = "Success"

        #write fetched details into csv
        if ($numberOfEntitiesFetched -gt 0) {
            WriteToCsvMasterReport -reportData $response.value -csvPath $csvPath -csvName $csvName -log ([ref]$log) -isChannel $false
        }

    }
    catch {
        $status = "Failed"
        $log += ((Get-Date).tostring() + ' ' + $_.Exception + "`n")
        
        $log += ((Get-Date).tostring() + ' Error: ' + $Error + "`n")

        #If API or write to CSV fails then status should be written as FAILED in State.csv for this url call.
        # $status = "Failed"

        if ($_.Exception.Response.StatusCode.value__ -eq 401) {
            $statusCode = $_.Exception.Response.StatusCode.value__

            Write-Host "========Enter new token and start the script again======="
        }
            
        if ($_.Exception.Response.StatusCode.value__ -eq 400) {
            $statusCode = $_.Exception.Response.StatusCode.value__
            $errorMessage = $Error[0].ErrorDetails
            Write-Host "$errorMessage"
        }

        Write-Host "`nSome error occurred. Check logs for more info.`n"
    }

    return @($status, $log, $statusCode, $flag, $url, $numberOfEntitiesFetched)
}

Function GetCWCDetailsAndWriteToCsv {
    Param(
        $baseUrl,
        $token,
        $skip,
        $top,
        $csvPath,
        $csvName
    )

    $log = ''
    $status = ''
    $statusCode = ''
    $flag = $true

    $url = -join ("$baseUrl", "channels?adminmode=true&`$skip=$skip&`$top=$top&`$orderby=metrics%2Ffollows%20desc&`$expand=creator,group&`$filter=isDefault%20eq%20false&api-version=1.4-private")
    "`ntrying $url" | Out-File $masterListLogFilePath -Append
    $headers = @{
        Authorization = "Bearer $token"
    }

    $log += ((Get-Date).tostring() + ' Trying to fetch CompanyWideChannels URI: ' + $reportUri + "`n")

    try {
        $response = Invoke-RestMethod -Uri $url -Method Get -Headers $Headers -ContentType "application/json"
        $numberOfEntitiesFetched = $response.value.Length
        if ($numberOfEntitiesFetched -lt $top) { 
            #do not fetch next batch
            $flag = $false 
        }
        
        $status = "Success"

        #write fetched details into csv
        if ($numberOfEntitiesFetched -gt 0) {
            WriteToCsvMasterReport -reportData $response.value -csvPath $csvPath -csvName $csvName -log ([ref]$log) -isChannel $true
        }
    }
    catch {
        $status = "Failed"
        $log += ((Get-Date).tostring() + ' ' + $_.Exception + "`n")
        
        $log += ((Get-Date).tostring() + ' Error: ' + $Error + "`n")

        #If API or write to CSV fails then status should be written as FAILED in State.csv for this url call.
        # $status = "Failed"

        if ($_.Exception.Response.StatusCode.value__ -eq 401) {
            $statusCode = $_.Exception.Response.StatusCode.value__

            Write-Host "========Enter new token and start the script again======="
        }
            
        if ($_.Exception.Response.StatusCode.value__ -eq 400) {
            $statusCode = $_.Exception.Response.StatusCode.value__
            $errorMessage = $Error[0].ErrorDetails
            Write-Host "$errorMessage"
        }

        Write-Host "`nSome error occurred. Check logs for more info.`n"
    }

    return @($status, $log, $statusCode, $flag, $url, $numberOfEntitiesFetched)
}

Function AddHeadersToStateFile {
    Param(
        $stateFilePath
    )

    $csvHeaders = 'OffsetId', 'RetryCount', 'Status';

    $lengthOfHeadersRow = (Get-Content $stateFilePath | Select-Object -First 1).Length;

    if ($lengthOfHeadersRow -eq 0) {
        # It will be 0, in case of newly created file
        # Put the headers in the file first
        Add-Content -Path $stateFilePath -Value (($csvHeaders | ForEach-Object { return $_ }) -join ',');
    }
}

Function AddHeadersToMasterReportStateFile {
    Param(
        $masterReportStateFilePath
    )

    $csvHeaders = 'Url', 'Top_Skip', 'RetryCount', 'Status';

    $lengthOfHeadersRow = (Get-Content $masterReportStateFilePath | Select-Object -First 1).Length;

    if ($lengthOfHeadersRow -eq 0) {
        # It will be 0, in case of newly created file
        # Put the headers in the file first
        Add-Content -Path $masterReportStateFilePath -Value (($csvHeaders | ForEach-Object { return $_ }) -join ',');
    }
}

Function CreateFile {
    Param(
        $filePath
    )

    if ((Test-Path $filePath -PathType leaf)) {
        Remove-Item ($filePath)
    }

    New-Item -Path $filePath -ItemType File | Out-Null

    Write-Host "Created $filePath file."
}

Function RetryFailedOffsetIds {
    Param(
        $stateFilePath,
        $csvPath,
        $csvNamePrefix,
        $baseUrl,
        $parallelCalls,
        $cwcFilePath
    )

    $i = 0

    ((Get-Date).tostring() + " Retrying failed offset Ids, if any.`n") | Out-File $logFilePath -Append

    $file = Import-csv $stateFilePath

    #Fetching those records from state.csv which are FAILED and have not been retried 5 times yet.
    $file = @($file | Where-Object { ($_.Status -eq "Failed") -and ([int]$_.RetryCount -lt 5) })

    $NumberOfFailedEntries = $file.Length

    while ($i -lt $NumberOfFailedEntries) {
        $j = 1

        #Retrying the failed Ids parallely in the batch of degreeOfParallelism.
        while (($j -le $parallelCalls) -and ($i -lt $NumberOfFailedEntries)) {
            $csvName = $timeStamp + '_' + $j.ToString() + '.csv'
            $cwcCsvName = 'cwcContainerId' + '_' + $timeStamp + '_' + $i.ToString() + '.csv' 
            Start-Job -Name $file[$i].OffsetId -InitializationScript $functions -ScriptBlock {

                param($baseUrl, $offsetId, $csvPath, $csvName, $token, $cwcFilePath, $cwcCsvName)

                #Fetches the content and write into individual temporary files.
                ReportDetailsAndWriteToCsv -baseUrl $baseUrl -offsetId $offsetId -csvPath $csvPath -csvName $csvName -token $token -cwcFilePath $cwcFilePath -cwcCsvName $cwcCsvName

            } -ArgumentList $baseUrl, $file[$i].OffsetId, $csvPath, $csvName, $token , $cwcFilePath, $cwcCsvName | Out-Null

            $j++
            $i++
        }

        Get-Job | Wait-Job | Out-Null

        $jobs = Get-job | Sort-Object -Property { [int]$_.Name }

        foreach ($job in $jobs) {
            $result = $job | Receive-job

            #If Job failes to complete then status should be 'Failed' for it to be retried on next run.
            #If Job succeeds, then status will be returned from the execution code.
            if ($job.JobStateInfo.State -ne 'Completed') {
                $result[0] = 'Failed'
            }

            $result[1] | Out-File $logFilePath -Append

            #UPDATE the state.csv file for previously failed OffsetId.

            $stateFile = $null

            $stateFile = Import-Csv -Path $stateFilePath

            if (@($stateFile).Length -gt 0) {
                $stateFile | ForEach-Object {
                    #update the status and retryCount for the matching OffsetId.
                    if ($job.Name -eq $_.OffsetId) {
                        if ($_.Status -eq "Failed") {
                            $_.Status = $result[0]
                            $_.RetryCount = ([int]$_.RetryCount + 1)
                        }
                    }
                }

                #Export the locally written data back to state.csv.
                $stateFile | Export-Csv -Path $stateFilePath -NoTypeInformation
            }

            if ($result[2] -eq 401) {
                exit
            }
        }

        Get-Job | Remove-Job
        Merge-CSVFiles -rootFolder $cwcFilePath
    }
}

Function RetryFailedUrls {
    Param(
        $masterListStateFilePath,
        $masterListCsvPath
    )

    $timeStamp = Get-Date -Format FileDateTime
    ((Get-Date).tostring() + " Retrying failed fetch M365Group/CWC urls, if any.`n") | Out-File $masterListLogFilePath -Append

    $stateFile = $null
    $stateFile = Import-csv $masterListStateFilePath

    #Fetching those records from state.csv which are FAILED and have not been retried 5 times yet.
    $file = @($stateFile | Where-Object { ($_.Status -eq "Failed") })

    $hashmap = @{}
    foreach ($row in $file) {
        $hashmap[$row.Url] = $row.Status
    }

    $NumberOfFailedEntries = $file.Length

    $counter = 0
    $i = 0
    while ($counter -lt $NumberOfFailedEntries) {
        if ($counter % 5 -eq 0) {
            $i++
        }
        $csvName = $timeStamp + '_' + $i.ToString() + '.csv'
        $retryUrl = $file[$counter].Url
        if ($retryUrl -match 'skip=(\d+)') {
            $skip = $matches[1]
        }
        if ($retryUrl -like "*api/channels*") {
            try {
            
                $result = GetCWCDetailsAndWriteToCsv -baseUrl $baseUrl -token $token -skip $skip -top $pageSize -csvPath $masterListCsvPath -csvName $csvName
                $status = $result[0]
                $log = $result[1]
                $statusCode = $result[2]
                $url = $result[4]
    
                #write to log file
                $log | Out-File $masterListLogFilePath -Append

                $hashmap[$url] = $status    

                #If API call failed with 401, then there is no need to continue and exit immediately as token is not valid anymore. Re-start the script with Resume flag set as true.
                if ($statusCode -eq 401) {
                    exit
                }
            }
            catch {
                $errorMessage = $_.Exception.Message
                Write-Host "`nSome error occurred. Check logs for more info. $errorMessage`n"
                $errorMessage | Out-File $masterListLogFilePath -Append
            }
    
        }
        else {
            try {
                $result = GetGroupDetailsAndWriteToCsv -baseUrl $baseUrl -token $token -skip $skip -top $pageSize -csvPath $masterListCsvPath -csvName $csvName
                $status = $result[0]
                $log = $result[1]
                $statusCode = $result[2]
                $url = $result[4]
    
                #write to log file
                $log | Out-File $masterListLogFilePath -Append

                $hashmap[$url] = $status 

                #If API call failed with 401, then there is no need to continue and exit immediately as token is not valid anymore. Re-start the script with Resume flag set as true. We wrote this call as well in state.csv so it will be picked by RetryFailedUrls function later.
    
                #Q - Since the call failed due to token expiry, should we avoid writing this call in state.csv and rather resume next call from here by computing this call from its previous call's url?
                if ($statusCode -eq 401) {
                    exit
                }
            }
            catch {
                $errorMessage = $_.Exception.Message
                Write-Host "`nSome error occurred. Check logs for more info. $errorMessage`n"
                $errorMessage | Out-File $masterListLogFilePath -Append
            }
        }
        
        $counter++
    
    }            
    #update state.csv for all retried entries with new statuses using hashmap
    foreach ( $row in $stateFile ) {
        if ($hashmap.ContainsKey($row.Url)) {
            $row.Status = $hashmap[$row.Url]
        }
    }

    #Export the locally written data back to state.csv.
    $stateFile | Export-Csv -Path $masterListStateFilePath -NoTypeInformation
}

function Merge-CSVFiles {
    param (
        [Parameter(Mandatory = $true)]
        [string]$rootFolder
    )

    $outputFilePath = $rootFolder + "\CWCContainerIds.csv"
    # Get all CSV files in the root folder
    if (-not (Test-Path $rootFolder)) {
        return
    }
    $csvFiles = Get-ChildItem -Path $rootFolder -Filter "*.csv"

    # Create an empty array to store the merged data
    $mergedData = @()

    # Iterate through each CSV file
    foreach ($csvFile in $csvFiles) {
        # Import the CSV file and append its data to the merged data array
        $data = Import-Csv -Path $csvFile.FullName | Select-Object -Property * -Unique
        $mergedData += $data
    }

    # Export the merged data to a new CSV file
    $mergedData | Export-Csv -Path $outputFilePath -NoTypeInformation

    #cleanup unused files
    $excludedFiles = @("CWCContainerIds.csv")
    foreach ($csvFile in $csvFiles) {
        if ($excludedFiles -notcontains $csvFile.Name) {
            Remove-Item -Path $csvFile.FullName -Force
        }
    }
}
    
Function GenerateFilterString {
    Param([ref]$filter)

    if ($PublishedDateLe.Length -ne 0) {
        $filter.Value = 'publishedDate le ' + $PublishedDateLe
    }

    if ($PublishedDateGe.Length -ne 0) {
        if ($filter.Value.Length -ne 0) {
            $filter.Value += ' and '
        }

        $filter.Value += 'publishedDate ge ' + $PublishedDateGe
    }
}

Function StartReportGeneration {

    $startDateTime = Get-Date
    Get-Job | Remove-Job

    $stateDirectoryPath = CreateDirectory 'state'

    Write-Host "------------------------------------------------------------`n"
    $reportFolderName = 'StreamClassicVideoReport'
    $reportPath = CreateDirectory $reportFolderName

    $stateFilePath = $stateDirectoryPath + '\state.csv'
    $lastRunFilePath = $stateDirectoryPath + '\lastRunFolder.txt'
    $cwcFilePath = CreateDirectory 'CWC'
    $offsetIdsList = @()

    # Initializing the report output directory filename
    $timeStamp = Get-Date -Format FileDateTime
    $csvPathFromOutDir = $reportFolderName + '\' + $timeStamp

    # If execution is being started from scratch, create the required files.
    if ($ResumeLastRun -eq 'false') {
        CreateFile $stateFilePath

        AddHeadersToStateFile $stateFilePath

        CreateFile $logFilePath
      
        CreateFile $lastRunFilePath
        if (Test-Path $cwcFilePath -PathType leaf) {
            Remove-Item -Path $cwcFilePath -Recurse | Out-Null
        }
        $cwcFilePath = CreateDirectory 'CWC'
        #In case of fresh run, create new output directory and save that path in the lastRunFolder file
        $csvPath = CreateDirectory $csvPathFromOutDir
        Set-Content -Path $lastRunFilePath -Value $csvPath

    }
    else { 
        #Checking for state.csv for resumption.
        Write-Host "`nSearching for last saved state..."

        if (!(Test-Path $logFilePath -PathType leaf)) {
            CreateFile $logFilePath
        }

        if (!(Test-Path $stateFilePath -PathType leaf)) {
            Write-Host "No saved state found. Preparing to fetch from scratch...`n"
            CreateFile $stateFilePath

            AddHeadersToStateFile $stateFilePath
        }
        else {
            #resume state
            $stateData = Import-Csv $stateFilePath
            $offsetIdsList = $stateData | Select-Object -ExpandProperty OffsetId
        }

        if (!(Test-Path $lastRunFilePath -PathType leaf)) {
            #For backward compatibility, as script executions prior to this version will not have the new lastRunFolder file
            #so creating new output directory and save that path in the lastRunFolder file
            CreateFile $lastRunFilePath
            $csvPath = CreateDirectory $csvPathFromOutDir
            Set-Content -Path $lastRunFilePath -Value $csvPath
        }
        else {
            #Fetch the LastRunFolder incase of resumed script run
            $csvPath = Get-Content $lastRunFilePath
        }
    }

    ((Get-Date).tostring() + " ------START------`nScript version: $scriptVersion`n") | Out-File $logFilePath -Append

    $baseUrl = GetBaseUrl
    $csvNamePrefix = ""

    Write-Host "------------------------------------------------------------`n"

    Do {
        $offsetId = $offsetIdsList | Select-Object -Last 1

        $orchestrationResponse = ReportOrchestration $baseUrl $offsetId
        $orchestrationResponse | Out-File $logFilePath -Append

        $offsetIdsList = $orchestrationResponse.offsetIds
        $parallelThreads = $orchestrationResponse.degreeOfParallelism
        #Check if no offset ids were returned for the first orchestration call in the script execution (whether resuming last run or starting new run)
        if ($firstOrchestration -and ($offsetIdsList.Length -eq 0)) {
            if ($ResumeLastRun -eq 'true') {
                Write-Host "`nNo new videos found" -ForegroundColor Yellow | Out-File $logFilePath -Append
                break #break from the loop as we need to retry for failed offsets, if any, incase it was a resume run
            }
            else {
                #remove report folder created before
                Remove-Item ($csvPathFromOutDir) -Confirm:$false -Force -Recurse
                Write-Host "`nNo videos found" -ForegroundColor Yellow | Out-File $logFilePath -Append
                exit
            }
        }   
        $firstOrchestration = $false;

        Write-Host "`nExtracting data..."

        $i = 1

        #Fetching ReportDetails in parallel.
        #Number of jobs = Number of Ids in offsetIdsList(fetched from Orchestration API) OR degree of parallelism.
        foreach ($Id in $offsetIdsList) {
            $csvName = $timeStamp + '_' + $i.ToString() + '.csv'
            $cwcCsvName = 'cwcContainerId' + '_' + $timeStamp + '_' + $i.ToString() + '.csv'

            Start-Job -Name $Id -InitializationScript $functions -ScriptBlock {

                param($baseUrl, $offsetId, $csvPath, $csvName, $token, $filter, $cwcFilePath, $cwcCsvName)

                #Fetches the content and write into individual temporary files.
                ReportDetailsAndWriteToCsv -baseUrl $baseUrl -offsetId $offsetId -csvPath $csvPath -csvName $csvName -token $token -filter $filter -cwcFilePath $cwcFilePath -cwcCsvName $cwcCsvName

            } -ArgumentList $baseUrl, $Id, $csvPath, $csvName, $token, $filter, $cwcFilePath, $cwcCsvName | Out-Null

            $i++
        }

        Get-Job | Wait-Job | Out-Null

        $jobs = Get-job | Sort-Object -Property { [int]$_.Name }

        foreach ($j in $jobs) {
            $result = $j | Receive-job

            #If Job failes to complete then status should be 'Failed' for it to be retried on next run.
            #If Job succeeds, then status will be returned from the execution code.
            if ($j.JobStateInfo.State -ne 'Completed') {
                $result[0] = 'Failed'
            }

            #Write logs, returned from the job, into the file.
            $result[1] | Out-File $logFilePath -Append

            # Write to state.csv
            ('"' + $j.Name.ToString() + '","0","' + $result[0] + '"') | Out-File $stateFilePath -encoding ASCII -Append

            #If API call failed with 401, then there is no need to continue and exit immediately as token is not valid anymore.
            if ($result[2] -eq 401) {
                exit
            }
        }

        Get-Job | Remove-Job
        Merge-CSVFiles -rootFolder $cwcFilePath

    }while ($offsetIdsList)

    #Once all OffsetIds have been processed, scan state.csv to retry failed ones.
    RetryFailedOffsetIds $stateFilePath $csvPath $csvNamePrefix $baseUrl $parallelThreads $cwcFilePath

    Write-Host "`n******************************************************************************"
    Write-Host "Reports are available at this location: $csvPath" -ForegroundColor Green
    Write-Host "******************************************************************************`n"

    $endDateTime = Get-Date

    Write-Host "Time elapsed: "($endDateTime - $startDateTime)"`n"
    ((Get-Date).tostring() + " Time elapsed: " + ($endDateTime - $startDateTime) + "`n") | Out-File $logFilePath -Append

    if (Test-Path $stateFilePath -PathType leaf) {
        #check for failed offsets in state even after retrying
        $stateCsvData = Import-csv $stateFilePath
        $failedOffsets = @($stateCsvData | Where-Object { ($_.Status -eq "Failed") })
    
        if ($failedOffsets.Length -gt 0) {
            # saving the failed offsets in logs
            $failedOffsets | Out-File $logFilePath -Append
            Write-Host "Failed to fetch details for some videos." -ForegroundColor Yellow
            Write-Host "Please re-run the script. If the error persists, kindly reach out to Customer Support and share the log file: $logFilePath"
        }
    }
}

Function StartMasterListReportGeneration {

    $startDateTime = Get-Date
    Get-Job | Remove-Job

    Write-Host "------------------------------------------------------------`n"
    $reportFolderName = 'StreamClassicMasterContainerList'
    $reportPath = CreateDirectory $reportFolderName

    # Initializing the report output directory filename
    $timeStamp = Get-Date -Format FileDateTime
    $masterListCsvPathFromOutDir = $reportFolderName + '\' + $timeStamp
    $lastRunRow = $null
    Write-Host "Resume mode: $ResumeLastRun`n"

    # If execution is being started from scratch, create the required files.
    #$ResumeLastRun = 'false'
    if ($ResumeLastRun -eq 'false') {
        CreateFile $masterListStateFilePath

        AddHeadersToMasterReportStateFile $masterListStateFilePath

        CreateFile $masterListLogFilePath
      
        CreateFile $masterListLastRunFilePath

        #In case of fresh run, create new output directory and save that path in the lastRunFolder file
        $masterListCsvPath = CreateDirectory $masterListCsvPathFromOutDir
        Set-Content -Path $masterListLastRunFilePath -Value $masterListCsvPath

    }
    else { 
        #Checking for state.csv for resumption.
        Write-Host "`nSearching for last saved state..."

        if (!(Test-Path $masterListLogFilePath -PathType leaf)) {
            CreateFile $masterListLogFilePath
        }

        if (!(Test-Path $masterListStateFilePath -PathType leaf)) {
            Write-Host "No saved state found. Preparing to fetch from scratch...`n"
            CreateFile $masterListStateFilePath

            AddHeadersToMasterReportStateFile $masterListStateFilePath
        }
        else {
            #resume state
            $masterListStateData = Import-Csv $masterListStateFilePath
            if ($null -eq $masterListStateData) {
                Write-Host "Empty saved state found. Preparing to fetch from scratch...`n"
                CreateFile $masterListStateFilePath

                AddHeadersToMasterReportStateFile $masterListStateFilePath
            }
            else {
                $lastRunRow = $masterListStateData[-1]
            }
        }

        if (!(Test-Path $masterListLastRunFilePath -PathType leaf)) {
            #For backward compatibility, as script executions prior to this version will not have the new lastRunFolder file
            #so creating new output directory and save that path in the lastRunFolder file
            CreateFile $masterListLastRunFilePath
            $masterListCsvPath = CreateDirectory $masterListCsvPathFromOutDir
            Set-Content -Path $masterListLastRunFilePath -Value $masterListCsvPath
        }
        else {
            #Fetch the LastRunFolder incase of resumed script run
            $masterListCsvPath = Get-Content $masterListLastRunFilePath
        }
    }

    ((Get-Date).tostring() + " ------START------`nScript version: $scriptVersion`n") | Out-File $masterListLogFilePath -Append

    $baseUrl = GetBaseUrl

    Write-Host "Fetching CompanyWideChannels...`n"
    "Starting fetch for CompanyWideChannels" | Out-File $masterListLogFilePath -Append

    $i = 0
    $counter = 0
    $flag = $true
    $skip = 0
    $pageSize = 100

    #if last row was a fetch for channels, we need to resume fetching channels from last successful call and then later start fetching groups from start. we maintain the below flag for this purpose of ensuring that for fetching groups we start from beginning.
    $resumeFetchForCWC = $false
    Do {
        #in case of resume, we pick the last run url and skip value and start from next skip value ( = skip + pageSize)
        #we also check if the last api call was for CWC or M365Group
        #counter = 0 check below because the below block should run only once
        if ($ResumeLastRun -eq 'true' -and $null -ne $lastRunRow -and $counter -eq 0) {
            $lastUrl = $lastRunRow.Url
            if ($lastUrl -like "*api/channels*") {
                #if last time top_skip = 24_200 => new skip = 24 + 200 (oldtop + oldskip) 
                $skip = [int](($lastRunRow.Top_Skip -split '_')[0]) + [int](($lastRunRow.Top_Skip -split '_')[1])
                $resumeFetchForCWC = $true
            }
            else {
                #if last api call was not for CWC, it means all CWCs have been fetched, so break from this loop
                break
            }
        }
        
        if ($counter % 5 -eq 0) {
            $i++
        }
        $csvName = $timeStamp + '_' + $i.ToString() + '.csv'
        
        try {
            $result = GetCWCDetailsAndWriteToCsv -baseUrl $baseUrl -token $token -skip $skip -top $pageSize -csvPath $masterListCsvPath -csvName $csvName
            $status = $result[0]
            $log = $result[1]
            $statusCode = $result[2]
            $flag = $result[3]
            $url = $result[4]
            $numberOfEntitiesFetched = $result[5]

            #write to log file
            $log | Out-File $masterListLogFilePath -Append

            # Write to state.csv
            $stateobject = [PSCustomObject]@{
                'Url'        = $url
                'Top_Skip'   = "$numberOfEntitiesFetched" + "_" + "$skip"
                'RetryCount' = '0'
                'Status'     = $status
            }

            $stateobject | Export-Csv $masterListStateFilePath -encoding ASCII -Append -NoTypeInformation

            #If API call failed with 401, then there is no need to continue as token is no more valid, so exit immediately. Re-start the script with Resume flag set as true.
            if ($statusCode -eq 401) {
                $flag = $false
                exit
            }
        }
        catch {
            $errorMessage = $_.Exception.Message
            Write-Host "`nSome error occurred. Check logs for more info. $errorMessage`n"
            $errorMessage | Out-File $masterListLogFilePath -Append
        }
 
        $skip += $numberOfEntitiesFetched
        $counter++
    } while ($flag)   
    
    Write-Host "Fetching M365Groups...`n"
    "Starting fetch for M365Groups" | Out-File $masterListLogFilePath -Append

    $i++
    $counter = 0
    $flag = $true
    $skip = 0
    Do {
        #counter = 0 check below because the below block should run only once
        if ($ResumeLastRun -eq 'true' -and $null -ne $lastRunRow -and $resumeFetchForCWC -eq $false -and $counter -eq 0) {
            $lastUrl = $lastRunRow.Url
            if ($lastUrl -like "*api/groups*") {
                #if last time top_skip = 24_200 => new skip = 24 + 200 (oldtop + oldskip) 
                $skip = [int](($lastRunRow.Top_Skip -split '_')[0]) + [int](($lastRunRow.Top_Skip -split '_')[1])
            }
            else {
                #if last api call was not for groups, it means all CWCs and M365Groups have been fetched, so break from this loop and retry failed calls if any
                break
            }
        }
        #using counter variable for file naming - for every 5 runs we make a new file, so each intermediate .csv will have 500 rows at max
        if ($counter % 5 -eq 0) {
            $i++
        }
        $csvName = $timeStamp + '_' + $i.ToString() + '.csv'
        
        try {
            $result = GetGroupDetailsAndWriteToCsv -baseUrl $baseUrl -token $token -skip $skip -top $pageSize -csvPath $masterListCsvPath -csvName $csvName
            $status = $result[0]
            $log = $result[1]
            $statusCode = $result[2]
            $flag = $result[3]
            $url = $result[4]
            $numberOfEntitiesFetched = $result[5]

            #write to log file
            $log | Out-File $masterListLogFilePath -Append

            # Write to state.csv
            $stateobject = [PSCustomObject]@{
                'Url'        = $url
                'Top_Skip'   = "$numberOfEntitiesFetched" + "_" + "$skip"
                'RetryCount' = '0'
                'Status'     = $status
            }
            $stateobject | Export-Csv $masterListStateFilePath -encoding ASCII -Append -NoTypeInformation

            #If API call failed with 401, then there is no need to continue and exit immediately as token is not valid anymore. Re-start the script with Resume flag set as true. We wrote this call as well in state.csv so it will be picked by RetryFailedUrls function later.
            if ($statusCode -eq 401) {
                $flag = $false
                exit
            }
        }
        catch {
            $errorMessage = $_.Exception.Message
            Write-Host "`nSome error occurred. Check logs for more info. $errorMessage`n"
            $errorMessage | Out-File $masterListLogFilePath -Append
        }
 
        $skip += $numberOfEntitiesFetched
        $counter++
    } while ($flag)


    #Once all OffsetIds have been processed, scan state.csv to retry failed ones.
    RetryFailedUrls -masterListStateFilePath $masterListStateFilePath -masterListCsvPath $masterListCsvPath

    #merge all csv files and export at one location with each file containing 10,000 records
    $finalContainerCount = DedupeMasterList -csvPath $masterListCsvPath -dateTimeNameSuffix $timeStamp

    #log the final file count
    "$finalContainerCount containers fetched" | Out-File $masterListLogFilePath -Append

    Write-Host "`n******************************************************************************"
    Write-Host "Total number of containers fetched: $finalContainerCount" -ForegroundColor Green
    Write-Host "Reports are available at this location: $masterListCsvPath" -ForegroundColor Green
    Write-Host "******************************************************************************`n"

    $endDateTime = Get-Date

    Write-Host "Time elapsed: "($endDateTime - $startDateTime)"`n"
    ((Get-Date).tostring() + " Time elapsed: " + ($endDateTime - $startDateTime) + "`n") | Out-File $masterListLogFilePath -Append

    if (Test-Path $masterListStateFilePath -PathType leaf) {
        #check for failed offsets in state even after retrying
        $masterListStateCsvData = Import-csv $masterListStateFilePath
        $failedUrls = @($masterListStateCsvData | Where-Object { ($_.Status -eq "Failed") })
    
        if ($failedUrls.Length -gt 0) {
            # saving the failed offsets in logs
            $failedUrls | Out-File $masterListLogFilePath -Append
            Write-Host "Failed to fetch details of some M365Groups/CompanyWideChannels" -ForegroundColor Yellow
            Write-Host "Please re-run the script. If the error persists, kindly reach out to Customer Support and share the log file: $masterListLogFilePath"
        }
    }
}
Function GetM365GroupSharepointDocumentUrl {
    Param(
        $M365GroupEmailId
    )

    try {
        # Get sharepoint destination url of default site lib, for the given M365 group email id
        $output = Get-UnifiedGroup -Identity $M365GroupEmailId | Select SharepointDocumentsUrl
        return $output
    }
    catch {
        #Log error.
        ((Get-Date).tostring() + ' ' + $Error[0].ErrorDetails + "`n") | Out-File $logFilePath -Append
        ((Get-Date).tostring() + ' ' + $Error[0].Exception + "`n") | Out-File $logFilePath -Append
    }
}

Function CreateMigrationDestinationPathMapping {
    Write-Host "Started process for creating Auto-Mapping of M365 group containers to Migration destination path"
    "Started process for creating Auto-Mapping of M365 group containers to Migration destination path" | Out-File $logFilePath -Append
    
    $rowHashSet = @{}
    if ($MigrationDestinationCsvFilePath -eq $null -Or $MigrationDestinationCsvFilePath -eq $empty) {
        Write-Host "MigrationDestinationCsvFilePath is required. Currently its null or empty"
        "MigrationDestinationCsvFilePath is required. Currently its null or empty" | Out-File $logFilePath -Append
        exit
    }
    else {
        Write-Host "Reading csv file:" + $MigrationDestinationCsvFilePath
        "Reading csv file:" + $MigrationDestinationCsvFilePath | Out-File $logFilePath -Append
    }
    
    # Read migration destination csv file
    # Try to read the migration destination csv file with the default delimiter (',')
    $csvFile = Import-Csv -Path $MigrationDestinationCsvFilePath 
    $first_row_elements = $csvFile[0] -split ';'
    # Could not read file at MigrationDestinationCsvFilePath with default delimiter (',') if number of columns is not equal to 6. Try reading with ';' delimiter and -useCulture
    if ($first_row_elements.Count -ne 6) { 
        $csvFile = Import-Csv -Path $MigrationDestinationCsvFilePath  -Delimiter ';'
        $first_row_elements = $csvFile[0] -split ';'
        if ($first_row_elements.Count -ne 6) {
            $csvFile = Import-Csv -Path $MigrationDestinationCsvFilePath -useCulture
            $first_row_elements = $csvFile[0] -split ';'
            if ($first_row_elements.Count -ne 6) {
                Write-Host "This file cannot be parsed. Invalid number of columns. Please check the file format, it doesn't seem to have the correct schema"
                "This file cannot be parsed. Invalid number of columns. Please check the file format, it doesn't seem to have the correct schema" | Out-File $logFilePath -Append
                exit
            }
        }
    }

    if (@($csvFile).Length -gt 0) {
        $csvFile | ForEach-Object {
            #extract ContainerId and ContainerEmailId
            #Pick the row element by its column index
            $row_elements = $_ -split ';'
            if (("" -eq ($row_elements[2] -split '=')[1] -Or "null" -eq ($row_elements[2] -split '=')[1]) -And ($row_elements[1] -split '=')[1].EndsWith('M365Group')) {
                $row = '';
                $row += ($row_elements[0] -split '=')[1] + ',';
                $row += ($row_elements[1] -split '=')[1];
                $row += ',#DestinationPathPlaceHolder#,'
                $row += ($row_elements[3] -split '=')[1] + ',"';
                $row += ($row_elements[4] -split '=')[1] + '",';
                $row += ($row_elements[5] -split '=')[1].Replace('}', '');
                
                $sourcePathFragments = ($row_elements[1] -split '=')[1].Split("|")
                if ($sourcePathFragments.Count -eq 3) {
                    $groupType = $sourcePathFragments[2]
                    $groupEmail = $sourcePathFragments[0]
                
                    # If container type is M365Group, extract the same
                    if ($groupType -eq 'M365Group' -And !$rowHashSet.ContainsKey($row)) {
                        $rowHashSet.Add($row, $groupEmail)
                    }
                }
                else {
                    "Could not parse source path (for reference only). Incorrect value:" + ($row_elements[1] -split '=')[1] | Out-File $logFilePath -Append
                } 
            }
        }
    }
    else {
        Write-Host "Could not read file at MigrationDestinationCsvFilePath. Please check path"
        "Could not read file at MigrationDestinationCsvFilePath. Please check path" | Out-File $logFilePath -Append
        exit
    }

    #Pick the headers for the output file from the input file headers. 
    $mappingFilePath = $OutDir + '\MigrationDestinationMappingForM365GroupContainers_' + $randomGuidForOutFiles.Guid + '.csv'
    $headers = Get-Content -Path $MigrationDestinationCsvFilePath -TotalCount 1
    $headers | Out-File $mappingFilePath -encoding utf8 -Append
    $atleastOneRowOutputIsDone = $false

    Foreach ($row in $rowHashSet.Keys) {
        # Get Mapping of Default site lib path
        $emaildId = $rowHashSet[$row]
        $DocumentLibrary = GetM365GroupSharepointDocumentUrl $emaildId
        $destinationPath = $DocumentLibrary.SharePointDocumentsUrl

        if ($null -eq $destinationPath -or $destinationPath.ToString().Trim(' ') -eq "") {
            $logVal = "Failed to retrive or No SharePointDocumentsUrl discovered for row:" + $row.ToString() 
            $logVal += " .Please check if sharepoint site exists for this group. Please check log, if there is an error. If no error is thrown, SharepointDocumentUrl doesn't exist for your site."
            $logVal += " You can manually assign Migration destination path for this case."
            $logVal | Out-File $logFilePath -Append 
        }
        else {
            $updatedRow = $row.ToString() -replace '#DestinationPathPlaceHolder#', $destinationPath
        
            # Push mapping data out to csv file
            $updatedRow | Out-File $mappingFilePath -encoding utf8 -Append
            $atleastOneRowOutputIsDone = $true
        }
    }

    if ($atleastOneRowOutputIsDone -eq $true) {
        Write-Host "Done. Exiting auto-mapping. A new csv with destination path mapping for M365Group containers, has been created in the OutDir. Please use this to upload destination path in migration tool. FilePath:" + $mappingFilePath
        "Done. Exiting auto-mapping. A new csv with destination path mapping for M365Group containers, has been created in the OutDir. Please use this to upload destination path in migration tool. FilePath:" + $mappingFilePath | Out-File $logFilePath -Append 
    }
    else {
        Write-Host "Done. Empty output file generated. Either no M365Group container with unassigned path was detected or there was error in fetching path. Please check logs. LogFilePath:" $logFilePath
    }
}

function Test-IsGuid {
    param([string]$guid, $logFilepath)
    $regex = [regex]::new('^[A-Fa-f0-9]{8}-([A-Fa-f0-9]{4}-){3}[A-Fa-f0-9]{12}$')
    $isGuid = $regex.IsMatch($guid)
    if ($isGuid -eq $false) {
        ((Get-Date).tostring() + " Invalid CWC Id $guid") | Out-File -FilePath $logFilepath -Append
    }
    return $isGuid
}

Function FetchCreatorDetails {
    Param(
        $baseUrl,
        $containerIds,
        $outputFilePath,
        $logFilepath,
        $stateFilepath
    )
    $cwcCreatorDetailsUri = $baseUrl + 'migrationReports/cwcdetails?api-version=1.0-odsp-migration'

    $headers = @{
        Authorization = "Bearer $token"
    }
    $requestBody = @{
        ChannelIds = @()
    }     

    $expandedContainerIds = $containerIds | Select-object -ExpandProperty 'containerId'
    $requestBody.ChannelIds += $expandedContainerIds
    $apiRequestBody = $requestBody | ConvertTo-Json -Depth 10

    try {
        ((Get-Date).tostring() + " Calling the creator Uri $cwcCreatorDetailsUri with request body $apiRequestBody ") | Out-File -FilePath $logFilepath -Append
        $response = Invoke-RestMethod -Uri $cwcCreatorDetailsUri -Method Post -Body $apiRequestBody -Headers $headers -ContentType "application/json"

        foreach ($value in $response.value) {
            if ([string]::IsNullOrEmpty($value.creatorEmailId)) {
                $status = 'failed'
            }
            else {
                $status = 'success'
            }
            $creatorobj = [PSCustomObject]@{
                'ContainerId'    = $value.Id
                'Name'           = $value.Name
                'Description'    = $value.Description
                'CreatorEmailId' = $value.CreatorEmailId
            }
            $stateObj = [PSCustomObject]@{
                'ContainerId' = $value.Id
                'status'      = $status
            }
            ((Get-Date).tostring() + " channel creator output CWC Id $value.Id") | Out-File -FilePath $logFilepath -Append
            $creatorobj | Export-Csv -Path $outputFilePath -Append -NoTypeInformation
            $stateObj | Export-Csv -Path $stateFilePath -Append -NoTypeInformation
        }
    }
    catch {
        $errorMessage = $_.Exception.Message
        $headers = $_.Exception.Response.Headers
        if ($headers.Contains('x-ms-request-id')) {
            $msRequestId = $headers.GetValues('x-ms-request-id')[0]
        }
        else {
            $msRequestId = ''
        }
        
        ((Get-Date).tostring() + " Failed to fetch Creator email ERROR Message $errorMessage having requestId $msRequestId") | Out-File -FilePath $logFilepath -Append

        if ($_.Exception.Response.StatusCode.value__ -eq 401) {
            $statusCode = $_.Exception.Response.StatusCode.value__

            Write-Host "========Enter new token and start the script again======="
        }
            
        if ($_.Exception.Response.StatusCode.value__ -eq 400) {
            $statusCode = $_.Exception.Response.StatusCode.value__
            Write-Host "$errorMessage"
        }

        Write-Host "`nSome error occurred. Check logs for more info.`n"
    }
}

Function GetCWCCreatorDetails {
    Param(
        $cwcOutputDirName,
        $logFilepath,
        $stateFilepath
    )
    
    # If execution is being started from scratch, create the create the new state, log and last run files .
    if ($ResumeLastRun -eq 'false') {
        $csvHeaders = 'ContainerId', 'Status';
        CreateFile $stateFilePath
        Add-Content -Path $stateFilePath -Value (($csvHeaders | ForEach-Object { return $_ }) -join ',');

        CreateFile $logFilePath
    }
    else {
        #Checking for state.csv for resumption.
        Write-Host "`nSearching for last saved state..."

        if (!(Test-Path $logFilePath -PathType leaf)) {
            CreateFile $logFilePath
        }

        if (!(Test-Path $stateFilePath -PathType leaf)) {
            Write-Host "No saved state found. Preparing to fetch from scratch...`n"
            $csvHeaders = 'ContainerId', 'Status';
            CreateFile $stateFilePath
            Add-Content -Path $stateFilePath -Value (($csvHeaders | ForEach-Object { return $_ }) -join ',')
        }
    }

    ((Get-Date).tostring() + " ------START------`nScript version: $scriptVersion`n") | Out-File $logFilePath -Append

    $cwcOutputDirPath = $OutDir + "\" + $cwcOutputDirName
    $outFileName = "CWCCreatorValues.csv"
    $outputFilePath = "$cwcOutputDirPath\" + $outFileName

    $baseUrl = GetBaseUrl
    ((Get-Date).tostring() + " Base Url $baseUrl") | Out-File -FilePath  $logFilePath -Append
    Write-Host "------------------------------------------------------------`n"
    Write-Host "Started process for fetching creator email for input CWC Container ID's"
    "Started process for fetching creator email for input CWC Container ID's" | Out-File $logFilePath -Append
    $stateData = Import-Csv $stateFilePath

    # Read CWC containerId's csv file
    $cwcCsvInput = Import-Csv -Path $CWCCreatorCsvFilePath
    $csvInputData = @($cwcCsvInput)
    $batchSize = 100
    $j = 0
    while ($j -lt $csvInputData.Count) {
        if ($csvInputData.Count -eq 1) {
            $batch = $csvInputData[0]
        }
        else {
            $batch = $csvInputData[$j..($j + $batchSize - 1)]
        }
        $data = $batch | Where-Object { ($_.'ContainerId' -notin $stateData.'ContainerId') -and (Test-IsGuid $_.'ContainerId' $logFilePath) }
        if ([string]::IsNullOrEmpty($data)) {
            $j = $j + $batchSize
            continue
        }
        if (!(Test-Path $cwcOutputDirName)) {
            $cwcOutputDirPath = CreateDirectory $cwcOutputDirName
        }
        if (!(Test-Path $outputFilePath -PathType leaf)) {
            CreateFile $outputFilePath 
        }

        FetchCreatorDetails -baseUrl $baseUrl -containerIds $data -outputFilePath $outputFilePath -logFilepath $logFilePath -stateFilepath $stateFilePath
        $j = $j + $batchSize
    }
}

$scriptVersion = '1.14'

Write-Host "Script version: $scriptVersion`n"

if ($GenerateMasterContainerList -eq 'true') {
    Write-Host "Generating Stream Classic master list...`n"
    $token = Get-Content -Path $InputFile
    $logFilePath = $OutDir + '\logs.txt'
    $masterListLogFilePath = $OutDir + '\masterListLog.txt'
    $masterListStateDirectoryPath = CreateDirectory 'masterListState'
    $masterListStateFilePath = $masterListStateDirectoryPath + '\masterListState.csv'
    $masterListLastRunFilePath = $masterListStateDirectoryPath + '\masterListLastRunFolder.txt'

    StartMasterListReportGeneration
    exit
}

if ($CreateDestinationPathMappingForM365GroupContainers -ne 'true') {
    Write-Host "`n////////////////////////////////////////////////////"
    Write-Host "Generating Stream Classic video report..."
    Write-Host "////////////////////////////////////////////////////`n"

    $token = Get-Content -Path $InputFile

    $logFilePath = $OutDir + '\logs.txt'

    $parallelThreads = 0
    $firstOrchestration = $true;

    [string]$filter = ''

    GenerateFilterString ([ref]$filter)

    StartReportGeneration
}
else {

    $randomGuidForOutFiles = New-Guid
    $logFilePath = $OutDir + '\logs_' + $randomGuidForOutFiles.Guid + '.txt'

    if (!(Test-Path $logFilePath -PathType leaf)) {
        CreateFile $logFilePath
    }

    Write-Host "Installing module ExchangeOnlineManagement, which requires admin priviledge on script"
    #Install Exchange Online Management Shell. Needs admin access to run
    Install-Module -Name ExchangeOnlineManagement
 
    #Connect to Exchange Online to allow M365Group details to be fetched. User will be prompted to login. This will work only for Exchange online admins. Documentation: https://learn.microsoft.com/en-us/powershell/module/exchange/get-unifiedgroup?view=exchange-ps
    Connect-ExchangeOnline

    try {
        #Invoke function to generate destination path mapping.
        CreateMigrationDestinationPathMapping
    }
    catch {
        #Log error.
        ((Get-Date).tostring() + ' ' + $Error[0].ErrorDetails + "`n") | Out-File $logFilePath -Append
        ((Get-Date).tostring() + ' ' + $Error[0].Exception + "`n") | Out-File $logFilePath -Append
        
        Write-Host "Please check error for details. If required, kindly reach out to Customer Support and share the log file: $logFilePath"
    }    
} 


if ($CreateDestinationPathMappingForM365GroupContainers -ne 'true' -and $CWCCreatorDetailMapping -eq 'true') {

    $token = Get-Content -Path $InputFile
    $cwcreportFolderName = 'CWCCreatorReport'
    $CWCCreatorReport = CreateDirectory $cwcreportFolderName
    $timeStamp = Get-Date -Format FileDateTime
    # Creating the new output directory 
    $cwcOutputDirName = $cwcreportFolderName + '\' + $timeStamp
    
    $stateDirectoryPath = CreateDirectory 'CWCCreatorReport\state'
    $stateFilePath = $stateDirectoryPath + '\state_cwc.csv'
    $logFilePath = $CWCCreatorReport + '\logs.txt'
    if (!(Test-Path $logFilePath -PathType leaf)) {
        CreateFile $logFilePath
    }
    $cwcStartTime = Get-Date
    $isValidPath = $true
    try {
        $cwcDirectoryPath = $OutDir + '\CWC'
        $cwcDirectoryCsvPath = $cwcDirectoryPath + '\CWCContainerIds.csv'
        if ([string]::IsNullOrEmpty($CWCCreatorCsvFilePath)) {
            $CWCCreatorCsvFilePath = $cwcDirectoryCsvPath
            if (Test-Path $CWCCreatorCsvFilePath) {
                $csvContent = Get-Content $CWCCreatorCsvFilePath
                if ($csvContent.Length -gt 0) {
                    Write-Host "Reading csv file on CWCCreatorCsvFilePath: $CWCCreatorCsvFilePath"
                    "Reading csv file on CWCCreatorCsvFilePath: $CWCCreatorCsvFilePath" | Out-File $logFilePath -Append
                    
                }
                else {
                    Write-Host "CSV file exists at CWCCreatorCsvFilePath $CWCCreatorCsvFilePath but is empty." -ForegroundColor Yellow
                    "CSV file exists at CWCCreatorCsvFilePath $CWCCreatorCsvFilePath but is empty" | Out-File $logFilePath -Append
                    $isValidPath = $false
                }
            }
            else {
                Write-Host "CWCCreatorCsvFilePath is required. CSV file does not exist at $cwcDirectoryCsvPath" -ForegroundColor Yellow
                "CWCCreatorCsvFilePath is required. CSV file does not exist at $cwcDirectoryCsvPath" | Out-File $logFilePath -Append
                $isValidPath = $false
            }
           
        }
        #Invoke function to generate CWC creator mapping.
        if ( $isValidPath) {
            GetCWCCreatorDetails $cwcOutputDirName $logFilepath $stateFilepath
        }
        if (Test-Path $cwcDirectoryPath) {
            Remove-Item $cwcDirectoryPath -Force -Recurse
        }
        $cwcEndTime = Get-Date
        Write-Host "Time elapsed to fetch CWC creator details: "($cwcEndTime - $cwcStartTime)"`n"
        ((Get-Date).tostring() + " Time elapsed to fetch CWC creator details: " + ($cwcEndTime - $cwcStartTime) + "`n") | Out-File $logFilePath -Append
    }
    catch {
        ((Get-Date).tostring() + ' ' + $Error[0].ErrorDetails + "`n") | Out-File $logFilePath -Append
        ((Get-Date).tostring() + ' ' + $Error[0].Exception + "`n") | Out-File $logFilePath -Append

        Write-Host "Please check error for details. If required, kindly reach out to Customer Support and share the log file: $logFilePath"  -ForegroundColor Red
    }
}
# SIG # Begin signature block
# MIIoKgYJKoZIhvcNAQcCoIIoGzCCKBcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCATJ/TN3np3ViPU
# dAPR8LxciF0d4Qmi3yGOFzJjbgo7eqCCDXYwggX0MIID3KADAgECAhMzAAADrzBA
# DkyjTQVBAAAAAAOvMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjMxMTE2MTkwOTAwWhcNMjQxMTE0MTkwOTAwWjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDOS8s1ra6f0YGtg0OhEaQa/t3Q+q1MEHhWJhqQVuO5amYXQpy8MDPNoJYk+FWA
# hePP5LxwcSge5aen+f5Q6WNPd6EDxGzotvVpNi5ve0H97S3F7C/axDfKxyNh21MG
# 0W8Sb0vxi/vorcLHOL9i+t2D6yvvDzLlEefUCbQV/zGCBjXGlYJcUj6RAzXyeNAN
# xSpKXAGd7Fh+ocGHPPphcD9LQTOJgG7Y7aYztHqBLJiQQ4eAgZNU4ac6+8LnEGAL
# go1ydC5BJEuJQjYKbNTy959HrKSu7LO3Ws0w8jw6pYdC1IMpdTkk2puTgY2PDNzB
# tLM4evG7FYer3WX+8t1UMYNTAgMBAAGjggFzMIIBbzAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQURxxxNPIEPGSO8kqz+bgCAQWGXsEw
# RQYDVR0RBD4wPKQ6MDgxHjAcBgNVBAsTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEW
# MBQGA1UEBRMNMjMwMDEyKzUwMTgyNjAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzci
# tW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEG
# CCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0
# MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAISxFt/zR2frTFPB45Yd
# mhZpB2nNJoOoi+qlgcTlnO4QwlYN1w/vYwbDy/oFJolD5r6FMJd0RGcgEM8q9TgQ
# 2OC7gQEmhweVJ7yuKJlQBH7P7Pg5RiqgV3cSonJ+OM4kFHbP3gPLiyzssSQdRuPY
# 1mIWoGg9i7Y4ZC8ST7WhpSyc0pns2XsUe1XsIjaUcGu7zd7gg97eCUiLRdVklPmp
# XobH9CEAWakRUGNICYN2AgjhRTC4j3KJfqMkU04R6Toyh4/Toswm1uoDcGr5laYn
# TfcX3u5WnJqJLhuPe8Uj9kGAOcyo0O1mNwDa+LhFEzB6CB32+wfJMumfr6degvLT
# e8x55urQLeTjimBQgS49BSUkhFN7ois3cZyNpnrMca5AZaC7pLI72vuqSsSlLalG
# OcZmPHZGYJqZ0BacN274OZ80Q8B11iNokns9Od348bMb5Z4fihxaBWebl8kWEi2O
# PvQImOAeq3nt7UWJBzJYLAGEpfasaA3ZQgIcEXdD+uwo6ymMzDY6UamFOfYqYWXk
# ntxDGu7ngD2ugKUuccYKJJRiiz+LAUcj90BVcSHRLQop9N8zoALr/1sJuwPrVAtx
# HNEgSW+AKBqIxYWM4Ev32l6agSUAezLMbq5f3d8x9qzT031jMDT+sUAoCw0M5wVt
# CUQcqINPuYjbS1WgJyZIiEkBMIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQg
# Q29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03
# a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akr
# rnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0Rrrg
# OGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy
# 4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9
# sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAh
# dCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8k
# A/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTB
# w3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmn
# Eyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90
# lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0w
# ggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2o
# ynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBa
# BgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
# bC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsG
# AQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNV
# HSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsG
# AQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABl
# AG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKb
# C5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11l
# hJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6
# I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0
# wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560
# STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQam
# ASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGa
# J+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ah
# XJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA
# 9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33Vt
# Y5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr
# /Xmfwb1tbWrJUnMTDXpQzTGCGgowghoGAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAAOvMEAOTKNNBUEAAAAAA68wDQYJYIZIAWUDBAIB
# BQCgga4wGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIGBEpJI71uSXDyaFFkdKbkr0
# RVayIUfmYreLeOo6mw4rMEIGCisGAQQBgjcCAQwxNDAyoBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20wDQYJKoZIhvcNAQEB
# BQAEggEARP7Z3Q3F7VWSYTD5CgzCWZVGzwVOJ5sl27X4JWEYjmQDjLGsU9ttWnY1
# ZDhfsEoK38kqT5Q1PcufhxqnBL/UoqFa1js+vRiU/gQYuMk7OVJjTsV8lIofNlYl
# LhOOmfgGjFl/tq6r+OJ9CvAOWMNCsBPkSubiahG5/+Nbs04TFz9HSe22FtF+rKce
# eK92LLuUVbPxOk59sZdVOAQGPnUSscueFECIfzC1UKhwxmaIrxo2jwucogHTaKk3
# OkUG/9Dpy5HqTLis2B5QcOC9E8y+iCB+9geMAAd+J8chYNPjQ5L1lryTrQHjbBI8
# TPINzQJyo6YMjqSbNwJQbbnkEV82UqGCF5QwgheQBgorBgEEAYI3AwMBMYIXgDCC
# F3wGCSqGSIb3DQEHAqCCF20wghdpAgEDMQ8wDQYJYIZIAWUDBAIBBQAwggFSBgsq
# hkiG9w0BCRABBKCCAUEEggE9MIIBOQIBAQYKKwYBBAGEWQoDATAxMA0GCWCGSAFl
# AwQCAQUABCDHypNvHOb++TDFQSXdr7IQ77Q3OKONn60B4yn74GeDtQIGZc3461rB
# GBMyMDI0MDMwNTE2Mjg1OS44MzZaMASAAgH0oIHRpIHOMIHLMQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1l
# cmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046QTkzNS0w
# M0UwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2Wg
# ghHqMIIHIDCCBQigAwIBAgITMwAAAekPcTB+XfESNgABAAAB6TANBgkqhkiG9w0B
# AQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYD
# VQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAeFw0yMzEyMDYxODQ1
# MjZaFw0yNTAzMDUxODQ1MjZaMIHLMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25z
# MScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046QTkzNS0wM0UwLUQ5NDcxJTAjBgNV
# BAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQCsmowxQRVgp4TSc3nTa6yrAPJnV6A7aZYnTw/yx90u
# 1DSH89nvfQNzb+5fmBK8ppH76TmJzjHUcImd845A/pvZY5O8PCBu7Gq+x5Xe6plQ
# t4xwVUUcQITxklOZ1Rm9fJ5nh8gnxOxaezFMM41sDI7LMpKwIKQMwXDctYKvCyQy
# 6kO2sVLB62kF892ZwcYpiIVx3LT1LPdMt1IeS35KY5MxylRdTS7E1Jocl30NgcBi
# JfqnMce05eEipIsTO4DIn//TtP1Rx57VXfvCO8NSCh9dxsyvng0lUVY+urq/G8QR
# FoOl/7oOI0Rf8Qg+3hyYayHsI9wtvDHGnT30Nr41xzTpw2I6ZWaIhPwMu5DvdkEG
# zV7vYT3tb9tTviY3psul1T5D938/AfNLqanVCJtP4yz0VJBSGV+h66ZcaUJOxpbS
# IjImaOLF18NOjmf1nwDatsBouXWXFK7E5S0VLRyoTqDCxHG4mW3mpNQopM/U1WJn
# jssWQluK8eb+MDKlk9E/hOBYKs2KfeQ4HG7dOcK+wMOamGfwvkIe7dkylzm8BeAU
# QC8LxrAQykhSHy+FaQ93DAlfQYowYDtzGXqE6wOATeKFI30u9YlxDTzAuLDK073c
# ndMV4qaD3euXA6xUNCozg7rihiHUaM43Amb9EGuRl022+yPwclmykssk30a4Rp3v
# 9QIDAQABo4IBSTCCAUUwHQYDVR0OBBYEFJF+M4nFCHYjuIj0Wuv+jcjtB+xOMB8G
# A1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY3Jvc29mdCUy
# MFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNybDBsBggrBgEFBQcBAQRgMF4w
# XAYIKwYBBQUHMAKGUGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY2Vy
# dHMvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3J0MAwG
# A1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwDgYDVR0PAQH/BAQD
# AgeAMA0GCSqGSIb3DQEBCwUAA4ICAQBWsSp+rmsxFLe61AE90Ken2XPgQHJDiS4S
# bLhvzfVjDPDmOdRE75uQohYhFMdGwHKbVmLK0lHV1Apz/HciZooyeoAvkHQaHmLh
# wBGkoyAAVxcaaUnHNIUS9LveL00PwmcSDLgN0V/Fyk20QpHDEukwKR8kfaBEX83A
# yvQzlf/boDNoWKEgpdAsL8SzCzXFLnDozzCJGq0RzwQgeEBr8E4K2wQ2WXI/ZJxZ
# S/+d3FdwG4ErBFzzUiSbV2m3xsMP3cqCRFDtJ1C3/JnjXMChnm9bLDD1waJ7TPp5
# wYdv0Ol9+aN0t1BmOzCj8DmqKuUwzgCK9Tjtw5KUjaO6QjegHzndX/tZrY792dfR
# AXr5dGrKkpssIHq6rrWO4PlL3OS+4ciL/l8pm+oNJXWGXYJL5H6LNnKyXJVEw/1F
# bO4+Gz+U4fFFxs2S8UwvrBbYccVQ9O+Flj7xTAeITJsHptAvREqCc+/YxzhIKkA8
# 8Q8QhJKUDtazatJH7ZOdi0LCKwgqQO4H81KZGDSLktFvNRhh8ZBAenn1pW+5UBGY
# z2GpgcxVXKT1CuUYdlHR9D6NrVhGqdhGTg7Og/d/8oMlPG3YjuqFxidiIsoAw2+M
# hI1zXrIi56t6JkJ75J69F+lkh9myJJpNkx41sSB1XK2jJWgq7VlBuP1BuXjZ3qgy
# m9r1wv0MtTCCB3EwggVZoAMCAQICEzMAAAAVxedrngKbSZkAAAAAABUwDQYJKoZI
# hvcNAQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAw
# DgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24x
# MjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1dGhvcml0eSAy
# MDEwMB4XDTIxMDkzMDE4MjIyNVoXDTMwMDkzMDE4MzIyNVowfDELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgUENBIDIwMTAwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoIC
# AQDk4aZM57RyIQt5osvXJHm9DtWC0/3unAcH0qlsTnXIyjVX9gF/bErg4r25Phdg
# M/9cT8dm95VTcVrifkpa/rg2Z4VGIwy1jRPPdzLAEBjoYH1qUoNEt6aORmsHFPPF
# dvWGUNzBRMhxXFExN6AKOG6N7dcP2CZTfDlhAnrEqv1yaa8dq6z2Nr41JmTamDu6
# GnszrYBbfowQHJ1S/rboYiXcag/PXfT+jlPP1uyFVk3v3byNpOORj7I5LFGc6XBp
# Dco2LXCOMcg1KL3jtIckw+DJj361VI/c+gVVmG1oO5pGve2krnopN6zL64NF50Zu
# yjLVwIYwXE8s4mKyzbnijYjklqwBSru+cakXW2dg3viSkR4dPf0gz3N9QZpGdc3E
# XzTdEonW/aUgfX782Z5F37ZyL9t9X4C626p+Nuw2TPYrbqgSUei/BQOj0XOmTTd0
# lBw0gg/wEPK3Rxjtp+iZfD9M269ewvPV2HM9Q07BMzlMjgK8QmguEOqEUUbi0b1q
# GFphAXPKZ6Je1yh2AuIzGHLXpyDwwvoSCtdjbwzJNmSLW6CmgyFdXzB0kZSU2LlQ
# +QuJYfM2BjUYhEfb3BvR/bLUHMVr9lxSUV0S2yW6r1AFemzFER1y7435UsSFF5PA
# PBXbGjfHCBUYP3irRbb1Hode2o+eFnJpxq57t7c+auIurQIDAQABo4IB3TCCAdkw
# EgYJKwYBBAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3FQIEFgQUKqdS/mTEmr6CkTxG
# NSnPEP8vBO4wHQYDVR0OBBYEFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMFwGA1UdIARV
# MFMwUQYMKwYBBAGCN0yDfQEBMEEwPwYIKwYBBQUHAgEWM2h0dHA6Ly93d3cubWlj
# cm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0b3J5Lmh0bTATBgNVHSUEDDAK
# BggrBgEFBQcDCDAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMC
# AYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTV9lbLj+iiXGJo0T2UkFvX
# zpoYxDBWBgNVHR8ETzBNMEugSaBHhkVodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20v
# cGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcmwwWgYI
# KwYBBQUHAQEETjBMMEoGCCsGAQUFBzAChj5odHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNydDANBgkqhkiG
# 9w0BAQsFAAOCAgEAnVV9/Cqt4SwfZwExJFvhnnJL/Klv6lwUtj5OR2R4sQaTlz0x
# M7U518JxNj/aZGx80HU5bbsPMeTCj/ts0aGUGCLu6WZnOlNN3Zi6th542DYunKmC
# VgADsAW+iehp4LoJ7nvfam++Kctu2D9IdQHZGN5tggz1bSNU5HhTdSRXud2f8449
# xvNo32X2pFaq95W2KFUn0CS9QKC/GbYSEhFdPSfgQJY4rPf5KYnDvBewVIVCs/wM
# nosZiefwC2qBwoEZQhlSdYo2wh3DYXMuLGt7bj8sCXgU6ZGyqVvfSaN0DLzskYDS
# PeZKPmY7T7uG+jIa2Zb0j/aRAfbOxnT99kxybxCrdTDFNLB62FD+CljdQDzHVG2d
# Y3RILLFORy3BFARxv2T5JL5zbcqOCb2zAVdJVGTZc9d/HltEAY5aGZFrDZ+kKNxn
# GSgkujhLmm77IVRrakURR6nxt67I6IleT53S0Ex2tVdUCbFpAUR+fKFhbHP+Crvs
# QWY9af3LwUFJfn6Tvsv4O+S3Fb+0zj6lMVGEvL8CwYKiexcdFYmNcP7ntdAoGokL
# jzbaukz5m/8K6TT4JDVnK+ANuOaMmdbhIurwJ0I9JZTmdHRbatGePu1+oDEzfbzL
# 6Xu/OHBE0ZDxyKs6ijoIYn/ZcGNTTY3ugm2lBRDBcQZqELQdVTNYs6FwZvKhggNN
# MIICNQIBATCB+aGB0aSBzjCByzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEn
# MCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOkE5MzUtMDNFMC1EOTQ3MSUwIwYDVQQD
# ExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloiMKAQEwBwYFKw4DAhoDFQCr
# aYf1xDk2rMnU/VJo2GGK1nxo8aCBgzCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFBDQSAyMDEwMA0GCSqGSIb3DQEBCwUAAgUA6ZGCKTAiGA8yMDI0MDMwNTExMzYw
# OVoYDzIwMjQwMzA2MTEzNjA5WjB0MDoGCisGAQQBhFkKBAExLDAqMAoCBQDpkYIp
# AgEAMAcCAQACAhiWMAcCAQACAhMWMAoCBQDpktOpAgEAMDYGCisGAQQBhFkKBAIx
# KDAmMAwGCisGAQQBhFkKAwKgCjAIAgEAAgMHoSChCjAIAgEAAgMBhqAwDQYJKoZI
# hvcNAQELBQADggEBAEPMChrbNTADzeSh+int8bbXCHX/bm4naxVPyp9IGqHejDKj
# Dxx6kajCz8rqnANLoghtEftyaIkxOLTy6C0SjrtSmQnMVy413ukpYIbXuNON1NU/
# PdOie3o1nGzt05FZv4Rquzo4mHInZMj8jRl0w5DZWu1kCnD/i2zVxQSzUMyTgvys
# qERSnvpd560wT1hGQoHc81VN2X+rEOq2YnQoBPEFnQL4d8qFpbQLCL63AJcaeHd3
# 7BO6kNQbpKSYAXy9k7w8DSYN6cCtuPlKjSzuhzupxjw1+KiGUvZ2E0mI53lp7xYX
# At6hw52R1u0CA7m31wxn/ED3RjcyC2LaRE5sKkwxggQNMIIECQIBATCBkzB8MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNy
# b3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAekPcTB+XfESNgABAAAB6TAN
# BglghkgBZQMEAgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMC8G
# CSqGSIb3DQEJBDEiBCC6awjNikF/rZL9i8WEGbtbV/iISfA0mJy13T69qKsKpTCB
# +gYLKoZIhvcNAQkQAi8xgeowgecwgeQwgb0EIKSQkniXaTcmj1TKQWF+x2U4riVo
# rGD8TwmgVbN9qsQlMIGYMIGApH4wfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldh
# c2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
# b3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIw
# MTACEzMAAAHpD3Ewfl3xEjYAAQAAAekwIgQgWg3PHbWeCdLmt2xeBQ6VDQj4Qu0P
# e2/sNQsBWYQZi/UwDQYJKoZIhvcNAQELBQAEggIAondeQml6T94VrPDKSIvNp4Ys
# WFkJoWmq5+JUc+2o9a83lq+bLDKUZ2Cc8vVm8lEIXPZuBEiQS3HQKCjGvc9XKD+V
# U8A6Bm9XRQvfxaa6iJU+fH18gNlN4C6uXtqSSZDY/mva7+7UKc3onkoCcVn3FeWd
# +/bHE7R0baBFx/IwjEnVi+y0nfXVA683Ab0CRrW0Gy6OTi53HSK2GIycMiQB7Wug
# cSpDphfSnHvPThxe/aaY/aKD4ZvFIGgy39piLO/vo9Zn2h9KKwCEFgQB7Vf3uYKT
# YPVFMZBpLH1Oq/7NLrhSoX4zR2fvXwjF/tfK3YPUS3y0u54DmsE27PtDyy3Tbt+T
# KOmQQsT4t9u1OAQiXV1SVCcA97WIVVS0WsjT6h0b9sDbHi7U8zUODQssYjaguJzN
# au70xsOV9RqPVbMIYtEJdPQvXa9L/wLjETHIxNbEKVKgE3HeTA4pZB0HRJi0FVWt
# JOg1Aacetv4wNplwM8T1q/RiyMnXx0dID8YkTu3GZLNERBhgV+D7QuMxPZ3lg4hu
# z7I1H4oKlY/ZpExdqcjs9MvrSK56zMx8yfRij31NzAkNjSQePeACZGUHdHy1mdye
# +bUn3dg/dmWVSoO9LEEgxtLBXpIvmx/nrnkfPX+nyU4lDVF2bbb3f/bmHxe+gEps
# uvHBmyPBLJejDq0BHK8=
# SIG # End signature block
