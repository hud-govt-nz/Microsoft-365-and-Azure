# Simple Compliance Report Script with optimized performance
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string[]]$ComplianceTagScope = @("all"),
    
    [Parameter(Mandatory = $false)]
    [int]$MaxParallelJobs = 5,  # Reduced default parallel jobs for better stability
    
    [Parameter(Mandatory = $false)]
    [int]$BatchSize = 2000      # Batch size for processing items
)

# Initialize basic configuration
$domain = "mhud"
$adminSiteURL = "https://$domain-Admin.SharePoint.com"
$TenantURL = "https://$domain.sharepoint.com"
$dateTime = (Get-Date).ToString("dd-MM-yyyy-HH-mm")
$outputPath = "C:\HUD\06_Reporting\SPO\Reports\SimpleLabelsReport_$dateTime.csv"
$tempFolder = "C:\HUD\06_Reporting\SPO\Temp"

# Create required directories
$paths = @($outputPath, $tempFolder) | ForEach-Object { Split-Path $_ -Parent }
foreach ($path in $paths) {
    if (-not (Test-Path $path)) { New-Item -ItemType Directory -Path $path -Force | Out-Null }
}

# Create CSV file with headers
@('FilePath', 'ComplianceTag', 'SiteUrl', 'Library') | ConvertTo-Csv -NoTypeInformation | Set-Content $outputPath

# Define libraries to exclude
$ExcludedLibraries = @(
    "Form Templates", "Preservation Hold Library", "Site Assets", "Site Pages", 
    "Images", "Pages", "Settings", "Videos", "Site Collection Documents", 
    "Site Collection Images", "Style Library", "AppPages", "Apps for SharePoint", 
    "Apps for Office"
)

function Process-Site {
    param(
        $siteUrl,
        $tempFile,
        $clientId,
        $tenant,
        $thumbprint,
        $excludedLibs,
        $tagScope,
        $batchSize
    )
    
    try {
        Write-Host "Processing site: $siteUrl" -ForegroundColor Cyan
        
        # Connect to the site with retry logic
        $maxRetries = 3
        $retryCount = 0
        $connected = $false
        
        while (-not $connected -and $retryCount -lt $maxRetries) {
            try {
                Import-Module PnP.PowerShell -ErrorAction Stop
                Connect-PnPOnline -url $siteUrl -ClientId $clientId -Tenant $tenant -Thumbprint $thumbprint
                $connected = $true
            }
            catch {
                $retryCount++
                Write-Warning "Connection attempt $retryCount failed for site $siteUrl"
                Start-Sleep -Seconds ($retryCount * 5)
            }
        }
        
        if (-not $connected) {
            throw "Failed to connect to site after $maxRetries attempts"
        }
        
        # Create temporary CSV
        @('FilePath', 'ComplianceTag', 'SiteUrl', 'Library') | ConvertTo-Csv -NoTypeInformation | Set-Content $tempFile
        
        # Get document libraries
        $DocLibraries = Get-PnPList -Includes BaseType, Hidden, Title | Where-Object {
            $_.BaseType -eq "DocumentLibrary" -and 
            $_.Hidden -eq $False -and 
            $_.Title -notin $excludedLibs
        }
        
        $libraryCount = $DocLibraries.Count
        $currentLibrary = 0
        
        foreach ($library in $DocLibraries) {
            $currentLibrary++
            Write-Progress -Id 2 -Activity "Processing Libraries" -Status "Library $currentLibrary of $libraryCount : $($library.Title)" -PercentComplete (($currentLibrary / $libraryCount) * 100)
            
            try {
                # Process items in batches
                $position = $null
                $processedItems = 0
                
                do {
                    $items = Get-PnPListItem -List $library.Title -Fields "FileRef","_ComplianceTag" -PageSize $batchSize -Position $position
                    $position = $items.ListItemCollectionPosition
                    
                    $labeledItems = $items | Where-Object { 
                        $_.FieldValues["_ComplianceTag"] -and
                        ($tagScope.Count -eq 1 -and $tagScope[0] -eq "all" -or 
                         $_.FieldValues["_ComplianceTag"] -in $tagScope)
                    }
                    
                    foreach ($item in $labeledItems) {
                        [PSCustomObject]@{
                            FilePath = $item.FieldValues["FileRef"]
                            ComplianceTag = $item.FieldValues["_ComplianceTag"]
                            SiteUrl = $siteUrl
                            Library = $library.Title
                        } | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 | Add-Content -Path $tempFile
                        
                        $processedItems++
                        if ($processedItems % 100 -eq 0) {
                            Write-Progress -Id 3 -Activity "Processing Files" -Status "Found $processedItems tagged files" -PercentComplete -1
                        }
                    }
                    
                    # Force garbage collection to manage memory
                    if ($null -ne $position) {
                        [System.GC]::Collect()
                        Start-Sleep -Milliseconds 500
                    }
                    
                } while ($null -ne $position)
                
                Write-Progress -Id 3 -Activity "Processing Files" -Completed
            }
            catch {
                Write-Warning "Error processing library $($library.Title): $_"
            }
        }
        
        Write-Progress -Id 2 -Activity "Processing Libraries" -Completed
        return $true
    }
    catch {
        Write-Warning "Error processing site $siteUrl : $_"
        return $false
    }
    finally {
        try {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
        }
        catch { }
    }
}

try {
    # Connect to admin site
    Write-Host "Connecting to SharePoint Online Admin Center..." -ForegroundColor Cyan
    Connect-PnPOnline -Url $adminSiteURL -ClientId $env:DigitalSupportAppID -Tenant 'mhud.onmicrosoft.com' -Thumbprint $env:DigitalSupportCertificateThumbprint
    
    # Get all sites
    Write-Host "Retrieving SharePoint sites..." -ForegroundColor Cyan
    $sites = Get-PnPTenantSite -Filter "Url -like '$TenantURL'" | Where-Object { $_.Template -ne 'RedirectSite#0' }
    Write-Host "Found $($sites.Count) sites to process" -ForegroundColor Yellow
    
    $total = $sites.Count
    $processed = 0
    $jobs = @{}
    
    foreach ($site in $sites) {
        while ($jobs.Count -ge $MaxParallelJobs) {
            $completed = $jobs.Keys | Where-Object { $jobs[$_].Job.State -eq 'Completed' }
            foreach ($jobKey in $completed) {
                $jobData = $jobs[$jobKey]
                
                if (Test-Path $jobData.TempFile) {
                    $content = Get-Content $jobData.TempFile | Select-Object -Skip 1
                    if ($content) {
                        Add-Content -Path $outputPath -Value $content
                    }
                    Remove-Item $jobData.TempFile -Force
                }
                
                Remove-Job $jobData.Job -Force
                $jobs.Remove($jobKey)
                $processed++
                
                Write-Progress -Id 1 -Activity "Processing SharePoint Sites" -Status "Site $processed of $total" -PercentComplete (($processed / $total) * 100)
            }
            Start-Sleep -Seconds 2
        }
        
        $tempFile = Join-Path $tempFolder "site_$($processed)_$((New-Guid).ToString()).csv"
        $job = Start-Job -ScriptBlock ${function:Process-Site} -ArgumentList @(
            $site.Url,
            $tempFile,
            $env:DigitalSupportAppID,
            'mhud.onmicrosoft.com',
            $env:DigitalSupportCertificateThumbprint,
            $ExcludedLibraries,
            $ComplianceTagScope,
            $BatchSize
        )
        
        $jobs.Add($site.Url, @{
            Job = $job
            TempFile = $tempFile
        })
        
        Write-Host "Started processing: $($site.Url)" -ForegroundColor Gray
    }
    
    # Wait for remaining jobs
    Write-Host "Waiting for remaining jobs to complete..." -ForegroundColor Cyan
    while ($jobs.Count -gt 0) {
        $completed = $jobs.Keys | Where-Object { $jobs[$_].Job.State -eq 'Completed' }
        foreach ($jobKey in $completed) {
            $jobData = $jobs[$jobKey]
            
            if (Test-Path $jobData.TempFile) {
                $content = Get-Content $jobData.TempFile | Select-Object -Skip 1
                if ($content) {
                    Add-Content -Path $outputPath -Value $content
                }
                Remove-Item $jobData.TempFile -Force
            }
            
            Remove-Job $jobData.Job -Force
            $jobs.Remove($jobKey)
            $processed++
            
            Write-Progress -Id 1 -Activity "Processing SharePoint Sites" -Status "Site $processed of $total" -PercentComplete (($processed / $total) * 100)
        }
        Start-Sleep -Seconds 2
    }
    
    Write-Progress -Id 1 -Activity "Processing SharePoint Sites" -Completed
}
catch {
    Write-Warning "Critical error in main processing: $_"
}
finally {
    # Cleanup
    if (Test-Path $tempFolder) {
        Remove-Item $tempFolder -Recurse -Force
    }
    
    Write-Host "`nReport generation complete!" -ForegroundColor Green
    Write-Host "Report saved to: $outputPath" -ForegroundColor Green
    Write-Host "Total sites processed: $processed" -ForegroundColor Green
}