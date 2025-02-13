param (
    [Parameter(Mandatory=$true)]
    [ValidateSet('All', 'Selected')]
    [string]$Mode = 'All',
    [switch]$AllowPrerelease,
    [string]$Name = '*',
    [ValidateSet('AllUsers', 'CurrentUser')][string]$Scope = 'AllUsers',
    [switch]$WhatIf
)

function Update-AllModules {
    # Test admin privileges
    if ($Scope -eq 'AllUsers') {
        if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
            Write-Warning ("Function {0} needs admin privileges. Break now." -f $MyInvocation.MyCommand)
            return
        }
    }
   
    # Get installed modules first
    Write-Host ("Retrieving installed modules...") -ForegroundColor Green
    $InstalledModules = Get-InstalledModule -Name $Name -ErrorAction SilentlyContinue | 
        Select-Object Name, Version | 
        Sort-Object Name

    if (-not $InstalledModules) {
        Write-Host ("No modules found.") -ForegroundColor Gray
        return
    }

    # If Mode is Selected, show selection grid immediately
    if ($Mode -eq 'Selected') {
        $SelectedModules = $InstalledModules | Out-GridView -Title "Select modules to check for updates" -PassThru
        if (-not $SelectedModules) {
            Write-Host ("No modules selected for update.") -ForegroundColor Yellow
            return
        }
        $InstalledModules = $SelectedModules
    }

    # Now check for updates only on selected/all modules
    Write-Host ("Checking for available updates...") -ForegroundColor Green
    $CurrentModules = $InstalledModules | ForEach-Object {
        $Current = $_
        Write-Host ("Checking {0}..." -f $Current.Name) -ForegroundColor Gray
        $Latest = Find-Module -Name $Current.Name -ErrorAction SilentlyContinue
        [PSCustomObject]@{
            Name = $Current.Name
            CurrentVersion = $Current.Version
            LatestVersion = $Latest.Version
            UpdateAvailable = ([version]$Latest.Version -gt [version]$Current.Version)
        }
    }
    
    $ModulesCount = $CurrentModules.Count
    $DigitsLength = $ModulesCount.ToString().Length
    Write-Host ("{0} modules will be processed for update." -f $ModulesCount) -ForegroundColor Gray
   
    # Show status of AllowPrerelease Switch
    ''
    if ($AllowPrerelease) {
        Write-Host ("Updating installed modules to the latest PreRelease version ...") -ForegroundColor Green
    } else {
        Write-Host ("Updating installed modules to the latest Production version ...") -ForegroundColor Green
    }
   
    # Update modules that need updating
    $i = 0
    foreach ($Module in $CurrentModules) {
        $i++
        $Counter = ("[{0,$DigitsLength}/{1,$DigitsLength}]" -f $i, $ModulesCount)
        $CounterLength = $Counter.Length

        if ($Module.UpdateAvailable) {
            Write-Host ('{0} Updating module {1} from version {2} to {3} ...' -f $Counter, $Module.Name, $Module.CurrentVersion, $Module.LatestVersion) -ForegroundColor Green
            try {
                Update-Module -Name $Module.Name -AllowPrerelease:$AllowPrerelease -AcceptLicense -Scope:$Scope -Force:$True -ErrorAction Stop -WhatIf:$WhatIf.IsPresent
                
                # Remove old versions
                $AllVersions = Get-InstalledModule -Name $Module.Name -AllVersions | Sort-Object PublishedDate -Descending
                if ($AllVersions.Count -gt 1) {
                    foreach ($Version in $AllVersions | Select-Object -Skip 1) {
                        try {
                            Write-Host ("{0,$CounterLength} Uninstalling previous version {1} of module {2} ..." -f ' ', $Version.Version, $Module.Name) -ForegroundColor Gray
                            Uninstall-Module -Name $Module.Name -RequiredVersion $Version.Version -Force:$True -ErrorAction Stop -AllowPrerelease -WhatIf:$WhatIf.IsPresent
                        } catch {
                            Write-Warning ("{0,$CounterLength} Error uninstalling previous version {1} of module {2}!" -f ' ', $Version.Version, $Module.Name)
                        }
                    }
                }
            } catch {
                Write-Host ("{0,$CounterLength} Error updating module {1}!" -f ' ', $Module.Name) -ForegroundColor Red
            }
        } else {
            Write-Host ('{0} Module {1} is already at latest version {2}' -f $Counter, $Module.Name, $Module.CurrentVersion) -ForegroundColor Gray
        }
    }
}

# Run the function
Update-AllModules