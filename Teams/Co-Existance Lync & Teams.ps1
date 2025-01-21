# Set-TeamsDefaultIMApp.ps1
# Change the current user's Default IM App to Teams

# IM Provider registry path
$imProviderPath = "HKCU:\Software\IM Providers"

# Retrieve current IM Provider Information
$imProvider = Get-ItemProperty -Path $imProviderPath

# If there is a current provider set, set this as the previous IM provider
# This is if a user unticks the setting in Teams, Teams knows what to fallback to
if ($imProvider.DefaultIMApp) {
    
    # Build Path for Teams
    $teamsPath = Join-Path -Path $imProviderPath -ChildPath "Teams"

    # Check Teams IM Provider path exists (it should if Teams has been run before)
    if (Test-Path $teamsPath) {

        Write-Host "$teamsPath already exists, no action needed..."

    }
    else {

        Write-Warning "$teamsPath does not exist, creating..."
        New-Item -Path $imProviderPath -Name "Teams"

    }

    # Path should now be created
    if (Test-Path $teamsPath) {

        Set-ItemProperty -Path $teamsPath PreviousDefaultIMApp -Value $imProvider.DefaultIMApp -Type String

    }
    else {

        Write-Warning "Unable to create $teamsPath!"

    }

}

# Set Teams as Deafult IM App
Set-ItemProperty -Path $imProviderPath -Name "DefaultIMApp" -Value "Teams"