# Define the compliance tag names to search for.
$complianceTags = @(
            "DA-2.1.2 Strategic relationship management with NZ organisations",
            "DA-2.1.3 Operational relationship management records",
            "DA-2.1.4 Reporting and analysis records",
            "DA-2.2.1 Event photographs with a description of the content",
            "DA-2.2.3 Multimedia resources that can be reused for other use",
            "DA607-2.1.3 Operational relationship management records",
            "DA697-2.1.1 Stakeholder relationship programme management",
            "DA697-2.1.2 Strategic relationship management",
            "DA697-2.1.3 Operational relationship management records",
            "DA697-2.2.3 Multimedia Resources"
)


# Build the search query.
$queryParts = $complianceTags | ForEach-Object { '("ComplianceTag:' + $_ + '")' }
$query = $queryParts -join ' OR '
Write-Host "Constructed Search Query: $query"

# Name for the compliance search.
$searchName = "ScanComplianceTags"

# Create a new compliance search targeting SharePoint sites and M365 groups.
New-ComplianceSearch -Name $searchName -SharePointLocation All -ExchangeLocation M365Group -ContentMatchQuery $query

# Start the compliance search.
Start-ComplianceSearch -Identity $searchName

# Optionally, display the search status.
Get-ComplianceSearch -Identity $searchName | Format-List
