[CmdletBinding()]
param()

function Read-UserInput {
    # Get Site URL
    do {
        $SiteUrl = Read-Host "`nEnter SharePoint site URL (e.g., https://contoso.sharepoint.com/sites/mysite)"
    } while ([string]::IsNullOrWhiteSpace($SiteUrl))

    # Get Max Depth with default of -1 (unlimited)
    $MaxDepth = Read-Host "`nEnter maximum depth level (press Enter for unlimited)"
    if ([string]::IsNullOrWhiteSpace($MaxDepth)) {
        $MaxDepth = -1
    }

    # Get Output Path with default
    $defaultPath = "C:\HUD\06_Reporting\SPO"
    $OutputPath = Read-Host "`nEnter output path (press Enter for default: $defaultPath)"
    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        $OutputPath = $defaultPath
    }

    # Create output directory if it doesn't exist
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        Write-Host "Created output directory: $OutputPath" -ForegroundColor Green
    }

    return @{
        SiteUrl = $SiteUrl
        MaxDepth = [int]$MaxDepth
        OutputPath = $OutputPath
    }
}

# Get user input at the start
$userInput = Read-UserInput

# Add variables to script scope
$script:SiteUrl = $userInput.SiteUrl
$script:MaxDepth = $userInput.MaxDepth
$script:OutputPath = $userInput.OutputPath

Write-Host "`nUsing the following settings:" -ForegroundColor Cyan
Write-Host "Site URL: $SiteUrl"
Write-Host "Max Depth: $(if ($MaxDepth -eq -1) { 'Unlimited' } else { $MaxDepth })"
Write-Host "Output Path: $OutputPath`n"

Add-Type -AssemblyName System.Web

# Check if PnP PowerShell is installed
if (!(Get-Module -ListAvailable -Name "PnP.PowerShell")) {
    Install-Module -Name "PnP.PowerShell" -Force -Scope CurrentUser
}

# Add ImportExcel module check
if (!(Get-Module -ListAvailable -Name "ImportExcel")) {
    Install-Module -Name "ImportExcel" -Force -Scope CurrentUser
}
Import-Module ImportExcel

# Add class definition after the param block
class SharePointNode {
    [string]$Name
    [string]$Type  # 🌐, 📚, or 📁
    [string]$Url
    [string]$ComplianceTag
    [int]$Level
    [System.Collections.ArrayList]$Children

    SharePointNode([string]$name, [string]$type, [string]$url, [int]$level) {
        $this.Name = $name
        $this.Type = $type
        $this.Url = $url
        $this.Level = $level
        $this.Children = [System.Collections.ArrayList]::new()
        $this.ComplianceTag = "No Label"
    }

    [void]AddChild([SharePointNode]$child) {
        $this.Children.Add($child) | Out-Null
    }
}

# Function to get indent string for tree view
function Get-IndentString {
    param(
        $level,
        [switch]$Markdown
    )
    if ($Markdown) {
        return "    " * $level + "- "
    }
    return "  " * $level + "├─"
}

# Function to get compliance tag
function Get-ComplianceTag {
    param (
        $List,
        $FolderPath
    )
    try {
        if ($List) {
            Write-Verbose "Checking label for list: $($List.Title)"
            
            # Only attempt to get label once
            $label = Get-PnPLabel -List $List.Title -ValuesOnly -ErrorAction SilentlyContinue
            
            if ($null -ne $label -and $label.TagName) {
                Write-Verbose "  ✓ Found label for $($List.Title): $($label.TagName)"
                return $label.TagName
            }
            
            # If we're checking a folder and no direct label was found, check for inherited label
            if ($FolderPath) {
                $parentLabel = Get-PnPLabel -List $List.Title -ValuesOnly -ErrorAction SilentlyContinue
                if ($null -ne $parentLabel -and $parentLabel.TagName) {
                    Write-Verbose "  ✓ Found inherited label: $($parentLabel.TagName)"
                    return "$($parentLabel.TagName) (Inherited)"
                }
            }
            
            Write-Verbose "  ℹ No label found for $($List.Title)"
        }
    }
    catch {
        Write-Verbose "  ! Error getting label for $($List.Title): $_"
    }
    return "No Label"
}

# Function to recursively get folders and files
function Get-FolderHierarchy {
    param (
        $List,
        $FolderUrl = "",
        $Level = 0,
        [int]$MaxDepth = -1,
        [SharePointNode]$ParentNode
    )
    
    try {
        if ($MaxDepth -ne -1 -and $Level -ge $MaxDepth) {
            return
        }

        Write-Host "Processing folders in '$($List.Title)'" -ForegroundColor Yellow
        
        $folders = if ([string]::IsNullOrEmpty($FolderUrl)) {
            Write-Host "  Getting root folders..." -ForegroundColor Gray
            Get-PnPFolder -ListRootFolder $List | Get-PnPFolderInFolder | 
                Where-Object { $_.Name -notmatch '^_|^Forms$' }
        } else {
            Write-Host "  Getting folders from: $FolderUrl" -ForegroundColor Gray
            Get-PnPFolder -Url $FolderUrl | Get-PnPFolderInFolder |
                Where-Object { $_.Name -notmatch '^_|^Forms$' }
        }
        
        if ($folders.Count -gt 0) {
            Write-Host "  Found $($folders.Count) folders" -ForegroundColor Gray
        }
        
        foreach ($folder in $folders) {
            Write-Host "    Processing folder: $($folder.Name)" -ForegroundColor Cyan
            $baseUrl = [regex]::Match($List.Context.Url, 'https://[^/]+').Value
            $relativePath = $folder.ServerRelativeUrl.Split('/') | 
                Where-Object { $_ } | 
                ForEach-Object { [Uri]::EscapeDataString($_) } | 
                Join-String -Separator '/'
            $absoluteUrl = "$baseUrl/$relativePath"
            
            $folderNode = [SharePointNode]::new($folder.Name, "📁", $absoluteUrl, $Level)
            $folderNode.ComplianceTag = Get-ComplianceTag -List $List -FolderPath $folder.ServerRelativeUrl
            $ParentNode.AddChild($folderNode)
            
            Get-FolderHierarchy -List $list -FolderUrl $folder.ServerRelativeUrl -Level ($Level + 1) -MaxDepth $MaxDepth -ParentNode $folderNode
        }
    }
    catch {
        Write-Warning "Error accessing folder in $($List.Title) : $_"
    }
}

# Function to get document libraries
function Get-LibraryHierarchy {
    param (
        $SiteUrl,
        $Level = 0,
        [int]$MaxDepth = -1,
        [SharePointNode]$ParentNode
    )
    
    Write-Host "`nProcessing site: $SiteUrl" -ForegroundColor Green
    
    Connect-PnPOnline -Url $SiteUrl -ClientId $env:DigitalSupportAppID -Tenant 'mhud.onmicrosoft.com' -Thumbprint $env:DigitalSupportCertificateThumbprint
    $lists = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and $_.Hidden -eq $false }
    
    if ($lists.Count -gt 0) {
        Write-Host "Found $($lists.Count) document libraries" -ForegroundColor White
    }
    
    foreach ($list in $lists) {
        Write-Host "`nProcessing library: $($list.Title)" -ForegroundColor Yellow
        $baseUrl = [regex]::Match($SiteUrl, 'https://[^/]+').Value
        $relativePath = $list.RootFolder.ServerRelativeUrl.Split('/') | 
            Where-Object { $_ } | 
            ForEach-Object { [Uri]::EscapeDataString($_) } | 
            Join-String -Separator '/'
        $absoluteUrl = "$baseUrl/$relativePath"
        
        $libraryNode = [SharePointNode]::new($list.Title, "📚", $absoluteUrl, $Level)
        $libraryNode.ComplianceTag = Get-ComplianceTag -List $list
        $ParentNode.AddChild($libraryNode)
        
        Get-FolderHierarchy -List $list -FolderUrl $list.RootFolder.ServerRelativeUrl -Level ($Level + 1) -MaxDepth $MaxDepth -ParentNode $libraryNode
    }
}

# Function to recursively get subsites
function Get-SiteHierarchy {
    param (
        $SiteUrl,
        $Level = 0,
        [int]$MaxDepth = -1,
        [SharePointNode]$ParentNode,
        [switch]$IsRoot = $false
    )
    
    Connect-PnPOnline -Url $SiteUrl -ClientId $env:DigitalSupportAppID -Tenant 'mhud.onmicrosoft.com' -Thumbprint $env:DigitalSupportCertificateThumbprint
    
    # Only create a new site node if this is not the root call
    if (-not $IsRoot) {
        $siteNode = [SharePointNode]::new($SiteUrl, "🌐", $SiteUrl, $Level)
        $ParentNode.AddChild($siteNode)
        $currentNode = $siteNode
    } else {
        $currentNode = $ParentNode
    }
    
    Get-LibraryHierarchy -SiteUrl $SiteUrl -Level ($Level + 1) -MaxDepth $MaxDepth -ParentNode $currentNode
    
    # Get subsites
    $subsites = Get-PnPSubWeb
    foreach ($subsite in $subsites) {
        Get-SiteHierarchy -SiteUrl $subsite.Url -Level ($Level + 1) -MaxDepth $MaxDepth -ParentNode $currentNode
    }
}

function Add-HierarchyToExcel {
    param (
        $Level,
        $Name,
        $Type,
        $Url,
        $ComplianceTag,
        $ExcelData
    )
    
    Write-Verbose "Adding to Excel: $Name with label: $ComplianceTag"
    $item = [PSCustomObject]@{
        Level = "  " * $Level + $Type
        Name = $Name
        URL = $Url
        "Retention Label" = if ($ComplianceTag) { $ComplianceTag } else { "No Label" }
    }
    $ExcelData.Add($item) | Out-Null
}

function Add-HierarchyToHTML {
    param (
        $Level,
        $Name,
        $Type,
        $Url,
        $ComplianceTag,
        $StringBuilder,
        [bool]$HasChildren = $false
    )
    
    $typeClass = switch ($Type) {
        "🌐" { "site" }
        "📚" { "library" }
        "📁" { "folder" }
        default { "" }
    }
    
    $indent = $Level * 30
    
    # Create main item div
    $line = "<div class='item level-$Level $typeClass' style='margin-left: ${indent}px'>"
    
    # Add toggle button if has children
    if ($HasChildren) {
        $line += "<span class='toggle' onclick='toggleChildren(this)'></span>"
    } else {
        $line += "<span style='margin-left: 20px'></span>"
    }
    
    $line += "<span class='icon'>$Type</span>"
    if ($Url) {
        $line += "<a href='$Url' class='link' target='_blank'>$Name</a>"
    } else {
        $line += "<span class='name'>$Name</span>"
    }
    if ($ComplianceTag) {
        $line += "<span class='tag'>$ComplianceTag</span>"
    }
    $line += "</div>"
    
    $StringBuilder.AppendLine($line) | Out-Null
    
    if ($HasChildren) {
        $StringBuilder.AppendLine("<div class='children'>") | Out-Null
    }
}

# Add a helper function to close children containers
function Add-HierarchyToHTMLClose {
    param (
        $StringBuilder,
        [bool]$HasChildren = $false
    )
    
    if ($HasChildren) {
        $StringBuilder.AppendLine("</div>") | Out-Null
    }
}

function New-FreeplaneNode {
    param (
        $Name,
        $Type,
        $Url,
        $Level,
        $ComplianceTag
    )

    $nodes = New-Object System.Collections.ArrayList

    $node = [PSCustomObject]@{
        ID = [guid]::NewGuid().ToString("N")
        CREATED = [int][double]::Parse((Get-Date -UFormat %s))
        MODIFIED = [int][double]::Parse((Get-Date -UFormat %s))
        TEXT = "$Type $Name"
        LINK = $Url
        STYLE = "bubble"
        Nodes = $nodes  # Initialize with ArrayList
    }

    if ($ComplianceTag) {
        $node | Add-Member -NotePropertyName "ComplianceTag" -NotePropertyValue $ComplianceTag
    }

    return $node
}

function Export-ToFreeplane {
    param (
        $RootNode,
        $FilePath
    )
    
    try {
        # Ensure we have a full path
        $fullPath = if ([System.IO.Path]::IsPathRooted($FilePath)) {
            $FilePath
        } else {
            Join-Path $PSScriptRoot $FilePath
        }
        
        # Create directory if it doesn't exist
        $directory = [System.IO.Path]::GetDirectoryName($fullPath)
        if (-not (Test-Path $directory)) {
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
        }

        # Create the XML document with proper encoding
        $xmlSettings = New-Object System.Xml.XmlWriterSettings
        $xmlSettings.Indent = $true
        $xmlSettings.IndentChars = "  "
        $xmlSettings.Encoding = [System.Text.Encoding]::UTF8

        Write-Verbose "Creating XML file at: $fullPath"
        $xmlWriter = [System.XML.XmlWriter]::Create($fullPath, $xmlSettings)

        # Write map element with proper schema
        $xmlWriter.WriteStartElement("map")
        $xmlWriter.WriteAttributeString("version", "1.0.1")

        function Write-NodeToXml {
            param($Node, $XmlWriter, $Position = "right")
            
            $xmlWriter.WriteStartElement("node")
            $xmlWriter.WriteAttributeString("ID", $Node.ID)
            $xmlWriter.WriteAttributeString("CREATED", $Node.CREATED)
            $xmlWriter.WriteAttributeString("MODIFIED", $Node.MODIFIED)
            $xmlWriter.WriteAttributeString("TEXT", $Node.TEXT)
            $xmlWriter.WriteAttributeString("POSITION", $Position)
            
            if ($Node.LINK) {
                $xmlWriter.WriteAttributeString("LINK", $Node.LINK)
            }
            
            if ($Node.ComplianceTag) {
                $xmlWriter.WriteStartElement("attribute")
                $xmlWriter.WriteAttributeString("NAME", "RetentionLabel")
                $xmlWriter.WriteAttributeString("VALUE", $Node.ComplianceTag)
                $xmlWriter.WriteEndElement()
            }

            # Write child nodes with alternating positions
            $childPosition = "right"
            foreach ($childNode in $Node.Nodes) {
                Write-NodeToXml -Node $childNode -XmlWriter $xmlWriter -Position $childPosition
                $childPosition = if ($childPosition -eq "right") { "left" } else { "right" }
            }
            
            $xmlWriter.WriteEndElement()
        }

        Write-NodeToXml -Node $RootNode -XmlWriter $xmlWriter

        $xmlWriter.WriteEndElement() # map
        $xmlWriter.WriteEndDocument()
        $xmlWriter.Flush()
        $xmlWriter.Close()
        
        return $true
    }
    catch {
        Write-Error "Failed to create Freeplane mind map: $_"
        return $false
    }
}

function Get-HTMLHeader {
    return @"
<!DOCTYPE html>
<html>
<head>
    <title>SharePoint Site Hierarchy</title>
    <style>
        body { 
            font-family: Arial, sans-serif; 
            margin: 20px;
            background: #f8f9fa;
        }
        h1 { 
            color: #0078d4;
            margin-bottom: 10px;
        }
        .container {
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .item { 
            margin: 5px 0;
            padding: 4px 0;
            display: flex;
            align-items: center;
        }
        .icon { 
            display: inline-block;
            width: 25px;
            text-align: center;
        }
        .link { 
            color: #0078d4;
            text-decoration: none;
            margin-right: 10px;
        }
        .link:hover { 
            text-decoration: underline;
        }
        .tag { 
            color: #666;
            font-size: 0.9em;
            background: #f0f0f0;
            padding: 2px 6px;
            border-radius: 3px;
            margin-left: auto;
        }
        .timestamp { 
            color: #666;
            font-size: 0.9em;
            margin-bottom: 20px;
        }
        .level-0 { margin-left: 0; }
        .level-1 { margin-left: 30px; }
        .level-2 { margin-left: 60px; }
        .level-3 { margin-left: 90px; }
        .level-4 { margin-left: 120px; }
        .level-5 { margin-left: 150px; }
        .site { background-color: #f8f9fa; }
        .library { background-color: #e3f2fd; }
        .folder { background-color: #f1f8e9; }
        .toggle {
            cursor: pointer;
            user-select: none;
            margin-right: 5px;
        }
        
        .toggle:before {
            content: '▼';
            display: inline-block;
            margin-right: 5px;
            transition: transform 0.2s;
        }
        
        .toggle.collapsed:before {
            transform: rotate(-90deg);
        }
        
        .children {
            transition: height 0.2s ease-out;
            overflow: hidden;
        }
        
        .children.collapsed {
            display: none;
        }
    </style>
    <script>
        function toggleChildren(element) {
            const parent = element.parentElement;
            const children = parent.nextElementSibling;
            element.classList.toggle('collapsed');
            children.classList.toggle('collapsed');
        }
        
        function expandAll() {
            document.querySelectorAll('.toggle.collapsed').forEach(toggle => {
                toggle.click();
            });
        }
        
        function collapseAll() {
            document.querySelectorAll('.toggle:not(.collapsed)').forEach(toggle => {
                toggle.click();
            });
        }
    </script>
</head>
<body>
    <div class="container">
        <h1>SharePoint Site Hierarchy</h1>
        <div class="controls">
            <button onclick="expandAll()">Expand All</button>
            <button onclick="collapseAll()">Collapse All</button>
        </div>
        <div class="timestamp">Generated on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</div>
"@
}

function Get-OutputFileName {
    param (
        [string]$SiteUrl,
        [string]$Extension
    )
    
    # Extract site name from URL - get everything after /sites/
    $siteName = $SiteUrl -replace '.*\/sites\/', '' -replace '[^a-zA-Z0-9]', '_'
    $timestamp = Get-Date -Format "yyyyMMdd"
    
    # Create consistent filename format
    $fileName = "${siteName}_hierarchy_${timestamp}.$Extension"
    return Join-Path $OutputPath $fileName
}

function Export-ToHTML {
    param (
        [SharePointNode]$RootNode,
        [string]$OutputPath
    )
    
    $htmlPath = Get-OutputFileName -SiteUrl $RootNode.Url -Extension "html"
    $StringBuilder = [System.Text.StringBuilder]::new()
    
    $StringBuilder.AppendLine($(Get-HTMLHeader)) | Out-Null
    
    function Write-NodeToHTML {
        param($Node, $StringBuilder)
        
        $hasChildren = $Node.Children.Count -gt 0
        Add-HierarchyToHTML -Level $Node.Level -Name $Node.Name -Type $Node.Type `
            -Url $Node.Url -ComplianceTag $Node.ComplianceTag `
            -StringBuilder $StringBuilder -HasChildren $hasChildren
        
        foreach ($child in $Node.Children) {
            Write-NodeToHTML -Node $child -StringBuilder $StringBuilder
        }
        
        Add-HierarchyToHTMLClose -StringBuilder $StringBuilder -HasChildren $hasChildren
    }
    
    Write-NodeToHTML -Node $RootNode -StringBuilder $StringBuilder
    
    $StringBuilder.AppendLine("    </div>") | Out-Null
    $StringBuilder.AppendLine("</body>") | Out-Null
    $StringBuilder.AppendLine("</html>") | Out-Null
    
    $StringBuilder.ToString() | Out-File -FilePath $htmlPath -Encoding UTF8
    Write-Host "✓ HTML file generated successfully at: $htmlPath" -ForegroundColor Green
}

function Export-ToExcel {
    param (
        [SharePointNode]$RootNode,
        [string]$OutputPath,
        [System.Collections.ArrayList]$ExcelData
    )
    
    $excelPath = Get-OutputFileName -SiteUrl $RootNode.Url -Extension "xlsx"
    
    function Write-NodeToExcel {
        param($Node, $ExcelData)
        
        Add-HierarchyToExcel -Level $Node.Level -Name $Node.Name -Type $Node.Type `
            -Url $Node.Url -ComplianceTag $Node.ComplianceTag -ExcelData $ExcelData
        
        foreach ($child in $Node.Children) {
            Write-NodeToExcel -Node $child -ExcelData $ExcelData
        }
    }
    
    Write-NodeToExcel -Node $RootNode -ExcelData $ExcelData
    
    $ExcelData | Export-Excel -Path $excelPath -AutoSize -FreezeTopRow -BoldTopRow `
        -WorksheetName "SharePoint Hierarchy" -TableStyle Medium9
    
    Write-Host "✓ Excel file generated successfully at: $excelPath" -ForegroundColor Green
}

function Export-ToMarkdown {
    param (
        [SharePointNode]$RootNode,
        [string]$OutputPath
    )
    
    # Get site name from RootNode URL instead of global $SiteUrl
    $siteName = $RootNode.Url -replace 'https://\w+\.sharepoint\.com/sites/', '' -replace '[^a-zA-Z0-9]', '_'
    $timestamp = Get-Date -Format "yyyyMMdd"
    $mdPath = Join-Path $OutputPath "$($siteName)_hierarchy_$($timestamp).md"
    $StringBuilder = [System.Text.StringBuilder]::new()
    
    $StringBuilder.AppendLine("# SharePoint Site Hierarchy") | Out-Null
    $StringBuilder.AppendLine("Generated on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')") | Out-Null
    $StringBuilder.AppendLine("") | Out-Null
    
    function Write-NodeToMarkdown {
        param($Node, $StringBuilder)
        
        # Remove this check to allow root node to be written
        $indent = "    " * $Node.Level
        $line = "$indent- $($Node.Type) [$($Node.Name)]($($Node.Url))"
        if ($Node.ComplianceTag -ne "No Label") {
            $line += " [🏷️ $($Node.ComplianceTag)]"
        }
        $StringBuilder.AppendLine($line) | Out-Null
        
        foreach ($child in $Node.Children) {
            Write-NodeToMarkdown -Node $child -StringBuilder $StringBuilder
        }
    }
    
    Write-NodeToMarkdown -Node $RootNode -StringBuilder $StringBuilder
    
    $StringBuilder.ToString() | Out-File -FilePath $mdPath -Encoding UTF8
    Write-Host "✓ Markdown file generated successfully at: $mdPath" -ForegroundColor Green
}

try {
    Write-Host "=== SharePoint Hierarchy Generation ===" -ForegroundColor Cyan
    
    # Create root node and collect all data
    $rootNode = [SharePointNode]::new($SiteUrl, "🌐", $SiteUrl, 0)
    Write-Host "`nCollecting SharePoint hierarchy data..."
    Get-SiteHierarchy -SiteUrl $SiteUrl -Level 0 -MaxDepth $MaxDepth -ParentNode $rootNode -IsRoot
    
    Write-Host "`nData collection complete. Choose export format(s):"
    Write-Host "1. Markdown"
    Write-Host "2. Excel"
    Write-Host "3. XML (Mind Map)"
    Write-Host "4. HTML"
    Write-Host "5. All formats"
    Write-Host "6. Exit without export"

    $choice = Read-Host "`nEnter your choice (multiple choices separated by comma)"
    $choices = $choice.Split(',').Trim()

    foreach ($c in $choices) {
        switch ($c) {
            "1" { Export-ToMarkdown -RootNode $rootNode -OutputPath $OutputPath }
            "2" { 
                $excelData = [System.Collections.ArrayList]::new()
                Export-ToExcel -RootNode $rootNode -OutputPath $OutputPath -ExcelData $excelData 
            }
            "3" { Export-ToFreeplane -RootNode $rootNode -FilePath (Get-OutputFileName -SiteUrl $SiteUrl -Extension "mm") }
            "4" { Export-ToHTML -RootNode $rootNode -OutputPath $OutputPath }
            "5" { 
                Export-ToMarkdown -RootNode $rootNode -OutputPath $OutputPath
                $excelData = [System.Collections.ArrayList]::new()
                Export-ToExcel -RootNode $rootNode -OutputPath $OutputPath -ExcelData $excelData
                Export-ToFreeplane -RootNode $rootNode -FilePath (Get-OutputFileName -SiteUrl $SiteUrl -Extension "mm")
                Export-ToHTML -RootNode $rootNode -OutputPath $OutputPath
            }
            "6" { Write-Host "Exiting without export." }
            default { Write-Warning "Invalid choice: $c" }
        }
    }
}
catch {
    Write-Error "An error occurred: $_"
}
finally {
    Disconnect-PnPOnline
}