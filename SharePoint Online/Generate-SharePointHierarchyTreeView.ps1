[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "C:\HUD\06_Reporting",

    [Parameter(Mandatory = $false)]
    [switch]$GenerateMarkdown,

    [Parameter(Mandatory = $false)]
    [switch]$GenerateExcel,

    [Parameter(Mandatory = $false)]
    [int]$MaxDepth = -1,

    [Parameter(Mandatory = $false)]
    [switch]$GenerateXML,

    [Parameter(Mandatory = $false)]
    [switch]$GenerateHTML
)

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
            $label = Get-PnPLabel -List $List.Title -ValuesOnly
            if ($null -ne $label) {
                Write-Host "  ✓ Found label for $($List.Title): $($label.TagName)" -ForegroundColor Green
                return $label.TagName
            }
            Write-Host "  ℹ No label found for $($List.Title)" -ForegroundColor Gray
        }
        if ($FolderPath) {
            Write-Verbose "Checking inherited label"
            $parentLabel = Get-PnPLabel -List $List.Title -ValuesOnly
            if ($null -ne $parentLabel) {
                Write-Host "  ✓ Found inherited label: $($parentLabel.TagName)" -ForegroundColor Green
                return "$($parentLabel.TagName) (Inherited)"
            }
        }
    }
    catch {
        Write-Warning "  ! Error getting label: $_"
    }
    return "No Label"
}

# Function to recursively get folders and files
function Get-FolderHierarchy {
    param (
        $List,
        $FolderUrl = "",
        $Level = 0,
        [switch]$Markdown,
        $StringBuilder,
        [int]$MaxDepth = -1,
        [PSCustomObject]$ParentNode
    )
    
    try {
        Write-Host "`n=== Processing Folder Structure ===" -ForegroundColor Cyan
        Write-Host "  List: $($List.Title)" -ForegroundColor White
        Write-Host "  Current Level: $Level" -ForegroundColor Gray
        
        if ($MaxDepth -ne -1 -and $Level -ge $MaxDepth) {
            return
        }

        $folders = if ([string]::IsNullOrEmpty($FolderUrl)) {
            Write-Verbose "  Getting root folders..."
            Get-PnPFolder -ListRootFolder $List | Get-PnPFolderInFolder | 
                Where-Object { $_.Name -notmatch '^_|^Forms$' }
        } else {
            Write-Verbose "  Getting folders from: $FolderUrl"
            Get-PnPFolder -Url $FolderUrl | Get-PnPFolderInFolder |
                Where-Object { $_.Name -notmatch '^_|^Forms$' }
        }
        
        Write-Host "  Found $($folders.Count) folders" -ForegroundColor Yellow
        
        foreach ($folder in $folders) {
            $indentString = Get-IndentString -level $Level -Markdown:$Markdown
            if ($Markdown) {
                # Create URL-friendly path with encoded spaces
                $baseUrl = [regex]::Match($List.Context.Url, 'https://[^/]+').Value
                $relativePath = $folder.ServerRelativeUrl.Split('/') | 
                    Where-Object { $_ } | 
                    ForEach-Object { [Uri]::EscapeDataString($_) } | 
                    Join-String -Separator '/'
                $absoluteUrl = "$baseUrl/$relativePath"
                $folderLine = "$indentString📁 [$($folder.Name)]($absoluteUrl)"
                $StringBuilder.AppendLine($folderLine) | Out-Null
            } elseif ($GenerateExcel) {
                $baseUrl = [regex]::Match($List.Context.Url, 'https://[^/]+').Value
                $relativePath = $folder.ServerRelativeUrl.Split('/') | 
                    Where-Object { $_ } | 
                    ForEach-Object { [Uri]::EscapeDataString($_) } | 
                    Join-String -Separator '/'
                $absoluteUrl = "$baseUrl/$relativePath"
                $complianceTag = Get-ComplianceTag -List $List -FolderPath $folder.ServerRelativeUrl
                Add-HierarchyToExcel -Level $Level -Name $folder.Name -Type "📁" -Url $absoluteUrl -ComplianceTag $complianceTag -ExcelData $script:ExcelData
            } elseif ($GenerateXML) {
                $baseUrl = [regex]::Match($List.Context.Url, 'https://[^/]+').Value
                $relativePath = $folder.ServerRelativeUrl.Split('/') | 
                    Where-Object { $_ } | 
                    ForEach-Object { [Uri]::EscapeDataString($_) } | 
                    Join-String -Separator '/'
                $absoluteUrl = "$baseUrl/$relativePath"
                $complianceTag = Get-ComplianceTag -List $List -FolderPath $folder.ServerRelativeUrl
                $folderNode = New-FreeplaneNode -Name $folder.Name -Type "📁" -Url $absoluteUrl -Level $Level -ComplianceTag $complianceTag
                $ParentNode.Nodes.Add($folderNode) | Out-Null
            } elseif ($GenerateHTML) {
                $baseUrl = [regex]::Match($List.Context.Url, 'https://[^/]+').Value
                $relativePath = $folder.ServerRelativeUrl.Split('/') | 
                    Where-Object { $_ } | 
                    ForEach-Object { [Uri]::EscapeDataString($_) } | 
                    Join-String -Separator '/'
                $absoluteUrl = "$baseUrl/$relativePath"
                $complianceTag = Get-ComplianceTag -List $List -FolderPath $folder.ServerRelativeUrl
                Add-HierarchyToHTML -Level $Level -Name $folder.Name -Type "📁" -Url $absoluteUrl -ComplianceTag $complianceTag -StringBuilder $script:StringBuilder
            } else {
                $folderLine = "$indentString 📁 $($folder.Name)"
                Write-Host $folderLine -ForegroundColor Yellow
            }
            
            Get-FolderHierarchy -List $List -FolderUrl $folder.ServerRelativeUrl -Level ($Level + 1) -Markdown:$Markdown -StringBuilder $StringBuilder -MaxDepth $MaxDepth -ParentNode $ParentNode
        }
    }
    catch {
        Write-Warning "Error accessing folder in $($List.Title) : $_"
        Write-Verbose $_.Exception.Message
    }
}

# Function to get document libraries
function Get-LibraryHierarchy {
    param (
        $SiteUrl,
        $Level = 0,
        [switch]$Markdown,
        $StringBuilder,
        [int]$MaxDepth = -1,
        [PSCustomObject]$ParentNode
    )
    
    Write-Host "`n=== Processing Document Libraries ===" -ForegroundColor Cyan
    Write-Host "  Site: $SiteUrl" -ForegroundColor White
    Write-Host "  Level: $Level" -ForegroundColor Gray
    
    Connect-PnPOnline -Url $SiteUrl -ClientId $env:DigitalSupportAppID -Tenant 'mhud.onmicrosoft.com' -Thumbprint $env:DigitalSupportCertificateThumbprint
    $lists = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and $_.Hidden -eq $false }
    Write-Host "  Found $($lists.Count) document libraries" -ForegroundColor Yellow
    
    foreach ($list in $lists) {
        $indentString = Get-IndentString -level $Level -Markdown:$Markdown
        if ($Markdown) {
            # Create URL-friendly path with encoded spaces
            $baseUrl = [regex]::Match($SiteUrl, 'https://[^/]+').Value
            $relativePath = $list.RootFolder.ServerRelativeUrl.Split('/') |
                Where-Object { $_ } |
                ForEach-Object { [Uri]::EscapeDataString($_) } |
                Join-String -Separator '/'
            $absoluteUrl = "$baseUrl/$relativePath"
            $libraryLine = "$indentString📚 [$($list.Title)]($absoluteUrl)"
            $StringBuilder.AppendLine($libraryLine) | Out-Null
        } elseif ($GenerateExcel) {
            $baseUrl = [regex]::Match($SiteUrl, 'https://[^/]+').Value
            $relativePath = $list.RootFolder.ServerRelativeUrl.Split('/') | 
                Where-Object { $_ } | 
                ForEach-Object { [Uri]::EscapeDataString($_) } | 
                Join-String -Separator '/'
            $absoluteUrl = "$baseUrl/$relativePath"
            
            Write-Verbose "Getting compliance tag for library: $($list.Title)"
            $complianceTag = Get-ComplianceTag -List $list
            Write-Verbose "Compliance tag for $($list.Title): $complianceTag"
            
            Add-HierarchyToExcel -Level $Level -Name $list.Title -Type "📚" `
                -Url $absoluteUrl -ComplianceTag $complianceTag -ExcelData $script:ExcelData
        } elseif ($GenerateXML) {
            $baseUrl = [regex]::Match($SiteUrl, 'https://[^/]+').Value
            $relativePath = $list.RootFolder.ServerRelativeUrl.Split('/') | 
                Where-Object { $_ } | 
                ForEach-Object { [Uri]::EscapeDataString($_) } | 
                Join-String -Separator '/'
            $absoluteUrl = "$baseUrl/$relativePath"
            $complianceTag = Get-ComplianceTag -List $list
            $libraryNode = New-FreeplaneNode -Name $list.Title -Type "📚" -Url $absoluteUrl -Level $Level -ComplianceTag $complianceTag
            $ParentNode.Nodes.Add($libraryNode) | Out-Null
            Get-FolderHierarchy -List $list -FolderUrl $list.RootFolder.ServerRelativeUrl -Level ($Level + 1) -Markdown:$Markdown -StringBuilder $StringBuilder -MaxDepth $MaxDepth -ParentNode $libraryNode
        } elseif ($GenerateHTML) {
            $baseUrl = [regex]::Match($SiteUrl, 'https://[^/]+').Value
            $relativePath = $list.RootFolder.ServerRelativeUrl.Split('/') | 
                Where-Object { $_ } | 
                ForEach-Object { [Uri]::EscapeDataString($_) } | 
                Join-String -Separator '/'
            $absoluteUrl = "$baseUrl/$relativePath"
            $complianceTag = Get-ComplianceTag -List $list
            Add-HierarchyToHTML -Level $Level -Name $list.Title -Type "📚" -Url $absoluteUrl -ComplianceTag $complianceTag -StringBuilder $script:StringBuilder
        } else {
            $libraryLine = "$indentString 📚 $($list.Title)"
            Write-Host $libraryLine
        }
        
        Get-FolderHierarchy -List $list -FolderUrl $list.RootFolder.ServerRelativeUrl -Level ($Level + 1) -Markdown:$Markdown -StringBuilder $StringBuilder -MaxDepth $MaxDepth -ParentNode $ParentNode
    }
}

# Function to recursively get subsites
function Get-SiteHierarchy {
    param (
        $SiteUrl,
        $Level = 0,
        [switch]$Markdown,
        $StringBuilder,
        [int]$MaxDepth = -1,
        [PSCustomObject]$ParentNode
    )
    
    Connect-PnPOnline -Url $SiteUrl -ClientId $env:DigitalSupportAppID -Tenant 'mhud.onmicrosoft.com' -Thumbprint $env:DigitalSupportCertificateThumbprint
    $indentString = Get-IndentString -level $Level -Markdown:$Markdown
    if ($Markdown) {
        $siteLine = "$indentString🌐 [$SiteUrl]($SiteUrl)"
        $StringBuilder.AppendLine($siteLine) | Out-Null
    } elseif ($GenerateExcel) {
        Add-HierarchyToExcel -Level $Level -Name $SiteUrl -Type "🌐" -Url $SiteUrl -ExcelData $script:ExcelData
    } elseif ($GenerateXML) {
        $siteNode = New-FreeplaneNode -Name $SiteUrl -Type "🌐" -Url $SiteUrl -Level $Level
        if ($ParentNode) {
            $ParentNode.Nodes.Add($siteNode) | Out-Null
            Get-LibraryHierarchy -SiteUrl $SiteUrl -Level ($Level + 1) -ParentNode $siteNode -MaxDepth $MaxDepth
        } else {
            # This is the root node
            $script:rootNode = $siteNode
            Get-LibraryHierarchy -SiteUrl $SiteUrl -Level ($Level + 1) -ParentNode $siteNode -MaxDepth $MaxDepth
        }
        return $siteNode
    } elseif ($GenerateHTML) {
        Add-HierarchyToHTML -Level $Level -Name $SiteUrl -Type "🌐" -Url $SiteUrl -StringBuilder $script:StringBuilder
    } else {
        $siteLine = "$indentString 🌐 $SiteUrl"
        Write-Host $siteLine
    }
    
    Get-LibraryHierarchy -SiteUrl $SiteUrl -Level ($Level + 1) -Markdown:$Markdown -StringBuilder $StringBuilder -MaxDepth $MaxDepth -ParentNode $ParentNode
    
    # Get subsites
    $subsites = Get-PnPSubWeb
    foreach ($subsite in $subsites) {
        Get-SiteHierarchy -SiteUrl $subsite.Url -Level ($Level + 1) -Markdown:$Markdown -StringBuilder $StringBuilder -MaxDepth $MaxDepth -ParentNode $ParentNode   
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
        $StringBuilder
    )
    
    # Determine item type class
    $typeClass = switch ($Type) {
        "🌐" { "site" }
        "📚" { "library" }
        "📁" { "folder" }
        default { "" }
    }
    
    $line = "<div class='item level-$Level $typeClass'>"
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
    </style>
</head>
<body>
    <div class="container">
        <h1>SharePoint Site Hierarchy</h1>
        <div class="timestamp">Generated on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</div>
"@
}

function Get-OutputFileName {
    param (
        [string]$SiteUrl,
        [string]$Extension
    )
    
    # Extract site name from URL
    $siteName = $SiteUrl -replace 'https://', '' -replace '\.sharepoint\.com.*', ''
    $date = Get-Date -Format 'yyyyMMdd'
    
    # Create the full path
    $fileName = "${siteName}_${date}_treeview.$Extension"
    return Join-Path "C:\HUD\06_Reporting" $fileName
}

# Main script
try {
    Write-Host "=== SharePoint Hierarchy Generation ===" -ForegroundColor Cyan
    Write-Host "Started at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor White
    Write-Host "Output Path: $OutputPath" -ForegroundColor Gray
    Write-Host "Site URL: $SiteUrl" -ForegroundColor White
    Write-Host "Maximum Depth: $(if ($MaxDepth -eq -1) { 'Unlimited' } else { $MaxDepth })" -ForegroundColor Gray
    Write-Host "`nInitializing..." -ForegroundColor Yellow
    
    if (-not $SiteUrl) {
        $SiteUrl = Read-Host "Enter your SharePoint tenant URL (e.g., https://contoso.sharepoint.com)"
    }
    Write-Verbose "Site URL: $SiteUrl"

    # Ensure output directory exists
    if (-not (Test-Path $OutputPath)) {
        Write-Verbose "Creating output directory: $OutputPath"
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }

    if ($GenerateMarkdown) {
        Write-Host "`n=== Generating Markdown Output ===" -ForegroundColor Cyan
        $markdownPath = Get-OutputFileName -SiteUrl $SiteUrl -Extension "md"
        $StringBuilder = [System.Text.StringBuilder]::new()
        $StringBuilder.AppendLine("# SharePoint Site Hierarchy") | Out-Null
        $StringBuilder.AppendLine("Generated on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')") | Out-Null
        $StringBuilder.AppendLine("") | Out-Null
        
        Write-Host "Generating SharePoint Hierarchy Tree as Markdown..."
        Get-SiteHierarchy -SiteUrl $SiteUrl -Level 0 -Markdown -StringBuilder $StringBuilder -MaxDepth $MaxDepth
        
        $StringBuilder.ToString() | Out-File -FilePath $markdownPath -Encoding UTF8
        Write-Host "✓ Markdown file generated successfully at: $markdownPath" -ForegroundColor Green
    } elseif ($GenerateExcel) {
        Write-Host "`n=== Generating Excel Output ===" -ForegroundColor Cyan
        $excelPath = Get-OutputFileName -SiteUrl $SiteUrl -Extension "xlsx"
        $script:ExcelData = [System.Collections.ArrayList]::new()
        Write-Host "`nGenerating SharePoint Hierarchy Tree as Excel...`n"
        Get-SiteHierarchy -SiteUrl $SiteUrl -Level 0 -MaxDepth $MaxDepth
        
        $script:ExcelData | Export-Excel -Path $excelPath -AutoSize -FreezeTopRow -BoldTopRow `
            -WorksheetName "SharePoint Hierarchy" -TableStyle Medium9
        Write-Host "✓ Excel file generated successfully at: $excelPath" -ForegroundColor Green
    } elseif ($GenerateXML) {
        Write-Host "`n=== Generating XML Mind Map ===" -ForegroundColor Cyan
        $xmlPath = Get-OutputFileName -SiteUrl $SiteUrl -Extension "mm"
        Write-Host "`nGenerating SharePoint Hierarchy as Freeplane Mind Map...`n"
        $script:rootNode = $null
        
        # Ensure XML path is properly resolved
        $xmlFullPath = if ([System.IO.Path]::IsPathRooted($xmlPath)) {
            $xmlPath
        } else {
            Join-Path $PSScriptRoot $xmlPath
        }
        
        Get-SiteHierarchy -SiteUrl $SiteUrl -Level 0 -MaxDepth $MaxDepth
        
        if ($script:rootNode) {
            if (Export-ToFreeplane -RootNode $script:rootNode -FilePath $xmlFullPath) {
                Write-Host "✓ Mind map generated successfully at: $xmlFullPath" -ForegroundColor Green
            }
        } else {
            Write-Error "Failed to generate mind map - no root node created"
        }
    } elseif ($GenerateHTML) {
        Write-Host "`n=== Generating HTML Output ===" -ForegroundColor Cyan
        $htmlPath = Get-OutputFileName -SiteUrl $SiteUrl -Extension "html"
        $script:StringBuilder = [System.Text.StringBuilder]::new()
        
        # Add HTML header with styles
        $script:StringBuilder.AppendLine($(Get-HTMLHeader)) | Out-Null
        
        Write-Host "Generating SharePoint Hierarchy as HTML..."
        Get-SiteHierarchy -SiteUrl $SiteUrl -Level 0 -MaxDepth $MaxDepth
        
        # Close HTML containers and document
        $script:StringBuilder.AppendLine("    </div>") | Out-Null
        $script:StringBuilder.AppendLine("</body>") | Out-Null
        $script:StringBuilder.AppendLine("</html>") | Out-Null
        
        $script:StringBuilder.ToString() | Out-File -FilePath $htmlPath -Encoding UTF8
        Write-Host "✓ HTML file generated successfully at: $htmlPath" -ForegroundColor Green
    } else {
        Write-Host "`n=== Generating Console Output ===" -ForegroundColor Cyan
        Get-SiteHierarchy -SiteUrl $SiteUrl -Level 0 -MaxDepth $MaxDepth
    }
    
    Write-Host "`n=== Operation Complete ===" -ForegroundColor Green
    Write-Host "Finished at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor White
}
catch {
    Write-Error "An error occurred: $_"
}
finally {
    Disconnect-PnPOnline
}
