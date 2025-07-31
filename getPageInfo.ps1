# SharePoint 2016 Page Information Extractor (Simplified)
# Run Connect-SharePoint.ps1 first to authenticate

param(
    [Parameter(Mandatory=$false)]
    [string]$PageUrl = "SitePages/stars.aspx",
    
    [Parameter(Mandatory=$false)]
    [string]$OutputFolder = "C:\Users\cmarala\Desktop\Ford\Output"
)

Write-Host "=== SharePoint Page Information Extractor ===" -ForegroundColor Cyan
Write-Host "Page: $PageUrl" -ForegroundColor Yellow

# Check if already connected
try {
    $context = Get-PnPContext -ErrorAction Stop
    if ($null -eq $context) {
        throw "No active connection"
    }
    Write-Host "Using existing SharePoint connection" -ForegroundColor Green
}
catch {
    Write-Error "No active SharePoint connection found."
    Write-Host "Please run Connect-SharePoint.ps1 first to authenticate." -ForegroundColor Red
    exit 1
}

# Create output folder if it doesn't exist
if (!(Test-Path $OutputFolder)) { 
    New-Item -Path $OutputFolder -ItemType Directory -Force
    Write-Host "Created output folder: $OutputFolder" -ForegroundColor Green
}

try {
    # Get the web information
    Write-Host "`nGetting web information..." -ForegroundColor Yellow
    $web = Get-PnPWeb
    Write-Host "Connected to: $($web.Title)" -ForegroundColor Green
    
    # Get the page file
    Write-Host "`nGetting page file..." -ForegroundColor Yellow
    $pageFile = Get-PnPFile -Url $PageUrl -ErrorAction Stop
    Write-Host "Page file: $($pageFile.Name)" -ForegroundColor Green
    
    # Get the page as list item
    Write-Host "`nGetting page list item..." -ForegroundColor Yellow
    $page = Get-PnPFile -Url $PageUrl -AsListItem -ErrorAction Stop
    Write-Host "Page title: $($page['Title'])" -ForegroundColor Green
    
    # Construct server relative URL for web parts
    $serverRelativePageUrl = $pageFile.ServerRelativeUrl
    if (-not $serverRelativePageUrl) {
        if ($web.ServerRelativeUrl -eq "/") {
            $serverRelativePageUrl = "/$PageUrl"
        } else {
            $serverRelativePageUrl = "$($web.ServerRelativeUrl)/$PageUrl"
        }
    }
    Write-Host "Server relative URL: $serverRelativePageUrl" -ForegroundColor Cyan
    
    # Get all web parts
    Write-Host "`nGetting all web parts..." -ForegroundColor Yellow
    $webParts = Get-PnPWebPart -ServerRelativePageUrl $serverRelativePageUrl
    Write-Host "Found $($webParts.Count) web parts" -ForegroundColor Green
    
    # Create XML document
    $xmlDoc = New-Object System.Xml.XmlDocument
    $xmlDeclaration = $xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", $null)
    $xmlDoc.AppendChild($xmlDeclaration) | Out-Null
    
    # Root element
    $rootElement = $xmlDoc.CreateElement("SharePointPageInfo")
    $xmlDoc.AppendChild($rootElement) | Out-Null
    
    # Page Properties
    Write-Host "`nExtracting page properties..." -ForegroundColor Yellow
    $pagePropsElement = $xmlDoc.CreateElement("PageProperties")
    $rootElement.AppendChild($pagePropsElement) | Out-Null
    
    # Basic page info
    $pagePropsElement.AppendChild($xmlDoc.CreateElement("PageUrl")).InnerText = $PageUrl
    $pagePropsElement.AppendChild($xmlDoc.CreateElement("ServerRelativeUrl")).InnerText = $serverRelativePageUrl
    $pagePropsElement.AppendChild($xmlDoc.CreateElement("SiteTitle")).InnerText = $web.Title
    $pagePropsElement.AppendChild($xmlDoc.CreateElement("SiteUrl")).InnerText = $web.Url
    
    if ($pageFile.Name) {
        $pagePropsElement.AppendChild($xmlDoc.CreateElement("FileName")).InnerText = $pageFile.Name
    }
    if ($pageFile.Length) {
        $pagePropsElement.AppendChild($xmlDoc.CreateElement("FileSize")).InnerText = $pageFile.Length.ToString()
    }
    if ($pageFile.TimeCreated) {
        $pagePropsElement.AppendChild($xmlDoc.CreateElement("FileCreated")).InnerText = $pageFile.TimeCreated.ToString("yyyy-MM-ddTHH:mm:ssZ")
    }
    if ($pageFile.TimeLastModified) {
        $pagePropsElement.AppendChild($xmlDoc.CreateElement("FileModified")).InnerText = $pageFile.TimeLastModified.ToString("yyyy-MM-ddTHH:mm:ssZ")
    }
    
    # Page list item properties
    if ($page) {
        if ($page["Title"]) {
            $pagePropsElement.AppendChild($xmlDoc.CreateElement("PageTitle")).InnerText = $page["Title"].ToString()
        }
        if ($page["ID"]) {
            $pagePropsElement.AppendChild($xmlDoc.CreateElement("PageID")).InnerText = $page["ID"].ToString()
        }
        if ($page["Created"]) {
            $pagePropsElement.AppendChild($xmlDoc.CreateElement("Created")).InnerText = ([DateTime]$page["Created"]).ToString("yyyy-MM-ddTHH:mm:ssZ")
        }
        if ($page["Modified"]) {
            $pagePropsElement.AppendChild($xmlDoc.CreateElement("Modified")).InnerText = ([DateTime]$page["Modified"]).ToString("yyyy-MM-ddTHH:mm:ssZ")
        }
        if ($page["Author"] -and $page["Author"].LookupValue) {
            $pagePropsElement.AppendChild($xmlDoc.CreateElement("CreatedBy")).InnerText = $page["Author"].LookupValue
        }
        if ($page["Editor"] -and $page["Editor"].LookupValue) {
            $pagePropsElement.AppendChild($xmlDoc.CreateElement("ModifiedBy")).InnerText = $page["Editor"].LookupValue
        }
    }
    
    # Web Parts
    Write-Host "`nExtracting web parts..." -ForegroundColor Yellow
    $webPartsElement = $xmlDoc.CreateElement("WebParts")
    $webPartsElement.SetAttribute("Count", $webParts.Count.ToString())
    $rootElement.AppendChild($webPartsElement) | Out-Null
    
    $counter = 0
    foreach ($wp in $webParts) {
        $counter++
        Write-Host "Processing web part $counter of $($webParts.Count): $($wp.WebPart.Title)" -ForegroundColor Cyan
        
        $webPartElement = $xmlDoc.CreateElement("WebPart")
        $webPartElement.SetAttribute("Index", $counter.ToString())
        $webPartsElement.AppendChild($webPartElement) | Out-Null
        
        # Basic web part properties
        if ($wp.Id) {
            $webPartElement.AppendChild($xmlDoc.CreateElement("Id")).InnerText = $wp.Id.ToString()
        }
        
        # Zone Information
        $zoneInfoElement = $xmlDoc.CreateElement("ZoneInformation")
        $webPartElement.AppendChild($zoneInfoElement) | Out-Null
        
        if ($wp.ZoneId) {
            $zoneInfoElement.AppendChild($xmlDoc.CreateElement("ZoneId")).InnerText = $wp.ZoneId
        }
        if ($wp.ZoneIndex -ne $null) {
            $zoneInfoElement.AppendChild($xmlDoc.CreateElement("ZoneIndex")).InnerText = $wp.ZoneIndex.ToString()
        }
        
        # Web part object properties
        if ($wp.WebPart) {
            if ($wp.WebPart.Title) {
                $webPartElement.AppendChild($xmlDoc.CreateElement("Title")).InnerText = $wp.WebPart.Title
            }
            if ($wp.WebPart.TitleUrl) {
                $webPartElement.AppendChild($xmlDoc.CreateElement("TitleUrl")).InnerText = $wp.WebPart.TitleUrl
            }
            if ($wp.WebPart.ExportMode) {
                $webPartElement.AppendChild($xmlDoc.CreateElement("ExportMode")).InnerText = $wp.WebPart.ExportMode.ToString()
            }
        }
        
        # Get all web part properties using Get-PnPWebPartProperty
        Write-Host "  Getting all web part properties..." -ForegroundColor Gray
        try {
            $webPartProperties = Get-PnPWebPartProperty -ServerRelativePageUrl $serverRelativePageUrl -Identity $wp.Id
            
            if ($webPartProperties) {
                Write-Host "  Successfully retrieved $($webPartProperties.Count) properties" -ForegroundColor Green
                
                $propertiesElement = $xmlDoc.CreateElement("WebPartProperties")
                $webPartElement.AppendChild($propertiesElement) | Out-Null
                
                # Process each property
                foreach ($property in $webPartProperties.GetEnumerator()) {
                    try {
                        $propElement = $xmlDoc.CreateElement("Property")
                        $propElement.SetAttribute("Name", $property.Key)
                        $propertiesElement.AppendChild($propElement) | Out-Null
                        
                        $propValue = $property.Value
                        if ($propValue -ne $null) {
                            # Handle different property types
                            if ($propValue -is [System.String]) {
                                $propElement.SetAttribute("Type", "String")
                                $propElement.InnerText = $propValue
                            }
                            elseif ($propValue -is [System.Boolean]) {
                                $propElement.SetAttribute("Type", "Boolean")
                                $propElement.InnerText = $propValue.ToString().ToLower()
                            }
                            elseif ($propValue -is [System.Int32] -or $propValue -is [System.Int64]) {
                                $propElement.SetAttribute("Type", "Integer")
                                $propElement.InnerText = $propValue.ToString()
                            }
                            elseif ($propValue -is [System.Guid]) {
                                $propElement.SetAttribute("Type", "Guid")
                                $propElement.InnerText = $propValue.ToString()
                            }
                            elseif ($propValue -is [System.Enum]) {
                                $propElement.SetAttribute("Type", "Enum")
                                $propElement.InnerText = $propValue.ToString()
                            }
                            else {
                                $propElement.SetAttribute("Type", $propValue.GetType().Name)
                                $propElement.InnerText = $propValue.ToString()
                            }
                        }
                        else {
                            $propElement.SetAttribute("Type", "Null")
                            $propElement.InnerText = ""
                        }
                    }
                    catch {
                        Write-Warning "    Error processing property '$($property.Key)': $($_.Exception.Message)"
                        $propElement = $xmlDoc.CreateElement("Property")
                        $propElement.SetAttribute("Name", $property.Key)
                        $propElement.SetAttribute("Type", "Error")
                        $propElement.InnerText = "Error: $($_.Exception.Message)"
                        $propertiesElement.AppendChild($propElement) | Out-Null
                    }
                }
            }
            else {
                Write-Warning "  Could not retrieve properties for web part: $($wp.WebPart.Title)"
                $webPartElement.AppendChild($xmlDoc.CreateElement("PropertiesStatus")).InnerText = "Could not retrieve properties"
            }
        }
        catch {
            Write-Warning "  Error getting properties for web part: $($_.Exception.Message)"
            $webPartElement.AppendChild($xmlDoc.CreateElement("PropertiesError")).InnerText = $_.Exception.Message
        }
        
        # Get web part XML using Get-PnPWebPartXml
        Write-Host "  Getting XML for web part: $($wp.WebPart.Title)" -ForegroundColor Gray
        try {
            $webPartXml = Get-PnPWebPartXml -ServerRelativePageUrl $serverRelativePageUrl -Identity $wp.Id
            
            if ($webPartXml) {
                Write-Host "  Successfully retrieved XML ($($webPartXml.Length) characters)" -ForegroundColor Green
                
                # Create XML Information section
                $xmlInfoElement = $xmlDoc.CreateElement("XMLInformation")
                $webPartElement.AppendChild($xmlInfoElement) | Out-Null
                
                # Save full XML
                $xmlInfoElement.AppendChild($xmlDoc.CreateElement("FullXML")).InnerText = $webPartXml
                
                # Parse XML to get type and content
                $xmlDoc2 = New-Object System.Xml.XmlDocument
                $xmlDoc2.LoadXml($webPartXml)
                
                # Detect web part format and type
                $webPartFormat = "Unknown"
                $actualTypeName = $null
                $partOrder = $null
                $zoneIdFromXml = $null
                
                # Check for v2 format (ContentEditor style) - root element WebPart with v2 namespace
                $v2RootNodes = $xmlDoc2.SelectNodes("//*[local-name()='WebPart' and namespace-uri()='http://schemas.microsoft.com/WebPart/v2']")
                if ($v2RootNodes.Count -gt 0) {
                    $webPartFormat = "WebPart_v2"
                    $typeNameNodes = $xmlDoc2.SelectNodes("//TypeName")
                    if ($typeNameNodes.Count -gt 0) {
                        $actualTypeName = $typeNameNodes[0].InnerText
                        Write-Host "  Found v2 WebPart: $actualTypeName" -ForegroundColor Green
                    }
                    # Extract PartOrder and ZoneID from v2 XML
                    $partOrderNodes = $xmlDoc2.SelectNodes("//PartOrder")
                    if ($partOrderNodes.Count -gt 0) {
                        $partOrder = $partOrderNodes[0].InnerText
                    }
                    $zoneIdNodes = $xmlDoc2.SelectNodes("//ZoneID")
                    if ($zoneIdNodes.Count -gt 0) {
                        $zoneIdFromXml = $zoneIdNodes[0].InnerText
                    }
                }
                # Check for v3 format (ListView style) - webParts/webPart structure
                elseif ($xmlDoc2.SelectNodes("//webParts/webPart").Count -gt 0) {
                    $webPartFormat = "webPart_v3"
                    $typeNodes = $xmlDoc2.SelectNodes("//type/@name")
                    if ($typeNodes.Count -gt 0) {
                        $actualTypeName = $typeNodes[0].Value
                        Write-Host "  Found v3 webPart: $actualTypeName" -ForegroundColor Green
                    }
                }
                
                $xmlInfoElement.AppendChild($xmlDoc.CreateElement("WebPartFormat")).InnerText = $webPartFormat
                if ($actualTypeName) {
                    $xmlInfoElement.AppendChild($xmlDoc.CreateElement("ActualTypeName")).InnerText = $actualTypeName
                }
                
                # Add additional zone information from XML
                if ($partOrder) {
                    $zoneInfoElement.AppendChild($xmlDoc.CreateElement("PartOrder")).InnerText = $partOrder
                }
                if ($zoneIdFromXml) {
                    $zoneInfoElement.AppendChild($xmlDoc.CreateElement("ZoneIdFromXML")).InnerText = $zoneIdFromXml
                }
                
                # Extract Content Editor Web Part content
                if ($actualTypeName -and $actualTypeName -like "*ContentEditorWebPart*") {
                    Write-Host "  Extracting CEWP content..." -ForegroundColor Green
                    
                    $cewpInfoElement = $xmlDoc.CreateElement("CEWPInformation")
                    $xmlInfoElement.AppendChild($cewpInfoElement) | Out-Null
                    
                    # Look for Content in ContentEditor namespace
                    $namespaceManager = New-Object System.Xml.XmlNamespaceManager($xmlDoc2.NameTable)
                    $namespaceManager.AddNamespace("ce", "http://schemas.microsoft.com/WebPart/v2/ContentEditor")
                    $ceContentNodes = $xmlDoc2.SelectNodes("//ce:Content", $namespaceManager)
                    
                    if ($ceContentNodes.Count -gt 0) {
                        $cewpContent = $ceContentNodes[0].InnerText
                        if ($cewpContent -and -not [string]::IsNullOrWhiteSpace($cewpContent)) {
                            $cewpInfoElement.AppendChild($xmlDoc.CreateElement("Content")).InnerText = $cewpContent
                            Write-Host "  Successfully extracted CEWP content ($($cewpContent.Length) characters)" -ForegroundColor Green
                        }
                    }
                    
                    # Look for ContentLink
                    $ceLinkNodes = $xmlDoc2.SelectNodes("//ce:ContentLink", $namespaceManager)
                    if ($ceLinkNodes.Count -gt 0 -and $ceLinkNodes[0].InnerText) {
                        $cewpInfoElement.AppendChild($xmlDoc.CreateElement("ContentLink")).InnerText = $ceLinkNodes[0].InnerText
                    }
                }
                
                # Extract ListView properties from XML
                elseif ($actualTypeName -and ($actualTypeName -like "*ListViewWebPart*" -or $actualTypeName -like "*XsltListViewWebPart*")) {
                    Write-Host "  Extracting ListView properties from XML..." -ForegroundColor Green
                    
                    $listViewInfoElement = $xmlDoc.CreateElement("ListViewInformation")
                    $xmlInfoElement.AppendChild($listViewInfoElement) | Out-Null
                    
                    # Get ListId from XML
                    $listIdNodes = $xmlDoc2.SelectNodes("//property[@name='ListId']")
                    if ($listIdNodes.Count -gt 0 -and $listIdNodes[0].InnerText) {
                        $listViewInfoElement.AppendChild($xmlDoc.CreateElement("ListId")).InnerText = $listIdNodes[0].InnerText
                    }
                    
                    # Get ListName from XML
                    $listNameNodes = $xmlDoc2.SelectNodes("//property[@name='ListName']")
                    if ($listNameNodes.Count -gt 0 -and $listNameNodes[0].InnerText) {
                        $listViewInfoElement.AppendChild($xmlDoc.CreateElement("ListName")).InnerText = $listNameNodes[0].InnerText
                    }
                    
                    # Get TitleUrl from XML
                    $titleUrlNodes = $xmlDoc2.SelectNodes("//property[@name='TitleUrl']")
                    if ($titleUrlNodes.Count -gt 0 -and $titleUrlNodes[0].InnerText) {
                        $listViewInfoElement.AppendChild($xmlDoc.CreateElement("TitleUrl")).InnerText = $titleUrlNodes[0].InnerText
                    }
                    
                    # Get View Definition
                    $xmlDefNodes = $xmlDoc2.SelectNodes("//property[@name='XmlDefinition']")
                    if ($xmlDefNodes.Count -gt 0 -and $xmlDefNodes[0].InnerText) {
                        $listViewInfoElement.AppendChild($xmlDoc.CreateElement("XmlDefinition")).InnerText = $xmlDefNodes[0].InnerText
                    }
                }
            }
            else {
                Write-Warning "  Could not retrieve XML for web part: $($wp.WebPart.Title)"
                $webPartElement.AppendChild($xmlDoc.CreateElement("XMLStatus")).InnerText = "Could not retrieve XML"
            }
        }
        catch {
            Write-Warning "  Error getting XML for web part: $($_.Exception.Message)"
            $webPartElement.AppendChild($xmlDoc.CreateElement("XMLError")).InnerText = $_.Exception.Message
        }
    }
    
    # Generate output filename with timestamp
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $outputFileName = "SharePoint_PageInfo_$timestamp.xml"
    $outputPath = Join-Path $OutputFolder $outputFileName
    
    # Save XML to file
    $xmlDoc.Save($outputPath)
    
    Write-Host "`n=== EXTRACTION COMPLETE ===" -ForegroundColor Green
    Write-Host "Page: $($page['Title'])" -ForegroundColor White
    Write-Host "Web Parts Found: $($webParts.Count)" -ForegroundColor White
    Write-Host "Output saved to: $outputPath" -ForegroundColor Yellow
    
    # Show web part summary with zone information
    Write-Host "`n=== WEB PARTS SUMMARY ===" -ForegroundColor Cyan
    $counter = 0
    foreach ($wp in $webParts) {
        $counter++
        $title = if ($wp.WebPart.Title) { $wp.WebPart.Title } else { "No Title" }
        $zone = if ($wp.ZoneId) { $wp.ZoneId } else { "No Zone" }
        $zoneIndex = if ($wp.ZoneIndex -ne $null) { $wp.ZoneIndex } else { "N/A" }
        Write-Host "$counter. $title" -ForegroundColor White
        Write-Host "    Zone: $zone | Zone Index: $zoneIndex | ID: $($wp.Id)" -ForegroundColor Gray
    }
}
catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
    Write-Host "Stack Trace: $($_.Exception.StackTrace)" -ForegroundColor Red
}

Write-Host "`nScript execution completed." -ForegroundColor Green