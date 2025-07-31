# SharePoint 2016 Page Metadata Extractor
# This script connects to SharePoint 2016 with MFA and extracts page metadata in XML format

param(
    [Parameter(Mandatory=$false)]
    [string]$SiteUrl = "https://consulting.global.deloitteonline.com/sites/Aflac/POC",
    
    [Parameter(Mandatory=$false)]
    [string]$PageUrl = "SitePages/stars.aspx",
    
    [Parameter(Mandatory=$false)]
    [string]$OutputFolder = "C:\Users\cmarala\Desktop\Ford\Output"
)

# Import required module
try {
    Import-Module SharePointPnPPowerShell2016 -ErrorAction Stop
    Write-Host "SharePointPnPPowerShell2016 module loaded successfully" -ForegroundColor Green
}
catch {
    Write-Error "Failed to load SharePointPnPPowerShell2016 module. Please ensure it's installed."
    Write-Host "Install with: Install-Module SharePointPnPPowerShell2016" -ForegroundColor Yellow
    exit 1
}

# Create output folder if it doesn't exist
if (!(Test-Path $OutputFolder)) { 
    New-Item -Path $OutputFolder -ItemType Directory -Force
    Write-Host "Created output folder: $OutputFolder" -ForegroundColor Green
}

try {
    # Connect to SharePoint with MFA support
    Write-Host "Connecting to SharePoint site: $SiteUrl" -ForegroundColor Yellow
    Write-Host "Please complete authentication in the browser window..." -ForegroundColor Yellow
    Connect-PnPOnline -Url $SiteUrl -UseWebLogin
    Write-Host "Successfully connected to SharePoint" -ForegroundColor Green
    
    # Get the web information first
    Write-Host "Getting web information..." -ForegroundColor Yellow
    $web = Get-PnPWeb
    if ($null -eq $web) {
        throw "Could not retrieve web information"
    }
    Write-Host "Web retrieved successfully: $($web.Title)" -ForegroundColor Green
    
    # Get the page file
    Write-Host "Retrieving page: $PageUrl" -ForegroundColor Yellow
    try {
        $pageFile = Get-PnPFile -Url $PageUrl -ErrorAction Stop
        if ($null -eq $pageFile) {
            throw "Page file not found: $PageUrl"
        }
        Write-Host "Page file retrieved successfully: $($pageFile.Name)" -ForegroundColor Green
    }
    catch {
        throw "Failed to retrieve page file '$PageUrl': $($_.Exception.Message)"
    }
    
    # Get the page as list item
    Write-Host "Getting page as list item..." -ForegroundColor Yellow
    try {
        $page = Get-PnPFile -Url $PageUrl -AsListItem -ErrorAction Stop
        if ($null -eq $page) {
            throw "Could not retrieve page as list item: $PageUrl"
        }
        Write-Host "Page list item retrieved successfully" -ForegroundColor Green
    }
    catch {
        throw "Failed to retrieve page as list item '$PageUrl': $($_.Exception.Message)"
    }
    
    # Get the list information
    Write-Host "Getting Site Pages list..." -ForegroundColor Yellow
    try {
        $list = Get-PnPList -Identity "Site Pages" -ErrorAction Stop
        if ($null -eq $list) {
            # Try alternative names for the Site Pages list
            $list = Get-PnPList -Identity "Pages" -ErrorAction SilentlyContinue
            if ($null -eq $list) {
                Write-Warning "Could not find Site Pages list, continuing without list information"
            }
        }
        if ($list) {
            Write-Host "Site Pages list retrieved successfully: $($list.Title)" -ForegroundColor Green
        }
    }
    catch {
        Write-Warning "Could not retrieve Site Pages list: $($_.Exception.Message)"
        $list = $null
    }
    
    # Create XML document
    $xmlDoc = New-Object System.Xml.XmlDocument
    $xmlDeclaration = $xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", $null)
    $xmlDoc.AppendChild($xmlDeclaration) | Out-Null
    
    # Root element
    $rootElement = $xmlDoc.CreateElement("SharePointPageMetadata")
    $xmlDoc.AppendChild($rootElement) | Out-Null
    
    # Site Information
    $siteInfo = $xmlDoc.CreateElement("SiteInformation")
    $rootElement.AppendChild($siteInfo) | Out-Null
    
    $siteInfo.AppendChild($xmlDoc.CreateElement("SiteUrl")).InnerText = $SiteUrl
    
    if ($web.Title) {
        $siteInfo.AppendChild($xmlDoc.CreateElement("WebTitle")).InnerText = $web.Title
    }
    if ($web.Description) {
        $siteInfo.AppendChild($xmlDoc.CreateElement("WebDescription")).InnerText = $web.Description
    }
    if ($web.Id) {
        $siteInfo.AppendChild($xmlDoc.CreateElement("WebId")).InnerText = $web.Id.ToString()
    }
    if ($web.ServerRelativeUrl) {
        $siteInfo.AppendChild($xmlDoc.CreateElement("WebServerRelativeUrl")).InnerText = $web.ServerRelativeUrl
    }
    if ($web.Created) {
        $siteInfo.AppendChild($xmlDoc.CreateElement("WebCreated")).InnerText = $web.Created.ToString("yyyy-MM-ddTHH:mm:ssZ")
    }
    if ($web.LastItemModifiedDate) {
        $siteInfo.AppendChild($xmlDoc.CreateElement("WebLastModified")).InnerText = $web.LastItemModifiedDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
    }
    
    # Page Information
    $pageInfo = $xmlDoc.CreateElement("PageInformation")
    $rootElement.AppendChild($pageInfo) | Out-Null
    
    $pageInfo.AppendChild($xmlDoc.CreateElement("PageUrl")).InnerText = $PageUrl
    
    if ($pageFile -and $pageFile.Name) {
        $pageInfo.AppendChild($xmlDoc.CreateElement("PageName")).InnerText = $pageFile.Name
    }
    if ($page -and $page["Title"]) {
        $pageInfo.AppendChild($xmlDoc.CreateElement("PageTitle")).InnerText = $page["Title"].ToString()
    }
    if ($page -and $page["ID"]) {
        $pageInfo.AppendChild($xmlDoc.CreateElement("PageId")).InnerText = $page["ID"].ToString()
    }
    if ($page -and $page["UniqueId"]) {
        $pageInfo.AppendChild($xmlDoc.CreateElement("PageUniqueId")).InnerText = $page["UniqueId"].ToString()
    }
    if ($pageFile -and $pageFile.ServerRelativeUrl) {
        $pageInfo.AppendChild($xmlDoc.CreateElement("PageServerRelativeUrl")).InnerText = $pageFile.ServerRelativeUrl
    }
    if ($pageFile -and $pageFile.Length) {
        $pageInfo.AppendChild($xmlDoc.CreateElement("PageFileSize")).InnerText = $pageFile.Length.ToString()
    }
    if ($pageFile -and $pageFile.CheckOutType) {
        $pageInfo.AppendChild($xmlDoc.CreateElement("PageCheckOutType")).InnerText = $pageFile.CheckOutType.ToString()
    }
    if ($pageFile -and $pageFile.Level) {
        $pageInfo.AppendChild($xmlDoc.CreateElement("PageLevel")).InnerText = $pageFile.Level.ToString()
    }
    
    # Dates
    $datesInfo = $xmlDoc.CreateElement("Dates")
    $pageInfo.AppendChild($datesInfo) | Out-Null
    
    if ($page["Created"]) {
        $datesInfo.AppendChild($xmlDoc.CreateElement("Created")).InnerText = ([DateTime]$page["Created"]).ToString("yyyy-MM-ddTHH:mm:ssZ")
    }
    if ($page["Modified"]) {
        $datesInfo.AppendChild($xmlDoc.CreateElement("Modified")).InnerText = ([DateTime]$page["Modified"]).ToString("yyyy-MM-ddTHH:mm:ssZ")
    }
    if ($pageFile.TimeCreated) {
        $datesInfo.AppendChild($xmlDoc.CreateElement("FileCreated")).InnerText = $pageFile.TimeCreated.ToString("yyyy-MM-ddTHH:mm:ssZ")
    }
    if ($pageFile.TimeLastModified) {
        $datesInfo.AppendChild($xmlDoc.CreateElement("FileModified")).InnerText = $pageFile.TimeLastModified.ToString("yyyy-MM-ddTHH:mm:ssZ")
    }
    
    # Authors
    $authorsInfo = $xmlDoc.CreateElement("Authors")
    $pageInfo.AppendChild($authorsInfo) | Out-Null
    
    if ($page -and $page["Author"] -and $page["Author"] -ne $null) {
        $author = $xmlDoc.CreateElement("CreatedBy")
        $authorsInfo.AppendChild($author) | Out-Null
        if ($page["Author"].LookupValue) {
            $author.AppendChild($xmlDoc.CreateElement("LoginName")).InnerText = $page["Author"].LookupValue
        }
        if ($page["Author"].Email) {
            $author.AppendChild($xmlDoc.CreateElement("Email")).InnerText = $page["Author"].Email
        }
    }
    
    if ($page -and $page["Editor"] -and $page["Editor"] -ne $null) {
        $editor = $xmlDoc.CreateElement("ModifiedBy")
        $authorsInfo.AppendChild($editor) | Out-Null
        if ($page["Editor"].LookupValue) {
            $editor.AppendChild($xmlDoc.CreateElement("LoginName")).InnerText = $page["Editor"].LookupValue
        }
        if ($page["Editor"].Email) {
            $editor.AppendChild($xmlDoc.CreateElement("Email")).InnerText = $page["Editor"].Email
        }
    }
    
    # Content Type Information
    $contentTypeInfo = $xmlDoc.CreateElement("ContentType")
    $pageInfo.AppendChild($contentTypeInfo) | Out-Null
    
    if ($page -and $page["ContentType"]) {
        $contentTypeInfo.AppendChild($xmlDoc.CreateElement("Name")).InnerText = $page["ContentType"].ToString()
    }
    if ($page -and $page["ContentTypeId"]) {
        $contentTypeInfo.AppendChild($xmlDoc.CreateElement("Id")).InnerText = $page["ContentTypeId"].ToString()
    }
    
    # Publishing Information (if available)
    $publishingInfo = $xmlDoc.CreateElement("PublishingInformation")
    $pageInfo.AppendChild($publishingInfo) | Out-Null
    
    # Check for common publishing fields
    $publishingFields = @(
        "PublishingPageLayout",
        "PublishingPageContent",
        "PublishingStartDate",
        "PublishingExpirationDate",
        "PublishingContact",
        "PublishingPageImage",
        "PublishingRollupImage",
        "PublishingHidden"
    )
    
    foreach ($field in $publishingFields) {
        if ($page -and $page.FieldValues -and $page.FieldValues.ContainsKey($field) -and $page[$field] -ne $null) {
            $publishingInfo.AppendChild($xmlDoc.CreateElement($field)).InnerText = $page[$field].ToString()
        }
    }
    
    # Custom Fields/Metadata
    $customFields = $xmlDoc.CreateElement("CustomFields")
    $pageInfo.AppendChild($customFields) | Out-Null
    
    # Get all field values (excluding system fields)
    $systemFields = @("ID", "Title", "Created", "Modified", "Author", "Editor", "ContentType", "ContentTypeId", 
                     "UniqueId", "GUID", "_UIVersionString", "FileRef", "FileDirRef", "FileLeafRef", 
                     "_Level", "_IsCurrentVersion", "ItemChildCount", "FolderChildCount", "_HasCopyDestinations",
                     "_CopySource", "owshiddenversion", "WorkflowVersion", "_UIVersion", "Attachments",
                     "_ModerationStatus", "_ModerationComments", "File_x0020_Type", "HTML_x0020_File_x0020_Type",
                     "Edit", "LinkTitleNoMenu", "LinkTitle", "DocIcon", "FileSizeDisplay", "ServerUrl",
                     "EncodedAbsUrl", "BaseName", "MetaInfo", "_EditMenuTableStart", "_EditMenuTableEnd",
                     "LinkFilenameNoMenu", "LinkFilename", "SelectTitle", "SelectFilename", "IndentLevel",
                     "PermMask", "CheckedOutUserId", "IsCheckedoutToLocal", "CheckedOutTitle", "CheckoutUser",
                     "ScopeId", "VirusStatus", "_CheckinComment", "LinkCheckedOutTitle", "Modified_x0020_By",
                     "Created_x0020_By", "File_x0020_Size", "InstanceID", "Order", "WorkflowInstanceID")
    
    if ($page -and $page.FieldValues) {
        foreach ($field in $page.FieldValues.Keys) {
            if ($systemFields -notcontains $field -and $page[$field] -ne $null -and $page[$field] -ne "") {
                try {
                    $customField = $xmlDoc.CreateElement("Field")
                    $customFields.AppendChild($customField) | Out-Null
                    $customField.SetAttribute("Name", $field)
                    $customField.InnerText = $page[$field].ToString()
                }
                catch {
                    Write-Warning "Could not process custom field '$field': $($_.Exception.Message)"
                }
            }
        }
    }
    
    # Web Parts Information
    $webPartsInfo = $xmlDoc.CreateElement("WebPartsInformation")
    $pageInfo.AppendChild($webPartsInfo) | Out-Null
    
    Write-Host "Retrieving web parts information..." -ForegroundColor Yellow
    try {
        # Get web parts from the page
        $webParts = Get-PnPClientSideComponent -Page $pageFile.Name -ErrorAction SilentlyContinue
        
        if ($webParts -and $webParts.Count -gt 0) {
            Write-Host "Found $($webParts.Count) web parts" -ForegroundColor Green
            
            # Group web parts by section/zone
            $webPartsByZone = @{}
            
            foreach ($webPart in $webParts) {
                try {
                    $zoneIndex = if ($webPart.Section -ne $null) { $webPart.Section } else { "0" }
                    $columnIndex = if ($webPart.Column -ne $null) { $webPart.Column } else { "0" }
                    $zoneKey = "Section_$zoneIndex" + "_Column_$columnIndex"
                    
                    if (-not $webPartsByZone.ContainsKey($zoneKey)) {
                        $webPartsByZone[$zoneKey] = @()
                    }
                    $webPartsByZone[$zoneKey] += $webPart
                }
                catch {
                    Write-Warning "Could not process web part: $($_.Exception.Message)"
                }
            }
            
            # Create XML structure for web parts by zone
            foreach ($zone in $webPartsByZone.Keys) {
                $zoneElement = $xmlDoc.CreateElement("WebPartZone")
                $webPartsInfo.AppendChild($zoneElement) | Out-Null
                $zoneElement.SetAttribute("Name", $zone)
                
                foreach ($webPart in $webPartsByZone[$zone]) {
                    try {
                        $webPartElement = $xmlDoc.CreateElement("WebPart")
                        $zoneElement.AppendChild($webPartElement) | Out-Null
                        
                        if ($webPart.Title) {
                            $webPartElement.AppendChild($xmlDoc.CreateElement("Title")).InnerText = $webPart.Title
                        }
                        if ($webPart.InstanceId) {
                            $webPartElement.AppendChild($xmlDoc.CreateElement("InstanceId")).InnerText = $webPart.InstanceId.ToString()
                        }
                        if ($webPart.WebPartType) {
                            $webPartElement.AppendChild($xmlDoc.CreateElement("WebPartType")).InnerText = $webPart.WebPartType
                        }
                        if ($webPart.Order) {
                            $webPartElement.AppendChild($xmlDoc.CreateElement("Order")).InnerText = $webPart.Order.ToString()
                        }
                        if ($webPart.Section) {
                            $webPartElement.AppendChild($xmlDoc.CreateElement("Section")).InnerText = $webPart.Section.ToString()
                        }
                        if ($webPart.Column) {
                            $webPartElement.AppendChild($xmlDoc.CreateElement("Column")).InnerText = $webPart.Column.ToString()
                        }
                        
                        # Get web part properties if available
                        if ($webPart.PropertiesJson) {
                            $propertiesElement = $xmlDoc.CreateElement("Properties")
                            $webPartElement.AppendChild($propertiesElement) | Out-Null
                            $propertiesElement.InnerText = $webPart.PropertiesJson
                        }
                    }
                    catch {
                        Write-Warning "Could not process web part details: $($_.Exception.Message)"
                    }
                }
            }
        }
        else {
            # Try alternative method for classic pages
            Write-Host "No modern web parts found, trying classic web parts..." -ForegroundColor Yellow
            try {
                # For classic pages, try to get web part manager
                $classicWebParts = Get-PnPWebPart -ServerRelativePageUrl $pageFile.ServerRelativeUrl -ErrorAction SilentlyContinue
                
                if ($classicWebParts -and $classicWebParts.Count -gt 0) {
                    Write-Host "Found $($classicWebParts.Count) classic web parts" -ForegroundColor Green
                    
                    foreach ($webPart in $classicWebParts) {
                        try {
                            $zoneElement = $xmlDoc.CreateElement("WebPartZone")
                            $webPartsInfo.AppendChild($zoneElement) | Out-Null
                            
                            if ($webPart.ZoneId) {
                                $zoneElement.SetAttribute("Name", $webPart.ZoneId)
                            } else {
                                $zoneElement.SetAttribute("Name", "UnknownZone")
                            }
                            
                            $webPartElement = $xmlDoc.CreateElement("WebPart")
                            $zoneElement.AppendChild($webPartElement) | Out-Null
                            
                            if ($webPart.WebPart -and $webPart.WebPart.Title) {
                                $webPartElement.AppendChild($xmlDoc.CreateElement("Title")).InnerText = $webPart.WebPart.Title
                            }
                            if ($webPart.Id) {
                                $webPartElement.AppendChild($xmlDoc.CreateElement("Id")).InnerText = $webPart.Id.ToString()
                            }
                            if ($webPart.WebPart -and $webPart.WebPart.GetType()) {
                                $webPartElement.AppendChild($xmlDoc.CreateElement("WebPartType")).InnerText = $webPart.WebPart.GetType().Name
                            }
                            if ($webPart.ZoneIndex) {
                                $webPartElement.AppendChild($xmlDoc.CreateElement("ZoneIndex")).InnerText = $webPart.ZoneIndex.ToString()
                            }
                        }
                        catch {
                            Write-Warning "Could not process classic web part: $($_.Exception.Message)"
                        }
                    }
                }
                else {
                    Write-Host "No classic web parts found either" -ForegroundColor Yellow
                    $noWebPartsElement = $xmlDoc.CreateElement("Message")
                    $webPartsInfo.AppendChild($noWebPartsElement) | Out-Null
                    $noWebPartsElement.InnerText = "No web parts found on this page"
                }
            }
            catch {
                Write-Warning "Could not retrieve classic web parts: $($_.Exception.Message)"
                $errorElement = $xmlDoc.CreateElement("Error")
                $webPartsInfo.AppendChild($errorElement) | Out-Null
                $errorElement.InnerText = "Error retrieving web parts: $($_.Exception.Message)"
            }
        }
    }
    catch {
        Write-Warning "Could not retrieve web parts: $($_.Exception.Message)"
        $errorElement = $xmlDoc.CreateElement("Error")
        $webPartsInfo.AppendChild($errorElement) | Out-Null
        $errorElement.InnerText = "Error retrieving web parts: $($_.Exception.Message)"
    }
    
    # Generate output filename with timestamp
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $outputFileName = "SharePoint_Page_Metadata_$timestamp.xml"
    $outputPath = Join-Path $OutputFolder $outputFileName
    
    # Save XML to file
    $xmlDoc.Save($outputPath)
    
    Write-Host "Page metadata successfully extracted!" -ForegroundColor Green
    Write-Host "Output file: $outputPath" -ForegroundColor Green
    
    # Display summary
    Write-Host "`n=== METADATA SUMMARY ===" -ForegroundColor Cyan
    
    if ($page -and $page['Title']) {
        Write-Host "Page Title: $($page['Title'])" -ForegroundColor White
    }
    if ($page -and $page['ID']) {
        Write-Host "Page ID: $($page['ID'])" -ForegroundColor White
    }
    if ($page -and $page['Created']) {
        Write-Host "Created: $($page['Created'])" -ForegroundColor White
    }
    if ($page -and $page['Modified']) {
        Write-Host "Modified: $($page['Modified'])" -ForegroundColor White
    }
    if ($pageFile -and $pageFile.Length) {
        Write-Host "File Size: $($pageFile.Length) bytes" -ForegroundColor White
    }
    if ($customFields) {
        Write-Host "Custom Fields Found: $($customFields.ChildNodes.Count)" -ForegroundColor White
    }
    
    if ($webPartsInfo -and $webPartsInfo.ChildNodes.Count -gt 0) {
        $webPartCount = 0
        foreach ($zone in $webPartsInfo.ChildNodes) {
            if ($zone.Name -eq "WebPartZone") {
                $webPartCount += $zone.ChildNodes.Count
            }
        }
        Write-Host "Web Parts Found: $webPartCount" -ForegroundColor White
        Write-Host "Web Part Zones: $($webPartsInfo.SelectNodes('WebPartZone').Count)" -ForegroundColor White
    }
    
    Write-Host "`nXML file saved to: $outputPath" -ForegroundColor Yellow
    
}
catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
    Write-Host "Stack Trace: $($_.Exception.StackTrace)" -ForegroundColor Red
}
finally {
    # Disconnect from SharePoint
    try {
        Disconnect-PnPOnline
        Write-Host "Disconnected from SharePoint" -ForegroundColor Green
    }
    catch {
        Write-Warning "Could not disconnect cleanly: $($_.Exception.Message)"
    }
}

Write-Host "`nScript execution completed." -ForegroundColor Green