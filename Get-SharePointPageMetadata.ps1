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
        # For SharePoint 2016 classic pages, get web parts using Get-PnPWebPart
        $classicWebParts = Get-PnPWebPart -ServerRelativePageUrl $pageFile.ServerRelativeUrl -ErrorAction Stop
        
        if ($classicWebParts -and $classicWebParts.Count -gt 0) {
            Write-Host "Found $($classicWebParts.Count) web parts" -ForegroundColor Green
            
            # Group web parts by zone
            $webPartsByZone = @{}
            
            foreach ($webPart in $classicWebParts) {
                try {
                    $zoneId = if ($webPart.ZoneId) { $webPart.ZoneId } else { "UnknownZone" }
                    
                    if (-not $webPartsByZone.ContainsKey($zoneId)) {
                        $webPartsByZone[$zoneId] = @()
                    }
                    $webPartsByZone[$zoneId] += $webPart
                }
                catch {
                    Write-Warning "Could not process web part for grouping: $($_.Exception.Message)"
                }
            }
            
            # Create XML structure for web parts by zone
            foreach ($zoneId in $webPartsByZone.Keys) {
                $zoneElement = $xmlDoc.CreateElement("WebPartZone")
                $webPartsInfo.AppendChild($zoneElement) | Out-Null
                $zoneElement.SetAttribute("Name", $zoneId)
                $zoneElement.SetAttribute("WebPartCount", $webPartsByZone[$zoneId].Count.ToString())
                
                foreach ($webPart in $webPartsByZone[$zoneId]) {
                    try {
                        $webPartElement = $xmlDoc.CreateElement("WebPart")
                        $zoneElement.AppendChild($webPartElement) | Out-Null
                        
                        # Basic web part information
                        if ($webPart.Id) {
                            $webPartElement.AppendChild($xmlDoc.CreateElement("Id")).InnerText = $webPart.Id.ToString()
                        }
                        
                        if ($webPart.ZoneIndex -ne $null) {
                            $webPartElement.AppendChild($xmlDoc.CreateElement("ZoneIndex")).InnerText = $webPart.ZoneIndex.ToString()
                        }
                        
                        # Web part object properties
                        if ($webPart.WebPart) {
                            $webPartDetailsElement = $xmlDoc.CreateElement("WebPartDetails")
                            $webPartElement.AppendChild($webPartDetailsElement) | Out-Null
                            
                            if ($webPart.WebPart.Title) {
                                $webPartDetailsElement.AppendChild($xmlDoc.CreateElement("Title")).InnerText = $webPart.WebPart.Title
                            }
                            
                            if ($webPart.WebPart.GetType()) {
                                $webPartDetailsElement.AppendChild($xmlDoc.CreateElement("WebPartType")).InnerText = $webPart.WebPart.GetType().Name
                                $webPartDetailsElement.AppendChild($xmlDoc.CreateElement("FullTypeName")).InnerText = $webPart.WebPart.GetType().FullName
                            }
                            
                            # Try to get common web part properties
                            try {
                                if ($webPart.WebPart.Description) {
                                    $webPartDetailsElement.AppendChild($xmlDoc.CreateElement("Description")).InnerText = $webPart.WebPart.Description
                                }
                            } catch { }
                            
                            try {
                                if ($webPart.WebPart.Hidden -ne $null) {
                                    $webPartDetailsElement.AppendChild($xmlDoc.CreateElement("Hidden")).InnerText = $webPart.WebPart.Hidden.ToString()
                                }
                            } catch { }
                            
                            try {
                                if ($webPart.WebPart.ChromeType) {
                                    $webPartDetailsElement.AppendChild($xmlDoc.CreateElement("ChromeType")).InnerText = $webPart.WebPart.ChromeType.ToString()
                                }
                            } catch { }
                            
                            try {
                                if ($webPart.WebPart.Width) {
                                    $webPartDetailsElement.AppendChild($xmlDoc.CreateElement("Width")).InnerText = $webPart.WebPart.Width.ToString()
                                }
                            } catch { }
                            
                            try {
                                if ($webPart.WebPart.Height) {
                                    $webPartDetailsElement.AppendChild($xmlDoc.CreateElement("Height")).InnerText = $webPart.WebPart.Height.ToString()
                                }
                            } catch { }
                            
                            # Try to get specific properties for common web part types
                            $webPartTypeName = $webPart.WebPart.GetType().Name
                            $specificPropsElement = $xmlDoc.CreateElement("SpecificProperties")
                            $webPartDetailsElement.AppendChild($specificPropsElement) | Out-Null
                            
                            switch ($webPartTypeName) {
                                "ContentEditorWebPart" {
                                    try {
                                        if ($webPart.WebPart.Content) {
                                            $specificPropsElement.AppendChild($xmlDoc.CreateElement("Content")).InnerText = $webPart.WebPart.Content.InnerText
                                        }
                                        if ($webPart.WebPart.ContentLink) {
                                            $specificPropsElement.AppendChild($xmlDoc.CreateElement("ContentLink")).InnerText = $webPart.WebPart.ContentLink
                                        }
                                    } catch { }
                                }
                                "XsltListViewWebPart" {
                                    try {
                                        if ($webPart.WebPart.ListName) {
                                            $specificPropsElement.AppendChild($xmlDoc.CreateElement("ListName")).InnerText = $webPart.WebPart.ListName
                                        }
                                        if ($webPart.WebPart.ViewGuid) {
                                            $specificPropsElement.AppendChild($xmlDoc.CreateElement("ViewGuid")).InnerText = $webPart.WebPart.ViewGuid.ToString()
                                        }
                                    } catch { }
                                }
                                "ScriptEditorWebPart" {
                                    try {
                                        if ($webPart.WebPart.Content) {
                                            $specificPropsElement.AppendChild($xmlDoc.CreateElement("Script")).InnerText = $webPart.WebPart.Content
                                        }
                                    } catch { }
                                }
                                "ImageWebPart" {
                                    try {
                                        if ($webPart.WebPart.ImageLink) {
                                            $specificPropsElement.AppendChild($xmlDoc.CreateElement("ImageLink")).InnerText = $webPart.WebPart.ImageLink
                                        }
                                        if ($webPart.WebPart.AlternativeText) {
                                            $specificPropsElement.AppendChild($xmlDoc.CreateElement("AlternativeText")).InnerText = $webPart.WebPart.AlternativeText
                                        }
                                    } catch { }
                                }
                            }
                            
                            # Remove SpecificProperties element if it's empty
                            if (-not $specificPropsElement.HasChildNodes) {
                                $webPartDetailsElement.RemoveChild($specificPropsElement) | Out-Null
                            }
                        }
                    }
                    catch {
                        Write-Warning "Could not process web part details for ID $($webPart.Id): $($_.Exception.Message)"
                        # Still add basic info even if details fail
                        if ($webPart.Id) {
                            $errorElement = $xmlDoc.CreateElement("ProcessingError")
                            $webPartElement.AppendChild($errorElement) | Out-Null
                            $errorElement.InnerText = "Error processing web part details: $($_.Exception.Message)"
                        }
                    }
                }
            }
        }
        else {
            Write-Host "No web parts found on this page" -ForegroundColor Yellow
            $noWebPartsElement = $xmlDoc.CreateElement("Message")
            $webPartsInfo.AppendChild($noWebPartsElement) | Out-Null
            $noWebPartsElement.InnerText = "No web parts found on this page"
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