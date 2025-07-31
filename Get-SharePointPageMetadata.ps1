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
    
    # Get the page file
    Write-Host "Retrieving page: $PageUrl" -ForegroundColor Yellow
    $page = Get-PnPFile -Url $PageUrl -AsListItem
    
    if ($null -eq $page) {
        throw "Page not found: $PageUrl"
    }
    
    # Get additional page properties
    $pageFile = Get-PnPFile -Url $PageUrl
    $web = Get-PnPWeb
    $list = Get-PnPList -Identity "Site Pages"
    
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
    $siteInfo.AppendChild($xmlDoc.CreateElement("WebTitle")).InnerText = $web.Title
    $siteInfo.AppendChild($xmlDoc.CreateElement("WebDescription")).InnerText = $web.Description
    $siteInfo.AppendChild($xmlDoc.CreateElement("WebId")).InnerText = $web.Id
    $siteInfo.AppendChild($xmlDoc.CreateElement("WebServerRelativeUrl")).InnerText = $web.ServerRelativeUrl
    $siteInfo.AppendChild($xmlDoc.CreateElement("WebCreated")).InnerText = $web.Created.ToString("yyyy-MM-ddTHH:mm:ssZ")
    $siteInfo.AppendChild($xmlDoc.CreateElement("WebLastModified")).InnerText = $web.LastItemModifiedDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
    
    # Page Information
    $pageInfo = $xmlDoc.CreateElement("PageInformation")
    $rootElement.AppendChild($pageInfo) | Out-Null
    
    $pageInfo.AppendChild($xmlDoc.CreateElement("PageUrl")).InnerText = $PageUrl
    $pageInfo.AppendChild($xmlDoc.CreateElement("PageName")).InnerText = $pageFile.Name
    $pageInfo.AppendChild($xmlDoc.CreateElement("PageTitle")).InnerText = $page["Title"]
    $pageInfo.AppendChild($xmlDoc.CreateElement("PageId")).InnerText = $page["ID"]
    $pageInfo.AppendChild($xmlDoc.CreateElement("PageUniqueId")).InnerText = $page["UniqueId"]
    $pageInfo.AppendChild($xmlDoc.CreateElement("PageServerRelativeUrl")).InnerText = $pageFile.ServerRelativeUrl
    $pageInfo.AppendChild($xmlDoc.CreateElement("PageFileSize")).InnerText = $pageFile.Length
    $pageInfo.AppendChild($xmlDoc.CreateElement("PageCheckOutType")).InnerText = $pageFile.CheckOutType
    $pageInfo.AppendChild($xmlDoc.CreateElement("PageLevel")).InnerText = $pageFile.Level
    
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
    
    if ($page["Author"]) {
        $author = $xmlDoc.CreateElement("CreatedBy")
        $authorsInfo.AppendChild($author) | Out-Null
        $author.AppendChild($xmlDoc.CreateElement("LoginName")).InnerText = $page["Author"].LookupValue
        $author.AppendChild($xmlDoc.CreateElement("Email")).InnerText = $page["Author"].Email
    }
    
    if ($page["Editor"]) {
        $editor = $xmlDoc.CreateElement("ModifiedBy")
        $authorsInfo.AppendChild($editor) | Out-Null
        $editor.AppendChild($xmlDoc.CreateElement("LoginName")).InnerText = $page["Editor"].LookupValue
        $editor.AppendChild($xmlDoc.CreateElement("Email")).InnerText = $page["Editor"].Email
    }
    
    # Content Type Information
    $contentTypeInfo = $xmlDoc.CreateElement("ContentType")
    $pageInfo.AppendChild($contentTypeInfo) | Out-Null
    
    if ($page["ContentType"]) {
        $contentTypeInfo.AppendChild($xmlDoc.CreateElement("Name")).InnerText = $page["ContentType"]
        $contentTypeInfo.AppendChild($xmlDoc.CreateElement("Id")).InnerText = $page["ContentTypeId"]
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
        if ($page.FieldValues.ContainsKey($field) -and $page[$field] -ne $null) {
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
    
    foreach ($field in $page.FieldValues.Keys) {
        if ($systemFields -notcontains $field -and $page[$field] -ne $null -and $page[$field] -ne "") {
            $customField = $xmlDoc.CreateElement("Field")
            $customFields.AppendChild($customField) | Out-Null
            $customField.SetAttribute("Name", $field)
            $customField.InnerText = $page[$field].ToString()
        }
    }
    
    # Version Information
    $versionInfo = $xmlDoc.CreateElement("VersionInformation")
    $pageInfo.AppendChild($versionInfo) | Out-Null
    
    if ($page["_UIVersionString"]) {
        $versionInfo.AppendChild($xmlDoc.CreateElement("UIVersion")).InnerText = $page["_UIVersionString"]
    }
    if ($page["_UIVersion"]) {
        $versionInfo.AppendChild($xmlDoc.CreateElement("VersionNumber")).InnerText = $page["_UIVersion"]
    }
    
    # Get version history
    try {
        $versions = Get-PnPFileVersion -Url $PageUrl
        if ($versions) {
            $versionHistory = $xmlDoc.CreateElement("VersionHistory")
            $versionInfo.AppendChild($versionHistory) | Out-Null
            
            foreach ($version in $versions) {
                $versionElement = $xmlDoc.CreateElement("Version")
                $versionHistory.AppendChild($versionElement) | Out-Null
                $versionElement.SetAttribute("VersionLabel", $version.VersionLabel)
                $versionElement.SetAttribute("Size", $version.Size)
                $versionElement.SetAttribute("Created", $version.Created.ToString("yyyy-MM-ddTHH:mm:ssZ"))
                $versionElement.SetAttribute("CreatedBy", $version.CreatedBy.LookupValue)
            }
        }
    }
    catch {
        Write-Warning "Could not retrieve version history: $($_.Exception.Message)"
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
    Write-Host "Page Title: $($page['Title'])" -ForegroundColor White
    Write-Host "Page ID: $($page['ID'])" -ForegroundColor White
    Write-Host "Created: $($page['Created'])" -ForegroundColor White
    Write-Host "Modified: $($page['Modified'])" -ForegroundColor White
    Write-Host "File Size: $($pageFile.Length) bytes" -ForegroundColor White
    Write-Host "Custom Fields Found: $($customFields.ChildNodes.Count)" -ForegroundColor White
    
    if ($versions) {
        Write-Host "Version History: $($versions.Count) versions" -ForegroundColor White
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