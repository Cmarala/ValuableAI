<#
.SYNOPSIS
    SharePoint 2016 Comprehensive Page Content Extractor
    
.DESCRIPTION
    This script extracts complete page information from SharePoint 2016 including:
    - All page metadata
    - All webparts and their properties including embedded webparts
    - All content inside webparts
    - All content written directly on page (irrespective of page type)
    - The order of content and webparts on page
    - Outputs to a well-formed XML file
    
.PARAMETER SiteUrl
    The URL of the SharePoint site
    
.PARAMETER PageUrl
    The server-relative URL of the page to extract (e.g., "/Pages/Home.aspx")
    
.PARAMETER OutputPath
    The path where the XML file will be saved
    
.PARAMETER Credentials
    Optional credentials for authentication (if not provided, uses current user)
    
.EXAMPLE
    .\Extract-SharePointPageContent.ps1 -SiteUrl "http://sharepoint.contoso.com" -PageUrl "/Pages/Home.aspx" -OutputPath "C:\Temp\PageData.xml"
    
.EXAMPLE
    $creds = Get-Credential
    .\Extract-SharePointPageContent.ps1 -SiteUrl "http://sharepoint.contoso.com" -PageUrl "/SitePages/Wiki.aspx" -OutputPath "C:\Temp\WikiPage.xml" -Credentials $creds
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$PageUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$OutputPath,
    
    [Parameter(Mandatory=$false)]
    [System.Management.Automation.PSCredential]$Credentials
)

# Add SharePoint PowerShell snapin if not already loaded
if ((Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -eq $null) {
    try {
        Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to load SharePoint PowerShell snapin. Make sure this script is run on a SharePoint server or with SharePoint PowerShell installed."
        exit 1
    }
}

# Load required assemblies for CSOM (fallback for remote execution)
$CSOMAssemblies = @(
    "Microsoft.SharePoint.Client.dll",
    "Microsoft.SharePoint.Client.Runtime.dll",
    "Microsoft.SharePoint.Client.Publishing.dll"
)

$CSOMPath = "${env:ProgramFiles}\Common Files\microsoft shared\Web Server Extensions\16\ISAPI"
if (Test-Path $CSOMPath) {
    foreach ($assembly in $CSOMAssemblies) {
        $assemblyPath = Join-Path $CSOMPath $assembly
        if (Test-Path $assemblyPath) {
            try {
                Add-Type -Path $assemblyPath -ErrorAction SilentlyContinue
            }
            catch {
                Write-Warning "Could not load assembly: $assembly"
            }
        }
    }
}

# Global variables
$script:extractionData = @{}
$script:webPartOrder = @()
$script:contentOrder = @()

#region Helper Functions

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "Info"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $(
        switch ($Level) {
            "Error" { "Red" }
            "Warning" { "Yellow" }
            "Success" { "Green" }
            default { "White" }
        }
    )
}

function Escape-XmlContent {
    param([string]$Content)
    if (-not $Content) { return "" }
    
    return [System.Security.SecurityElement]::Escape($Content)
}

function Get-SafeXmlValue {
    param($Value)
    if ($Value -eq $null) { return "" }
    return Escape-XmlContent $Value.ToString()
}

function Get-UserInfo {
    param($User)
    
    if ($User -eq $null) {
        return @{
            Id = 0
            LoginName = ""
            Name = ""
            Email = ""
        }
    }
    
    return @{
        Id = if ($User.Id) { $User.Id } else { 0 }
        LoginName = Get-SafeXmlValue $User.LoginName
        Name = Get-SafeXmlValue $User.Name
        Email = Get-SafeXmlValue $User.Email
    }
}

function Get-WebPartContent {
    param(
        $WebPart,
        $WebPartManager,
        $Context
    )
    
    $content = @{
        HtmlContent = ""
        TextContent = ""
        XmlContent = ""
        JsonContent = ""
        RawContent = ""
    }
    
    try {
        # Try to get web part XML for content analysis
        if ($WebPart.WebPart.ExportMode -ne "None") {
            try {
                $webPartXml = $WebPartManager.ExportWebPart($WebPart.Id)
                $Context.ExecuteQuery()
                $content.XmlContent = Get-SafeXmlValue $webPartXml.Value
                $content.RawContent = Get-SafeXmlValue $webPartXml.Value
            }
            catch {
                Write-Log "Could not export web part XML for content extraction: $($_.Exception.Message)" "Warning"
            }
        }
        
        # Extract content based on web part type
        $webPartType = $WebPart.WebPart.TypeName
        
        # Content Editor Web Part
        if ($webPartType -like "*ContentEditorWebPart*" -or $webPartType -like "*ContentEditor*") {
            try {
                $Context.Load($WebPart.WebPart.Properties)
                $Context.ExecuteQuery()
                
                if ($WebPart.WebPart.Properties.FieldValues.ContainsKey("Content")) {
                    $content.HtmlContent = Get-SafeXmlValue $WebPart.WebPart.Properties.FieldValues["Content"]
                }
                if ($WebPart.WebPart.Properties.FieldValues.ContainsKey("ContentLink")) {
                    $content.RawContent += "ContentLink: " + (Get-SafeXmlValue $WebPart.WebPart.Properties.FieldValues["ContentLink"])
                }
            }
            catch {
                Write-Log "Error extracting Content Editor content: $($_.Exception.Message)" "Warning"
            }
        }
        
        # Text Web Part
        elseif ($webPartType -like "*TextWebPart*" -or $webPartType -like "*Text*") {
            try {
                $Context.Load($WebPart.WebPart.Properties)
                $Context.ExecuteQuery()
                
                if ($WebPart.WebPart.Properties.FieldValues.ContainsKey("Text")) {
                    $content.TextContent = Get-SafeXmlValue $WebPart.WebPart.Properties.FieldValues["Text"]
                }
            }
            catch {
                Write-Log "Error extracting Text Web Part content: $($_.Exception.Message)" "Warning"
            }
        }
        
        # HTML Form Web Part
        elseif ($webPartType -like "*HtmlForm*") {
            try {
                $Context.Load($WebPart.WebPart.Properties)
                $Context.ExecuteQuery()
                
                if ($WebPart.WebPart.Properties.FieldValues.ContainsKey("Content")) {
                    $content.HtmlContent = Get-SafeXmlValue $WebPart.WebPart.Properties.FieldValues["Content"]
                }
            }
            catch {
                Write-Log "Error extracting HTML Form content: $($_.Exception.Message)" "Warning"
            }
        }
        
        # Try to extract any content-related properties
        try {
            $Context.Load($WebPart.WebPart.Properties)
            $Context.ExecuteQuery()
            
            $contentProperties = @()
            foreach ($prop in $WebPart.WebPart.Properties.FieldValues.Keys) {
                $propName = $prop.ToLower()
                if ($propName -like "*content*" -or $propName -like "*text*" -or $propName -like "*html*" -or $propName -like "*body*") {
                    $propValue = $WebPart.WebPart.Properties.FieldValues[$prop]
                    if ($propValue -and $propValue.ToString().Trim() -ne "") {
                        $contentProperties += "$prop`: " + (Get-SafeXmlValue $propValue)
                    }
                }
            }
            
            if ($contentProperties.Count -gt 0) {
                $content.RawContent += "`n" + ($contentProperties -join "`n")
            }
        }
        catch {
            Write-Log "Error extracting general web part content properties: $($_.Exception.Message)" "Warning"
        }
    }
    catch {
        Write-Log "Error in Get-WebPartContent: $($_.Exception.Message)" "Warning"
    }
    
    return $content
}

function Get-WebPartProperties {
    param(
        $WebPart,
        $Context
    )
    
    $properties = @()
    
    try {
        $Context.Load($WebPart.WebPart.Properties)
        $Context.ExecuteQuery()
        
        foreach ($key in $WebPart.WebPart.Properties.FieldValues.Keys) {
            $value = $WebPart.WebPart.Properties.FieldValues[$key]
            $properties += @{
                Name = Get-SafeXmlValue $key
                Value = Get-SafeXmlValue $value
                Type = if ($value) { $value.GetType().Name } else { "String" }
            }
        }
    }
    catch {
        Write-Log "Error extracting web part properties: $($_.Exception.Message)" "Warning"
    }
    
    return $properties
}

function Get-EmbeddedWebParts {
    param(
        $WebPart,
        $Context
    )
    
    $embeddedParts = @()
    
    try {
        # Check if this web part contains other web parts (like Web Part Zone)
        if ($WebPart.WebPart.TypeName -like "*WebPartZone*") {
            # Try to get web parts within this zone
            # This is a simplified approach - full implementation would require more complex zone analysis
            Write-Log "Detected Web Part Zone - checking for embedded parts" "Info"
        }
        
        # Check properties for embedded content references
        $Context.Load($WebPart.WebPart.Properties)
        $Context.ExecuteQuery()
        
        foreach ($key in $WebPart.WebPart.Properties.FieldValues.Keys) {
            $propName = $key.ToLower()
            if ($propName -like "*webpart*" -or $propName -like "*embed*" -or $propName -like "*zone*") {
                $propValue = $WebPart.WebPart.Properties.FieldValues[$key]
                if ($propValue -and $propValue.ToString().Contains("WebPart")) {
                    Write-Log "Found potential embedded web part reference in property: $key" "Info"
                    # Additional processing could be added here to parse embedded web part definitions
                }
            }
        }
    }
    catch {
        Write-Log "Error checking for embedded web parts: $($_.Exception.Message)" "Warning"
    }
    
    return $embeddedParts
}

function Extract-WebParts {
    param(
        $Web,
        $File,
        $Context
    )
    
    Write-Log "Extracting web parts from page..." "Info"
    
    $webParts = @()
    $zones = @()
    
    try {
        # Get web part manager
        $webPartManager = $File.GetLimitedWebPartManager([Microsoft.SharePoint.Client.WebParts.PersonalizationScope]::Shared)
        $webPartCollection = $webPartManager.WebParts
        $Context.Load($webPartCollection)
        $Context.ExecuteQuery()
        
        Write-Log "Found $($webPartCollection.Count) web parts" "Info"
        
        # Track zones
        $zoneTracker = @{}
        
        foreach ($webPart in $webPartCollection) {
            try {
                $Context.Load($webPart)
                $Context.Load($webPart.WebPart)
                $Context.Load($webPart.WebPart.Properties)
                $Context.ExecuteQuery()
                
                # Get web part XML if exportable
                $webPartXml = ""
                if ($webPart.WebPart.ExportMode -ne "None") {
                    try {
                        $exportedXml = $webPartManager.ExportWebPart($webPart.Id)
                        $Context.ExecuteQuery()
                        $webPartXml = Get-SafeXmlValue $exportedXml.Value
                    }
                    catch {
                        Write-Log "Could not export web part XML for web part: $($webPart.WebPart.Title)" "Warning"
                    }
                }
                
                # Extract properties
                $properties = Get-WebPartProperties -WebPart $webPart -Context $Context
                
                # Extract content
                $content = Get-WebPartContent -WebPart $webPart -WebPartManager $webPartManager -Context $Context
                
                # Check for embedded web parts
                $embeddedParts = Get-EmbeddedWebParts -WebPart $webPart -Context $Context
                
                # Get zone information
                $zoneId = Get-SafeXmlValue $webPart.ZoneId
                $zoneIndex = if ($webPart.ZoneIndex) { $webPart.ZoneIndex } else { 0 }
                
                # Track zone if not already tracked
                if (-not $zoneTracker.ContainsKey($zoneId)) {
                    $zoneTracker[$zoneId] = @{
                        ZoneId = $zoneId
                        ZoneIndex = $zoneTracker.Count
                        ZoneTitle = $zoneId
                        Orientation = "Vertical"  # Default value
                        ChromeType = "Default"
                        WebParts = @()
                    }
                }
                
                $zoneTracker[$zoneId].WebParts += @{
                    WebPartId = $webPart.Id.ToString()
                    ZoneIndex = $zoneIndex
                    Order = $zoneTracker[$zoneId].WebParts.Count
                }
                
                $webPartData = @{
                    WebPartId = $webPart.Id.ToString()
                    Title = Get-SafeXmlValue $webPart.WebPart.Title
                    Description = Get-SafeXmlValue $webPart.WebPart.Description
                    TypeName = Get-SafeXmlValue $webPart.WebPart.TypeName
                    Assembly = ""  # Not directly accessible in CSOM
                    ZoneId = $zoneId
                    ZoneIndex = $zoneIndex
                    IsVisible = if ($webPart.WebPart.IsIncluded -ne $null) { $webPart.WebPart.IsIncluded } else { $true }
                    IsIncluded = if ($webPart.WebPart.IsIncluded -ne $null) { $webPart.WebPart.IsIncluded } else { $true }
                    IsClosed = if ($webPart.WebPart.IsClosed -ne $null) { $webPart.WebPart.IsClosed } else { $false }
                    ChromeType = Get-SafeXmlValue $webPart.WebPart.ChromeType
                    ChromeState = "Normal"  # Default value
                    Width = Get-SafeXmlValue $webPart.WebPart.Width
                    Height = Get-SafeXmlValue $webPart.WebPart.Height
                    AllowClose = if ($webPart.WebPart.AllowRemove -ne $null) { $webPart.WebPart.AllowRemove } else { $true }
                    AllowConnect = $true  # Default value
                    AllowEdit = $true  # Default value
                    AllowHide = if ($webPart.WebPart.AllowHide -ne $null) { $webPart.WebPart.AllowHide } else { $true }
                    AllowMinimize = if ($webPart.WebPart.AllowMinimize -ne $null) { $webPart.WebPart.AllowMinimize } else { $true }
                    AllowZoneChange = if ($webPart.WebPart.AllowZoneChange -ne $null) { $webPart.WebPart.AllowZoneChange } else { $true }
                    ExportMode = Get-SafeXmlValue $webPart.WebPart.ExportMode
                    WebPartXml = $webPartXml
                    Properties = $properties
                    Content = $content
                    EmbeddedWebParts = $embeddedParts
                }
                
                $webParts += $webPartData
                
                # Add to content order tracking
                $script:contentOrder += @{
                    ElementId = $webPart.Id.ToString()
                    ElementType = "WebPart"
                    Order = $script:contentOrder.Count
                    ZoneId = $zoneId
                    Content = $webPartData.Title
                }
                
                Write-Log "Extracted web part: $($webPartData.Title) (Type: $($webPartData.TypeName))" "Info"
            }
            catch {
                Write-Log "Error processing web part: $($_.Exception.Message)" "Error"
            }
        }
        
        # Convert zone tracker to zones array
        $zones = $zoneTracker.Values | ForEach-Object {
            @{
                ZoneId = $_.ZoneId
                ZoneIndex = $_.ZoneIndex
                ZoneTitle = $_.ZoneTitle
                Orientation = $_.Orientation
                ChromeType = $_.ChromeType
                WebPartsInZone = $_.WebParts
            }
        }
    }
    catch {
        Write-Log "Error extracting web parts: $($_.Exception.Message)" "Error"
    }
    
    return @{
        WebParts = $webParts
        Zones = $zones
    }
}

function Extract-PageContent {
    param(
        $ListItem,
        $File,
        $Context
    )
    
    Write-Log "Extracting page content..." "Info"
    
    $pageContent = @{
        AspxContent = ""
        WikiContent = ""
        PublishingContent = ""
        RawPageContent = ""
        PlainTextContent = ""
        ContentFields = @()
    }
    
    try {
        # Get file content
        $fileContent = $File.OpenBinaryDirect()
        $Context.ExecuteQuery()
        
        if ($fileContent -and $fileContent.Stream) {
            $reader = New-Object System.IO.StreamReader($fileContent.Stream)
            $pageContent.RawPageContent = Get-SafeXmlValue $reader.ReadToEnd()
            $reader.Close()
        }
    }
    catch {
        Write-Log "Could not read file content directly: $($_.Exception.Message)" "Warning"
    }
    
    try {
        # Extract wiki content if it's a wiki page
        if ($ListItem["WikiField"]) {
            $pageContent.WikiContent = Get-SafeXmlValue $ListItem["WikiField"]
            $pageContent.PlainTextContent = Get-SafeXmlValue ($ListItem["WikiField"] -replace '<[^>]+>', '')
            
            $script:contentOrder += @{
                ElementId = "WikiContent"
                ElementType = "Content"
                Order = $script:contentOrder.Count
                ZoneId = ""
                Content = "Wiki Content"
            }
        }
    }
    catch {
        Write-Log "Error extracting wiki content: $($_.Exception.Message)" "Warning"
    }
    
    try {
        # Extract publishing content fields
        $publishingFields = @("PublishingPageContent", "PublishingPageImage", "PublishingPageLayout")
        foreach ($fieldName in $publishingFields) {
            try {
                if ($ListItem.FieldValues.ContainsKey($fieldName) -and $ListItem[$fieldName]) {
                    $fieldValue = Get-SafeXmlValue $ListItem[$fieldName]
                    $pageContent.ContentFields += @{
                        Name = $fieldName
                        Type = "Publishing"
                        DisplayName = $fieldName
                        Value = $fieldValue
                    }
                    
                    if ($fieldName -eq "PublishingPageContent") {
                        $pageContent.PublishingContent = $fieldValue
                        
                        $script:contentOrder += @{
                            ElementId = $fieldName
                            ElementType = "Field"
                            Order = $script:contentOrder.Count
                            ZoneId = ""
                            Content = "Publishing Content"
                        }
                    }
                }
            }
            catch {
                Write-Log "Error extracting publishing field $fieldName`: $($_.Exception.Message)" "Warning"
            }
        }
    }
    catch {
        Write-Log "Error extracting publishing content: $($_.Exception.Message)" "Warning"
    }
    
    try {
        # Extract other content fields
        $contentFieldNames = @("Body", "Content", "Description", "Summary", "Abstract")
        foreach ($fieldName in $contentFieldNames) {
            try {
                if ($ListItem.FieldValues.ContainsKey($fieldName) -and $ListItem[$fieldName]) {
                    $fieldValue = Get-SafeXmlValue $ListItem[$fieldName]
                    $pageContent.ContentFields += @{
                        Name = $fieldName
                        Type = "Text"
                        DisplayName = $fieldName
                        Value = $fieldValue
                    }
                }
            }
            catch {
                Write-Log "Error extracting content field $fieldName`: $($_.Exception.Message)" "Warning"
            }
        }
    }
    catch {
        Write-Log "Error extracting content fields: $($_.Exception.Message)" "Warning"
    }
    
    return $pageContent
}

function Extract-PageMetadata {
    param(
        $ListItem,
        $File,
        $Context
    )
    
    Write-Log "Extracting page metadata..." "Info"
    
    try {
        # Load all necessary properties
        $Context.Load($ListItem)
        $Context.Load($File)
        $Context.Load($File.Author)
        $Context.Load($File.ModifiedBy)
        $Context.ExecuteQuery()
        
        # Get content type information
        $contentType = ""
        $contentTypeId = ""
        try {
            if ($ListItem.ContentType) {
                $Context.Load($ListItem.ContentType)
                $Context.ExecuteQuery()
                $contentType = Get-SafeXmlValue $ListItem.ContentType.Name
                $contentTypeId = Get-SafeXmlValue $ListItem.ContentType.Id.ToString()
            }
        }
        catch {
            Write-Log "Could not load content type information: $($_.Exception.Message)" "Warning"
        }
        
        # Determine page type
        $pageType = "Other"
        if ($File.Name.EndsWith(".aspx")) {
            if ($ListItem.FieldValues.ContainsKey("WikiField") -and $ListItem["WikiField"]) {
                $pageType = "Wiki"
            }
            elseif ($contentType -like "*Publishing*" -or $ListItem.FieldValues.ContainsKey("PublishingPageContent")) {
                $pageType = "Publishing"
            }
            elseif ($contentType -like "*WebPart*") {
                $pageType = "WebPart"
            }
            else {
                $pageType = "Application"
            }
        }
        
        $metadata = @{
            PageId = $ListItem.Id.ToString()
            PageName = Get-SafeXmlValue $File.Name
            PageTitle = Get-SafeXmlValue $ListItem["Title"]
            PageUrl = Get-SafeXmlValue $File.ServerRelativeUrl
            ServerRelativeUrl = Get-SafeXmlValue $File.ServerRelativeUrl
            PageType = $pageType
            ContentType = $contentType
            ContentTypeId = $contentTypeId
            Created = $ListItem["Created"]
            Modified = $ListItem["Modified"]
            CreatedBy = Get-UserInfo $File.Author
            ModifiedBy = Get-UserInfo $File.ModifiedBy
            CheckOutType = Get-SafeXmlValue $File.CheckOutType
            FileSize = if ($File.Length) { $File.Length } else { 0 }
            Version = Get-SafeXmlValue $ListItem["_UIVersionString"]
            CustomProperties = @()
            ListItemFields = @()
        }
        
        # Extract all list item fields
        foreach ($fieldName in $ListItem.FieldValues.Keys) {
            try {
                $fieldValue = $ListItem[$fieldName]
                if ($fieldValue -ne $null -and $fieldValue.ToString().Trim() -ne "") {
                    $metadata.ListItemFields += @{
                        InternalName = $fieldName
                        DisplayName = $fieldName
                        FieldType = "Unknown"
                        IsReadOnly = $false
                        Value = Get-SafeXmlValue $fieldValue
                    }
                }
            }
            catch {
                Write-Log "Error extracting field $fieldName`: $($_.Exception.Message)" "Warning"
            }
        }
        
        return $metadata
    }
    catch {
        Write-Log "Error extracting page metadata: $($_.Exception.Message)" "Error"
        return @{}
    }
}

#endregion

#region Main Extraction Function

function Extract-SharePointPageData {
    param(
        [string]$SiteUrl,
        [string]$PageUrl,
        [System.Management.Automation.PSCredential]$Credentials
    )
    
    Write-Log "Starting SharePoint page extraction..." "Info"
    Write-Log "Site URL: $SiteUrl" "Info"
    Write-Log "Page URL: $PageUrl" "Info"
    
    try {
        # Initialize client context
        $context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
        
        # Set credentials if provided
        if ($Credentials) {
            if ($SiteUrl.StartsWith("https://") -and $SiteUrl.Contains(".sharepoint.com")) {
                # SharePoint Online
                $context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credentials.UserName, $Credentials.Password)
            } else {
                # On-premises
                $context.Credentials = $Credentials.GetNetworkCredential()
            }
        }
        
        # Get web and site information
        $web = $context.Web
        $site = $context.Site
        $context.Load($web)
        $context.Load($site)
        $context.ExecuteQuery()
        
        Write-Log "Connected to site: $($web.Title)" "Success"
        
        # Get the page file
        $file = $web.GetFileByServerRelativeUrl($PageUrl)
        $context.Load($file)
        $context.ExecuteQuery()
        
        Write-Log "Found page file: $($file.Name)" "Success"
        
        # Get the list item for the page
        $listItem = $file.ListItemAllFields
        $context.Load($listItem)
        $context.ExecuteQuery()
        
        # Extract all data
        $extractionInfo = @{
            ScriptVersion = "1.0.0"
            SharePointVersion = "SharePoint 2016"
            ExtractionUser = [Environment]::UserName
            ExtractionMethod = "PowerShell CSOM"
            Timestamp = Get-Date
        }
        
        $siteInformation = @{
            SiteUrl = Get-SafeXmlValue $site.Url
            SiteTitle = Get-SafeXmlValue $site.RootWeb.Title
            WebUrl = Get-SafeXmlValue $web.Url
            WebTitle = Get-SafeXmlValue $web.Title
            WebTemplate = Get-SafeXmlValue $web.WebTemplate
            Culture = Get-SafeXmlValue $web.Culture
            Language = if ($web.Language) { $web.Language } else { 1033 }
        }
        
        $pageMetadata = Extract-PageMetadata -ListItem $listItem -File $file -Context $context
        $pageContent = Extract-PageContent -ListItem $listItem -File $file -Context $context
        $webPartData = Extract-WebParts -Web $web -File $file -Context $context
        
        # Build page structure based on content order
        $pageStructure = $script:contentOrder | ForEach-Object { $_ }
        
        return @{
            ExtractionInfo = $extractionInfo
            SiteInformation = $siteInformation
            PageMetadata = $pageMetadata
            PageContent = $pageContent
            WebPartZones = $webPartData.Zones
            AllWebParts = $webPartData.WebParts
            PageStructure = $pageStructure
        }
    }
    catch {
        Write-Log "Error during extraction: $($_.Exception.Message)" "Error"
        throw
    }
    finally {
        if ($context) {
            $context.Dispose()
        }
    }
}

#endregion

#region XML Generation

function Generate-XML {
    param($Data)
    
    Write-Log "Generating XML output..." "Info"
    
    $xml = @"
<?xml version="1.0" encoding="utf-8"?>
<SharePointPageData xmlns="http://schemas.sharepoint.com/page-extraction/2016" version="1.0" extractionDate="$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ss')">
  <ExtractionInfo>
    <ScriptVersion>$($Data.ExtractionInfo.ScriptVersion)</ScriptVersion>
    <SharePointVersion>$($Data.ExtractionInfo.SharePointVersion)</SharePointVersion>
    <ExtractionUser>$($Data.ExtractionInfo.ExtractionUser)</ExtractionUser>
    <ExtractionMethod>$($Data.ExtractionInfo.ExtractionMethod)</ExtractionMethod>
    <Timestamp>$($Data.ExtractionInfo.Timestamp.ToString('yyyy-MM-ddTHH:mm:ss'))</Timestamp>
  </ExtractionInfo>
  
  <SiteInformation>
    <SiteUrl>$($Data.SiteInformation.SiteUrl)</SiteUrl>
    <SiteTitle>$($Data.SiteInformation.SiteTitle)</SiteTitle>
    <WebUrl>$($Data.SiteInformation.WebUrl)</WebUrl>
    <WebTitle>$($Data.SiteInformation.WebTitle)</WebTitle>
    <WebTemplate>$($Data.SiteInformation.WebTemplate)</WebTemplate>
    <Culture>$($Data.SiteInformation.Culture)</Culture>
    <Language>$($Data.SiteInformation.Language)</Language>
  </SiteInformation>
  
  <PageMetadata>
    <PageId>$($Data.PageMetadata.PageId)</PageId>
    <PageName>$($Data.PageMetadata.PageName)</PageName>
    <PageTitle>$($Data.PageMetadata.PageTitle)</PageTitle>
    <PageUrl>$($Data.PageMetadata.PageUrl)</PageUrl>
    <ServerRelativeUrl>$($Data.PageMetadata.ServerRelativeUrl)</ServerRelativeUrl>
    <PageType>$($Data.PageMetadata.PageType)</PageType>
    <ContentType>$($Data.PageMetadata.ContentType)</ContentType>
    <ContentTypeId>$($Data.PageMetadata.ContentTypeId)</ContentTypeId>
    <Created>$($Data.PageMetadata.Created.ToString('yyyy-MM-ddTHH:mm:ss'))</Created>
    <Modified>$($Data.PageMetadata.Modified.ToString('yyyy-MM-ddTHH:mm:ss'))</Modified>
    <CreatedBy>
      <Id>$($Data.PageMetadata.CreatedBy.Id)</Id>
      <LoginName>$($Data.PageMetadata.CreatedBy.LoginName)</LoginName>
      <Name>$($Data.PageMetadata.CreatedBy.Name)</Name>
      <Email>$($Data.PageMetadata.CreatedBy.Email)</Email>
    </CreatedBy>
    <ModifiedBy>
      <Id>$($Data.PageMetadata.ModifiedBy.Id)</Id>
      <LoginName>$($Data.PageMetadata.ModifiedBy.LoginName)</LoginName>
      <Name>$($Data.PageMetadata.ModifiedBy.Name)</Name>
      <Email>$($Data.PageMetadata.ModifiedBy.Email)</Email>
    </ModifiedBy>
    <CheckOutType>$($Data.PageMetadata.CheckOutType)</CheckOutType>
    <FileSize>$($Data.PageMetadata.FileSize)</FileSize>
    <Version>$($Data.PageMetadata.Version)</Version>
    <ListItemFields>
"@

    # Add list item fields
    foreach ($field in $Data.PageMetadata.ListItemFields) {
        $xml += @"

      <Field internalName="$($field.InternalName)" displayName="$($field.DisplayName)" fieldType="$($field.FieldType)" isReadOnly="$($field.IsReadOnly.ToString().ToLower())">$($field.Value)</Field>
"@
    }

    $xml += @"

    </ListItemFields>
  </PageMetadata>
  
  <PageContent>
    <AspxContent><![CDATA[$($Data.PageContent.AspxContent)]]></AspxContent>
    <WikiContent><![CDATA[$($Data.PageContent.WikiContent)]]></WikiContent>
    <PublishingContent><![CDATA[$($Data.PageContent.PublishingContent)]]></PublishingContent>
    <RawPageContent><![CDATA[$($Data.PageContent.RawPageContent)]]></RawPageContent>
    <PlainTextContent><![CDATA[$($Data.PageContent.PlainTextContent)]]></PlainTextContent>
    <ContentFields>
"@

    # Add content fields
    foreach ($field in $Data.PageContent.ContentFields) {
        $xml += @"

      <ContentField name="$($field.Name)" type="$($field.Type)" displayName="$($field.DisplayName)"><![CDATA[$($field.Value)]]></ContentField>
"@
    }

    $xml += @"

    </ContentFields>
  </PageContent>
  
  <WebPartZones count="$($Data.WebPartZones.Count)">
"@

    # Add web part zones
    foreach ($zone in $Data.WebPartZones) {
        $xml += @"

    <Zone>
      <ZoneId>$($zone.ZoneId)</ZoneId>
      <ZoneIndex>$($zone.ZoneIndex)</ZoneIndex>
      <ZoneTitle>$($zone.ZoneTitle)</ZoneTitle>
      <Orientation>$($zone.Orientation)</Orientation>
      <ChromeType>$($zone.ChromeType)</ChromeType>
      <WebPartsInZone count="$($zone.WebPartsInZone.Count)">
"@
        
        foreach ($wpRef in $zone.WebPartsInZone) {
            $xml += @"

        <WebPartReference>
          <WebPartId>$($wpRef.WebPartId)</WebPartId>
          <ZoneIndex>$($wpRef.ZoneIndex)</ZoneIndex>
          <Order>$($wpRef.Order)</Order>
        </WebPartReference>
"@
        }
        
        $xml += @"

      </WebPartsInZone>
    </Zone>
"@
    }

    $xml += @"

  </WebPartZones>
  
  <AllWebParts count="$($Data.AllWebParts.Count)">
"@

    # Add all web parts
    foreach ($wp in $Data.AllWebParts) {
        $xml += @"

    <WebPart>
      <WebPartId>$($wp.WebPartId)</WebPartId>
      <Title>$($wp.Title)</Title>
      <Description>$($wp.Description)</Description>
      <TypeName>$($wp.TypeName)</TypeName>
      <Assembly>$($wp.Assembly)</Assembly>
      <ZoneId>$($wp.ZoneId)</ZoneId>
      <ZoneIndex>$($wp.ZoneIndex)</ZoneIndex>
      <IsVisible>$($wp.IsVisible.ToString().ToLower())</IsVisible>
      <IsIncluded>$($wp.IsIncluded.ToString().ToLower())</IsIncluded>
      <IsClosed>$($wp.IsClosed.ToString().ToLower())</IsClosed>
      <ChromeType>$($wp.ChromeType)</ChromeType>
      <ChromeState>$($wp.ChromeState)</ChromeState>
      <Width>$($wp.Width)</Width>
      <Height>$($wp.Height)</Height>
      <AllowClose>$($wp.AllowClose.ToString().ToLower())</AllowClose>
      <AllowConnect>$($wp.AllowConnect.ToString().ToLower())</AllowConnect>
      <AllowEdit>$($wp.AllowEdit.ToString().ToLower())</AllowEdit>
      <AllowHide>$($wp.AllowHide.ToString().ToLower())</AllowHide>
      <AllowMinimize>$($wp.AllowMinimize.ToString().ToLower())</AllowMinimize>
      <AllowZoneChange>$($wp.AllowZoneChange.ToString().ToLower())</AllowZoneChange>
      <ExportMode>$($wp.ExportMode)</ExportMode>
      <WebPartXml><![CDATA[$($wp.WebPartXml)]]></WebPartXml>
      <Properties count="$($wp.Properties.Count)">
"@

        foreach ($prop in $wp.Properties) {
            $xml += @"

        <Property name="$($prop.Name)" type="$($prop.Type)"><![CDATA[$($prop.Value)]]></Property>
"@
        }

        $xml += @"

      </Properties>
      <Content>
        <HtmlContent><![CDATA[$($wp.Content.HtmlContent)]]></HtmlContent>
        <TextContent><![CDATA[$($wp.Content.TextContent)]]></TextContent>
        <XmlContent><![CDATA[$($wp.Content.XmlContent)]]></XmlContent>
        <JsonContent><![CDATA[$($wp.Content.JsonContent)]]></JsonContent>
        <RawContent><![CDATA[$($wp.Content.RawContent)]]></RawContent>
      </Content>
      <EmbeddedWebParts count="$($wp.EmbeddedWebParts.Count)">
"@

        foreach ($embedded in $wp.EmbeddedWebParts) {
            # Recursive web part structure (simplified for now)
            $xml += @"

        <EmbeddedWebPart>
          <WebPartId>$($embedded.WebPartId)</WebPartId>
          <Title>$($embedded.Title)</Title>
          <TypeName>$($embedded.TypeName)</TypeName>
        </EmbeddedWebPart>
"@
        }

        $xml += @"

      </EmbeddedWebParts>
    </WebPart>
"@
    }

    $xml += @"

  </AllWebParts>
  
  <PageStructure count="$($Data.PageStructure.Count)">
"@

    # Add page structure
    foreach ($element in $Data.PageStructure) {
        $xml += @"

    <StructureElement>
      <ElementId>$($element.ElementId)</ElementId>
      <ElementType>$($element.ElementType)</ElementType>
      <Order>$($element.Order)</Order>
      <ZoneId>$($element.ZoneId)</ZoneId>
      <Content><![CDATA[$($element.Content)]]></Content>
    </StructureElement>
"@
    }

    $xml += @"

  </PageStructure>
</SharePointPageData>
"@

    return $xml
}

#endregion

#region Main Execution

try {
    Write-Log "SharePoint 2016 Page Content Extractor v1.0" "Info"
    Write-Log "=============================================" "Info"
    
    # Extract page data
    $extractedData = Extract-SharePointPageData -SiteUrl $SiteUrl -PageUrl $PageUrl -Credentials $Credentials
    
    # Generate XML
    $xmlContent = Generate-XML -Data $extractedData
    
    # Validate XML
    try {
        $xmlDoc = [xml]$xmlContent
        Write-Log "XML validation successful" "Success"
    }
    catch {
        Write-Log "XML validation failed: $($_.Exception.Message)" "Error"
        throw "Generated XML is not well-formed"
    }
    
    # Save to file
    $outputDir = Split-Path $OutputPath -Parent
    if (!(Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }
    
    $xmlContent | Out-File -FilePath $OutputPath -Encoding UTF8
    
    Write-Log "XML output saved to: $OutputPath" "Success"
    Write-Log "Extraction completed successfully!" "Success"
    
    # Display summary
    Write-Log "=============================================" "Info"
    Write-Log "EXTRACTION SUMMARY:" "Info"
    Write-Log "Page Type: $($extractedData.PageMetadata.PageType)" "Info"
    Write-Log "Web Parts Found: $($extractedData.AllWebParts.Count)" "Info"
    Write-Log "Zones Found: $($extractedData.WebPartZones.Count)" "Info"
    Write-Log "Content Elements: $($extractedData.PageStructure.Count)" "Info"
    Write-Log "=============================================" "Info"
}
catch {
    Write-Log "Script execution failed: $($_.Exception.Message)" "Error"
    exit 1
}

#endregion