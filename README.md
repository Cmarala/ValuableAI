# SharePoint 2016 Comprehensive Page Content Extractor

A comprehensive PowerShell solution for extracting complete page information from SharePoint 2016, including all page metadata, webparts, content, and structure, outputting to a well-formed XML file.

## Features

This solution extracts:

✅ **All page metadata** (title, creation/modification dates, content type, etc.)  
✅ **All webparts and their properties** including embedded webparts  
✅ **All content inside webparts** (Content Editor, Text, HTML Form, etc.)  
✅ **All content written directly on page** (irrespective of page type)  
✅ **The order of content and webparts on page**  
✅ **Well-defined XML output** without errors and with proper validation

## Supported Page Types

- **Wiki Pages** - Extracts WikiField content and all webparts
- **Publishing Pages** - Extracts PublishingPageContent and related fields
- **Web Part Pages** - Extracts all webpart zones and webparts
- **Application Pages** - Basic content extraction
- **Custom Pages** - Fallback extraction for any .aspx page

## Requirements

### System Requirements
- **SharePoint 2016** (On-premises or hybrid)
- **PowerShell 3.0 or higher**
- **SharePoint PowerShell Snap-in** or **SharePoint Client-Side Object Model (CSOM)**
- **.NET Framework 4.5 or higher**

### Permissions Required
- **Read permissions** to the target site and page
- **Web permissions** to access web part information
- **List permissions** to read page metadata from Pages library

### Installation Prerequisites

#### Option 1: Run on SharePoint Server (Recommended)
If running directly on a SharePoint 2016 server:
```powershell
# SharePoint PowerShell snapin is automatically available
# No additional installation required
```

#### Option 2: Remote Execution with CSOM
If running from a remote machine:

1. **Download SharePoint 2016 Client Components SDK**
   - Download from Microsoft Download Center
   - Or install via NuGet in a .NET project

2. **Install required assemblies** (typical path):
   ```
   C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\
   - Microsoft.SharePoint.Client.dll
   - Microsoft.SharePoint.Client.Runtime.dll
   - Microsoft.SharePoint.Client.Publishing.dll
   ```

## Quick Start

### 1. Download the Files
Download the following files to your working directory:
- `Extract-SharePointPageContent.ps1` - Main PowerShell script
- `SharePointPageExtraction.xsd` - XML schema definition

### 2. Basic Usage
```powershell
# Simple extraction using current user credentials
.\Extract-SharePointPageContent.ps1 `
    -SiteUrl "http://sharepoint.contoso.com" `
    -PageUrl "/Pages/Home.aspx" `
    -OutputPath "C:\Temp\PageData.xml"
```

### 3. Using Custom Credentials
```powershell
# Get credentials and extract
$creds = Get-Credential
.\Extract-SharePointPageContent.ps1 `
    -SiteUrl "http://sharepoint.contoso.com" `
    -PageUrl "/SitePages/MyWikiPage.aspx" `
    -OutputPath "C:\Temp\WikiPage.xml" `
    -Credentials $creds
```

## Detailed Usage Instructions

### Step 1: Prepare Your Environment

#### For SharePoint Server Execution
1. Log into your SharePoint 2016 server
2. Open **SharePoint 2016 Management Shell** as Administrator
3. Navigate to the directory containing the script

#### For Remote Execution
1. Ensure CSOM assemblies are installed
2. Open **PowerShell ISE** or **PowerShell** as Administrator
3. Verify assemblies can be loaded:
```powershell
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
```

### Step 2: Identify Your Target Page

#### Get the Page URL
1. Navigate to your SharePoint page in a web browser
2. Copy the server-relative URL (everything after the domain)
   - Example: `/Pages/Home.aspx`
   - Example: `/sites/mysite/SitePages/Wiki.aspx`
   - Example: `/subweb/Lists/MyList/DispForm.aspx?ID=1`

#### Common Page Locations
- **Pages Library**: `/Pages/PageName.aspx`
- **Site Pages**: `/SitePages/PageName.aspx`
- **Subsite Pages**: `/subsite/Pages/PageName.aspx`

### Step 3: Run the Extraction

#### Example 1: Extract a Wiki Page
```powershell
.\Extract-SharePointPageContent.ps1 `
    -SiteUrl "http://intranet.contoso.com" `
    -PageUrl "/SitePages/TeamWiki.aspx" `
    -OutputPath "C:\Extractions\TeamWiki.xml"
```

#### Example 2: Extract a Publishing Page with Credentials
```powershell
$credentials = Get-Credential -Message "Enter SharePoint credentials"
.\Extract-SharePointPageContent.ps1 `
    -SiteUrl "http://portal.contoso.com" `
    -PageUrl "/Pages/NewsArticle.aspx" `
    -OutputPath "C:\Extractions\NewsArticle.xml" `
    -Credentials $credentials
```

#### Example 3: Extract Multiple Pages (Batch Processing)
```powershell
$siteUrl = "http://sharepoint.contoso.com"
$pages = @(
    "/Pages/Home.aspx",
    "/Pages/About.aspx",
    "/SitePages/Wiki.aspx"
)

foreach ($page in $pages) {
    $fileName = $page.Split('/')[-1].Replace('.aspx', '.xml')
    $outputPath = "C:\Extractions\$fileName"
    
    Write-Host "Extracting: $page"
    .\Extract-SharePointPageContent.ps1 `
        -SiteUrl $siteUrl `
        -PageUrl $page `
        -OutputPath $outputPath
}
```

### Step 4: Verify the Output

The script will display progress information and a summary:
```
[2024-01-15 14:30:15] [Info] SharePoint 2016 Page Content Extractor v1.0
[2024-01-15 14:30:15] [Info] =============================================
[2024-01-15 14:30:16] [Success] Connected to site: Contoso Intranet
[2024-01-15 14:30:17] [Success] Found page file: Home.aspx
[2024-01-15 14:30:18] [Info] Extracting page metadata...
[2024-01-15 14:30:19] [Info] Extracting page content...
[2024-01-15 14:30:20] [Info] Extracting web parts from page...
[2024-01-15 14:30:20] [Info] Found 3 web parts
[2024-01-15 14:30:21] [Info] Extracted web part: Welcome Text (Type: ContentEditorWebPart)
[2024-01-15 14:30:22] [Success] XML validation successful
[2024-01-15 14:30:22] [Success] XML output saved to: C:\Temp\PageData.xml
[2024-01-15 14:30:22] [Success] Extraction completed successfully!
[2024-01-15 14:30:22] [Info] =============================================
[2024-01-15 14:30:22] [Info] EXTRACTION SUMMARY:
[2024-01-15 14:30:22] [Info] Page Type: Wiki
[2024-01-15 14:30:22] [Info] Web Parts Found: 3
[2024-01-15 14:30:22] [Info] Zones Found: 2
[2024-01-15 14:30:22] [Info] Content Elements: 5
[2024-01-15 14:30:22] [Info] =============================================
```

## Understanding the XML Output

The extracted XML follows a comprehensive schema with the following main sections:

### XML Structure Overview
```xml
<SharePointPageData xmlns="http://schemas.sharepoint.com/page-extraction/2016">
  <ExtractionInfo>        <!-- Script and extraction metadata -->
  <SiteInformation>       <!-- Site and web information -->
  <PageMetadata>          <!-- Page properties and metadata -->
  <PageContent>           <!-- Raw page content and text -->
  <WebPartZones>          <!-- Web part zone definitions -->
  <AllWebParts>           <!-- Complete web part information -->
  <PageStructure>         <!-- Order and structure of page elements -->
</SharePointPageData>
```

### Key Sections Explained

#### 1. ExtractionInfo
Contains metadata about the extraction process:
```xml
<ExtractionInfo>
  <ScriptVersion>1.0.0</ScriptVersion>
  <SharePointVersion>SharePoint 2016</SharePointVersion>
  <ExtractionUser>DOMAIN\username</ExtractionUser>
  <ExtractionMethod>PowerShell CSOM</ExtractionMethod>
  <Timestamp>2024-01-15T14:30:22</Timestamp>
</ExtractionInfo>
```

#### 2. PageMetadata
Complete page metadata including custom fields:
```xml
<PageMetadata>
  <PageId>1</PageId>
  <PageName>Home.aspx</PageName>
  <PageTitle>Welcome to Our Site</PageTitle>
  <PageType>Wiki</PageType>
  <ContentType>Wiki Page</ContentType>
  <Created>2023-12-01T10:00:00</Created>
  <Modified>2024-01-15T14:00:00</Modified>
  <!-- All list item fields included -->
</PageMetadata>
```

#### 3. WebParts Section
Detailed information about each web part:
```xml
<AllWebParts count="3">
  <WebPart>
    <WebPartId>12345678-1234-1234-1234-123456789abc</WebPartId>
    <Title>Content Editor</Title>
    <TypeName>Microsoft.SharePoint.WebPartPages.ContentEditorWebPart</TypeName>
    <ZoneId>wpz</ZoneId>
    <Properties count="15">
      <Property name="Content" type="String"><![CDATA[<div>Hello World</div>]]></Property>
      <!-- All web part properties -->
    </Properties>
    <Content>
      <HtmlContent><![CDATA[<div>Hello World</div>]]></HtmlContent>
      <RawContent><![CDATA[...]]></RawContent>
    </Content>
  </WebPart>
</AllWebParts>
```

#### 4. PageStructure
Maintains the order of all content elements:
```xml
<PageStructure count="5">
  <StructureElement>
    <ElementId>WikiContent</ElementId>
    <ElementType>Content</ElementType>
    <Order>0</Order>
    <Content><![CDATA[Wiki page content...]]></Content>
  </StructureElement>
  <StructureElement>
    <ElementId>12345678-1234-1234-1234-123456789abc</ElementId>
    <ElementType>WebPart</ElementType>
    <Order>1</Order>
    <ZoneId>wpz</ZoneId>
    <Content>Content Editor</Content>
  </StructureElement>
</PageStructure>
```

## Advanced Configuration

### Handling Different Authentication Scenarios

#### Windows Authentication (Default)
```powershell
# Uses current user context - no credentials needed
.\Extract-SharePointPageContent.ps1 -SiteUrl "http://sharepoint.local" -PageUrl "/Pages/Home.aspx" -OutputPath "output.xml"
```

#### Forms-Based Authentication
```powershell
$creds = Get-Credential -Message "Enter FBA credentials"
.\Extract-SharePointPageContent.ps1 -SiteUrl "http://sharepoint.contoso.com" -PageUrl "/Pages/Home.aspx" -OutputPath "output.xml" -Credentials $creds
```

#### SharePoint Online (Office 365)
```powershell
# For SharePoint Online, ensure you have SharePointOnlineCredentials
$creds = Get-Credential
.\Extract-SharePointPageContent.ps1 -SiteUrl "https://contoso.sharepoint.com" -PageUrl "/Pages/Home.aspx" -OutputPath "output.xml" -Credentials $creds
```

### Bulk Extraction Script

Create a CSV file (`pages.csv`) with your pages:
```csv
SiteUrl,PageUrl,OutputFileName
http://sharepoint.local,/Pages/Home.aspx,Home.xml
http://sharepoint.local,/Pages/About.aspx,About.xml
http://sharepoint.local/sites/hr,/SitePages/Policies.aspx,HRPolicies.xml
```

Then use this script:
```powershell
$pages = Import-Csv "pages.csv"
$outputDir = "C:\SharePointExtractions"

foreach ($page in $pages) {
    $outputPath = Join-Path $outputDir $page.OutputFileName
    Write-Host "Processing: $($page.PageUrl) from $($page.SiteUrl)"
    
    try {
        .\Extract-SharePointPageContent.ps1 `
            -SiteUrl $page.SiteUrl `
            -PageUrl $page.PageUrl `
            -OutputPath $outputPath
        Write-Host "✓ Success: $outputPath" -ForegroundColor Green
    }
    catch {
        Write-Host "✗ Failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}
```

## Troubleshooting

### Common Issues and Solutions

#### 1. "Failed to load SharePoint PowerShell snapin"
**Problem**: SharePoint PowerShell snapin not available  
**Solution**: 
- Run on SharePoint server, or
- Install SharePoint Client Components SDK
- Use CSOM assemblies instead

#### 2. "Access Denied" Errors
**Problem**: Insufficient permissions  
**Solution**:
- Verify read permissions to site and page
- Check if page exists and is accessible
- Use appropriate credentials with `-Credentials` parameter

#### 3. "Could not export web part XML"
**Problem**: Web part export mode is set to "None"  
**Solution**:
- This is expected for security-restricted web parts
- Content will still be extracted from properties
- Check web part settings if you need full XML export

#### 4. "Page not found" Errors
**Problem**: Incorrect page URL format  
**Solution**:
- Use server-relative URLs (start with `/`)
- Include the full path including subsites
- Verify the page exists in the browser first

#### 5. Assembly Loading Issues
**Problem**: CSOM assemblies not found  
**Solution**:
```powershell
# Verify assembly path
$assemblyPath = "${env:ProgramFiles}\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Test-Path $assemblyPath

# If not found, install SharePoint Client Components
# Or copy assemblies to a local folder and update the script path
```

#### 6. Memory Issues with Large Pages
**Problem**: Out of memory with pages containing many web parts  
**Solution**:
- Process pages individually rather than in batch
- Increase PowerShell memory limits:
```powershell
# Increase memory limit
$env:PSMemoryQuotaKB = 2097152  # 2GB
```

### Debugging Mode

To enable verbose logging, modify the script to include debug information:
```powershell
# Add this at the beginning of the script
$VerbosePreference = "Continue"
$DebugPreference = "Continue"

# Run with -Verbose flag
.\Extract-SharePointPageContent.ps1 -SiteUrl "http://sharepoint.local" -PageUrl "/Pages/Home.aspx" -OutputPath "output.xml" -Verbose
```

### Performance Optimization

For better performance when processing multiple pages:

1. **Use Persistent Connections**:
   - Modify script to reuse ClientContext objects
   - Batch multiple page extractions in one session

2. **Parallel Processing**:
```powershell
# Process multiple pages in parallel
$pages | ForEach-Object -Parallel {
    .\Extract-SharePointPageContent.ps1 -SiteUrl $_.SiteUrl -PageUrl $_.PageUrl -OutputPath $_.OutputPath
} -ThrottleLimit 5
```

## XML Schema Validation

To validate your extracted XML against the schema:

```powershell
# Load the XML and schema
$xmlDoc = [xml](Get-Content "PageData.xml")
$schemaSet = New-Object System.Xml.Schema.XmlSchemaSet
$schemaSet.Add("http://schemas.sharepoint.com/page-extraction/2016", "SharePointPageExtraction.xsd")

# Validate
$xmlDoc.Schemas = $schemaSet
try {
    $xmlDoc.Validate($null)
    Write-Host "✓ XML is valid according to schema" -ForegroundColor Green
}
catch {
    Write-Host "✗ XML validation failed: $($_.Exception.Message)" -ForegroundColor Red
}
```

## What Gets Extracted

### Page Metadata
- Page ID, name, title, URLs
- Page type (Wiki, Publishing, WebPart, etc.)
- Content type information
- Creation and modification dates
- Author and modifier information
- File size and version information
- All list item fields and custom properties

### Web Part Information
- Web part GUID, title, description
- Web part type and assembly information
- Zone placement and ordering
- All web part properties
- Visibility and permission settings
- Export settings and capabilities
- Full web part XML (if exportable)

### Content Extraction
- **Wiki Pages**: WikiField content and plain text
- **Publishing Pages**: PublishingPageContent and related fields
- **Content Editor Web Parts**: HTML content and content links
- **Text Web Parts**: Text content
- **All Web Parts**: Content-related properties
- **Raw Page Content**: Complete ASPX file content

### Structure Information
- Order of all content elements on the page
- Web part zone definitions and organization
- Relationship between zones and web parts
- Sequential ordering for content reconstruction

## Security Considerations

### Data Privacy
- The extracted XML contains all page content including potentially sensitive information
- Store extracted files securely
- Consider encryption for sensitive content

### Credential Management
- Never hardcode credentials in scripts
- Use `Get-Credential` for interactive credential input
- Consider using service accounts for automated extractions

### Web Part Security
- Some web parts may be configured as non-exportable for security reasons
- The script respects these settings and will note when export is not possible
- Sensitive web part properties may be excluded from extraction

## Use Cases

### Content Migration
- **SharePoint Upgrades**: Extract pages before upgrading to preserve exact content and structure
- **Platform Migration**: Move from SharePoint 2016 to SharePoint Online or other platforms
- **Site Restructuring**: Backup complete page content before major site changes

### Documentation and Compliance
- **Content Auditing**: Document all content and web parts for compliance purposes
- **Backup and Recovery**: Create detailed backups of important pages
- **Change Management**: Track page content changes over time

### Development and Testing
- **Environment Synchronization**: Replicate production page content in development environments
- **Web Part Analysis**: Understand web part usage and configuration across sites
- **Content Analysis**: Analyze content patterns and usage across multiple pages

## Limitations

### Known Limitations
1. **Embedded Web Parts**: Complex nested web part scenarios may not be fully captured
2. **Custom Web Parts**: Properties of custom-developed web parts may need additional handling
3. **Dynamic Content**: Content loaded via JavaScript or AJAX is not captured
4. **File Attachments**: Referenced files are not downloaded, only URLs are captured
5. **User Context**: Some content may be user-specific and appear differently for different users

### SharePoint Version Compatibility
- **Designed for SharePoint 2016**: Primary target platform
- **SharePoint 2013**: Should work with minor modifications
- **SharePoint Online**: Works with appropriate authentication
- **SharePoint 2019/Subscription**: Should work but may need CSOM assembly updates

## Support and Contributing

### Getting Help
1. Check the troubleshooting section above
2. Verify your environment meets the requirements
3. Test with a simple page first before processing complex pages

### Extending the Script
The script is designed to be modular and extensible:

- **Add new web part types**: Extend the `Get-WebPartContent` function
- **Custom content extraction**: Modify the `Extract-PageContent` function
- **Additional metadata**: Enhance the `Extract-PageMetadata` function
- **Different output formats**: Create alternative output functions

### Best Practices
1. **Test thoroughly** with non-production content first
2. **Start small** - extract individual pages before batch processing
3. **Monitor resource usage** - large pages can consume significant memory
4. **Validate output** - always check the generated XML for completeness
5. **Document customizations** - keep track of any script modifications

---

**Version**: 1.0.0  
**Last Updated**: January 2024  
**Compatibility**: SharePoint 2016, PowerShell 3.0+