# SharePoint 2016 Metadata Extractor - Simple POC

A simple, single-file C# console application to extract SharePoint 2016 page metadata.

## Features

- Single file application (no logging, simplified for POC)
- Extracts comprehensive page metadata from SharePoint 2016
- Generates XML output matching PowerShell script format
- Prompts for credentials (works with username/password)
- Handles both SharePoint Online and on-premises authentication

## Quick Setup

1. **Prerequisites:**
   - .NET 6.0 SDK
   - Access to SharePoint 2016 site

2. **Configuration:**
   Edit the constants at the top of `SimpleSharePointExtractor.cs`:
   ```csharp
   private static readonly string SiteUrl = "https://your-sharepoint-site.com";
   private static readonly string PageUrl = "SitePages/your-page.aspx";
   private static readonly string OutputFolder = @"C:\Temp\SharePointOutput";
   ```

3. **Build and Run:**
   ```bash
   dotnet build SimpleSPExtractor.csproj
   dotnet run --project SimpleSPExtractor.csproj
   ```

## Usage

1. Run the application
2. Enter your SharePoint username when prompted
3. Enter your password (masked input)
4. Application will extract metadata and save XML file

## Output

Creates timestamped XML files in the specified output folder:
```
SharePoint_Page_Metadata_YYYYMMDD_HHMMSS.xml
```

## Sample Output Structure

```xml
<?xml version="1.0" encoding="UTF-8"?>
<SharePointPageMetadata>
  <SiteInformation>
    <SiteUrl>https://your-site.com</SiteUrl>
    <WebTitle>Site Title</WebTitle>
    <!-- ... site metadata ... -->
  </SiteInformation>
  <PageInformation>
    <PageUrl>SitePages/page.aspx</PageUrl>
    <PageTitle>Page Title</PageTitle>
    <!-- ... page metadata ... -->
    <CustomFields>
      <Field Name="CustomField1">Value1</Field>
      <!-- ... custom fields ... -->
    </CustomFields>
  </PageInformation>
</SharePointPageMetadata>
```

## Authentication

The application tries multiple authentication methods:
1. SharePoint Online credentials (for O365)
2. Network credentials (for on-premises)
3. Default credentials (as fallback)

## Files

- **SimpleSPExtractor.csproj** - Project file
- **SimpleSharePointExtractor.cs** - Single-file application with everything included
- **README_POC.md** - This documentation

## Troubleshooting

**Authentication Issues:**
- Ensure username includes domain (e.g., `user@domain.com` or `domain\user`)
- For MFA-enabled accounts, use app passwords if available
- Check SharePoint site permissions

**Connection Issues:**
- Verify site URL is accessible
- Check if page exists at the specified URL
- Ensure network connectivity to SharePoint

**Build Issues:**
- Ensure .NET 6.0 SDK is installed
- Run `dotnet restore` if package issues occur

This POC provides a simple way to extract SharePoint page metadata without complex authentication setup.