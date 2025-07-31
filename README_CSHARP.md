# SharePoint 2016 Page Metadata Extractor (C#)

A C# Console application that connects to SharePoint 2016 and extracts comprehensive page metadata in XML format. This application replicates the functionality of the PowerShell script `Get-SharePointPageMetadata.ps1`.

## Features

- **SharePoint 2016 Connectivity**: Uses SharePoint Client Side Object Model (CSOM) to connect to SharePoint sites
- **Comprehensive Metadata Extraction**: Extracts detailed information including:
  - Site information (title, description, creation dates)
  - Page properties (title, ID, file size, checkout status)
  - Date information (created, modified, file dates)
  - Author information (created by, modified by with email addresses)
  - Content type information
  - Publishing metadata (if available)
  - Custom field values
  - Version history
- **XML Output**: Generates structured XML output matching the PowerShell script format
- **Command Line Interface**: Easy-to-use command line interface with sensible defaults
- **Robust Error Handling**: Comprehensive error handling and logging
- **Flexible Authentication**: Supports both credential-based and modern authentication methods

## Prerequisites

- .NET 6.0 or later
- Access to SharePoint 2016 site
- Appropriate permissions to read page metadata

## Installation

1. Clone or download the source code
2. Restore NuGet packages:
   ```bash
   dotnet restore
   ```
3. Build the application:
   ```bash
   dotnet build
   ```

## Usage

### Basic Usage

```bash
dotnet run
```

This will use default values:
- Site URL: `https://consulting.global.deloitteonline.com/sites/Aflac/POC`
- Page URL: `SitePages/stars.aspx`
- Output Folder: `~/Desktop/SharePointMetadataOutput`

### Custom Parameters

```bash
dotnet run -- --site-url "https://your-sharepoint-site.com" --page-url "SitePages/your-page.aspx" --output-folder "C:\Output"
```

### With Authentication

```bash
dotnet run -- --site-url "https://your-sharepoint-site.com" --username "your-username" --password "your-password" --use-modern-auth false
```

### Command Line Options

| Option | Description | Default |
|--------|-------------|---------|
| `--site-url` | The SharePoint site URL | `https://consulting.global.deloitteonline.com/sites/Aflac/POC` |
| `--page-url` | The relative URL of the page | `SitePages/stars.aspx` |
| `--output-folder` | Output folder for XML files | `~/Desktop/SharePointMetadataOutput` |
| `--username` | Username for authentication | (optional) |
| `--password` | Password for authentication | (optional) |
| `--use-modern-auth` | Use modern authentication | `true` |

### Help

```bash
dotnet run -- --help
```

## Output

The application generates an XML file with a timestamp in the filename:
```
SharePoint_Page_Metadata_YYYYMMDD_HHMMSS.xml
```

### Sample XML Structure

```xml
<?xml version="1.0" encoding="UTF-8"?>
<SharePointPageMetadata>
  <SiteInformation>
    <SiteUrl>https://your-site.com</SiteUrl>
    <WebTitle>Site Title</WebTitle>
    <WebDescription>Site Description</WebDescription>
    <WebId>guid-here</WebId>
    <WebServerRelativeUrl>/sites/yoursite</WebServerRelativeUrl>
    <WebCreated>2023-01-01T00:00:00Z</WebCreated>
    <WebLastModified>2023-12-01T00:00:00Z</WebLastModified>
  </SiteInformation>
  <PageInformation>
    <PageUrl>SitePages/page.aspx</PageUrl>
    <PageName>page.aspx</PageName>
    <PageTitle>Page Title</PageTitle>
    <PageId>1</PageId>
    <PageUniqueId>guid-here</PageUniqueId>
    <PageServerRelativeUrl>/sites/yoursite/SitePages/page.aspx</PageServerRelativeUrl>
    <PageFileSize>12345</PageFileSize>
    <Dates>
      <Created>2023-01-01T00:00:00Z</Created>
      <Modified>2023-12-01T00:00:00Z</Modified>
      <FileCreated>2023-01-01T00:00:00Z</FileCreated>
      <FileModified>2023-12-01T00:00:00Z</FileModified>
    </Dates>
    <Authors>
      <CreatedBy>
        <LoginName>user@domain.com</LoginName>
        <Email>user@domain.com</Email>
      </CreatedBy>
      <ModifiedBy>
        <LoginName>editor@domain.com</LoginName>
        <Email>editor@domain.com</Email>
      </ModifiedBy>
    </Authors>
    <ContentType>
      <Name>Site Page</Name>
      <Id>content-type-id</Id>
    </ContentType>
    <PublishingInformation>
      <!-- Publishing fields if available -->
    </PublishingInformation>
    <CustomFields>
      <Field Name="CustomField1">Value1</Field>
      <Field Name="CustomField2">Value2</Field>
    </CustomFields>
    <VersionInformation>
      <UIVersion>1.0</UIVersion>
      <VersionNumber>1</VersionNumber>
      <VersionHistory>
        <Version VersionLabel="1.0" Size="12345" Created="2023-01-01T00:00:00Z" CreatedBy="user@domain.com" />
      </VersionHistory>
    </VersionInformation>
  </PageInformation>
</SharePointPageMetadata>
```

## Architecture

The application is structured with the following components:

- **Program.cs**: Main entry point with command line argument parsing
- **Models/SharePointPageMetadata.cs**: Data models for metadata structure
- **Services/SharePointMetadataService.cs**: SharePoint connectivity and metadata extraction
- **Services/XmlMetadataGenerator.cs**: XML generation and file output

## Error Handling

The application includes comprehensive error handling for:
- SharePoint connectivity issues
- Authentication failures
- Missing pages or permissions
- File system access problems
- Data conversion errors

All errors are logged with appropriate detail levels for troubleshooting.

## Authentication Notes

### Modern Authentication
Modern authentication (default) attempts to use interactive browser-based login. This is recommended for SharePoint Online and modern SharePoint 2016 configurations with ADFS.

### Credential-Based Authentication
When providing username and password, the application uses `SharePointOnlineCredentials` for authentication. This works with:
- SharePoint Online (Office 365)
- SharePoint 2016 with appropriate authentication configuration

### On-Premises SharePoint 2016
For traditional on-premises SharePoint 2016 with Windows Authentication, you may need to:
1. Run the application on a domain-joined machine
2. Use domain credentials
3. Ensure the SharePoint site is configured for the authentication method

## Dependencies

- **Microsoft.SharePointOnline.CSOM**: SharePoint Client Side Object Model
- **System.CommandLine**: Command line argument parsing
- **Microsoft.Extensions.Logging**: Logging framework

## Comparison with PowerShell Script

This C# application provides equivalent functionality to the PowerShell script with these advantages:
- **Cross-platform**: Runs on Windows, macOS, and Linux
- **Performance**: Generally faster execution
- **Type Safety**: Compile-time error checking
- **Distribution**: Can be compiled to a single executable
- **Integration**: Easier to integrate into larger .NET applications

## Troubleshooting

### Common Issues

1. **Authentication Failures**
   - Ensure credentials are correct
   - Check if MFA is required (use modern auth)
   - Verify site URL is accessible

2. **Page Not Found**
   - Verify the page URL is correct and relative to the site
   - Check permissions to access the page
   - Ensure the page exists in the Site Pages library

3. **Permission Errors**
   - Ensure the user has at least read permissions to the site and page
   - Check if the Site Pages library exists and is accessible

4. **Network/Connectivity Issues**
   - Verify network connectivity to SharePoint
   - Check firewall and proxy settings
   - Ensure SharePoint site is accessible from the machine

### Logging

The application uses structured logging. To increase log verbosity for troubleshooting, modify the logging configuration in `Program.cs`:

```csharp
builder.AddConsole().SetMinimumLevel(LogLevel.Debug);
```

## Contributing

Feel free to submit issues, feature requests, or pull requests to improve the application.

## License

This application is provided as-is for educational and operational purposes.