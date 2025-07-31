# SharePoint Web Part Reader

A simple C# .NET Framework Console Application that connects to SharePoint 2016 using WebLogin authentication and enumerates web parts on a specific page.

## Features

- Connects to SharePoint 2016 using WebLogin authentication with MFA support via PnP Core AuthenticationManager
- Opens browser window for interactive authentication (supports MFA/2FA)
- Loads a specific page (`/sites/dev/SitePages/TestPage.aspx`)
- Retrieves and displays information about all web parts on the page:
  - Web part title
  - Instance ID (wpDef.Id)
  - Zone ID
  - Zone Index
  - Rendered div ID as "MSOZoneCell_WebPartWPQ{ZoneIndex}"

## Requirements

- .NET Framework 4.8
- SharePoint 2016 site with WebLogin authentication and MFA enabled
- Default web browser (for interactive authentication)
- Required NuGet packages (automatically restored):
  - Microsoft.SharePointOnline.CSOM (16.1.23926.12001)
  - SharePointPnPCoreOnline (3.29.2101)

## Usage

1. Update the `siteUrl` variable in `Program.cs` with your SharePoint site URL
2. Build and run the application
3. A browser window will automatically open for SharePoint authentication
4. Complete the login process including MFA/2FA steps in the browser
5. The application will connect to SharePoint and display web part information

## CSOM References Used

- Microsoft.SharePoint.Client.dll
- Microsoft.SharePoint.Client.Runtime.dll
- Microsoft.SharePoint.Client.WebParts.dll

## Configuration

Before running, update the following in `Program.cs`:
```csharp
string siteUrl = "https://your-sharepoint-site.com";  // Update with your SharePoint site URL
```

The page URL is currently set to `/sites/dev/SitePages/TestPage.aspx` but can be modified as needed.