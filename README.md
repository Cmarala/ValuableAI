# SharePoint Web Part Reader

A simple C# .NET Framework Console Application that connects to SharePoint 2016 using WebLogin authentication and enumerates web parts on a specific page.

## Features

- Connects to SharePoint 2016 using WebLogin authentication via PnP Core AuthenticationManager
- Loads a specific page (`/sites/dev/SitePages/TestPage.aspx`)
- Retrieves and displays information about all web parts on the page:
  - Web part title
  - Instance ID (wpDef.Id)
  - Zone ID
  - Zone Index
  - Rendered div ID as "MSOZoneCell_WebPartWPQ{ZoneIndex}"

## Requirements

- .NET Framework 4.8
- SharePoint 2016 site with WebLogin authentication enabled
- Required NuGet packages (automatically restored):
  - Microsoft.SharePointOnline.CSOM (16.1.23926.12001)
  - SharePointPnPCoreOnline (3.29.2101)

## Usage

1. Update the `siteUrl` variable in `Program.cs` with your SharePoint site URL
2. Build and run the application
3. Enter your SharePoint credentials when prompted
4. The application will connect to SharePoint and display web part information

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