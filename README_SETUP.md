# SharePoint 2016 Page Metadata Extractor - Setup Guide

This C# Console application extracts comprehensive page metadata from SharePoint 2016 sites, including sites with MFA (Multi-Factor Authentication) support.

## Current Implementation Status

✅ **Completed Components:**
- Project structure with .NET 6.0 console application
- Command line interface with System.CommandLine
- Data models for SharePoint metadata
- XML generation service matching PowerShell script output
- PnP.Framework integration for SharePoint connectivity
- Comprehensive error handling and logging

⚠️ **Authentication Configuration Required:**
The application currently requires specific authentication setup for SharePoint 2016 with MFA support.

## Prerequisites

1. **.NET 6.0 SDK or later**
2. **SharePoint 2016 site access** with appropriate permissions
3. **Azure AD App Registration** (for MFA-enabled sites)
4. **Visual Studio 2022** or **VS Code** (recommended)

## Setting Up Authentication for MFA-Enabled SharePoint Sites

### Step 1: Azure AD App Registration

For SharePoint sites with MFA, you need to register an Azure AD application:

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations**
3. Click **New registration**
4. Configure the app:
   - **Name**: SharePoint Metadata Extractor
   - **Supported account types**: Accounts in this organizational directory only
   - **Redirect URI**: Public client/native (mobile & desktop) - `http://localhost`

5. After creation, note down:
   - **Application (client) ID**
   - **Directory (tenant) ID**

6. Configure **API permissions**:
   - Add **SharePoint** > **AllSites.Read** (or AllSites.FullControl if needed)
   - Add **Microsoft Graph** > **Sites.Read.All**
   - Grant admin consent for your organization

### Step 2: Authentication Methods

The application supports multiple authentication methods:

#### Method 1: Interactive Authentication (Recommended for MFA)
```bash
dotnet run -- --site-url "https://your-sharepoint-site.com" --use-modern-auth true
```

#### Method 2: Credential-based Authentication (Non-MFA)
```bash
dotnet run -- --site-url "https://your-sharepoint-site.com" --username "user@domain.com" --password "password" --use-modern-auth false
```

#### Method 3: App-Only Authentication (Service Principal)
For unattended scenarios, configure the service with client credentials:
```bash
dotnet run -- --site-url "https://your-sharepoint-site.com" --client-id "your-app-id" --client-secret "your-secret"
```

## Configuration Files

### appsettings.json (Optional)
Create an `appsettings.json` file for default configuration:

```json
{
  "SharePointSettings": {
    "DefaultSiteUrl": "https://your-sharepoint-site.com",
    "DefaultPageUrl": "SitePages/your-page.aspx",
    "DefaultOutputFolder": "C:\\Output\\SharePointMetadata"
  },
  "AzureAD": {
    "TenantId": "your-tenant-id",
    "ClientId": "your-client-id",
    "ClientSecret": "your-client-secret"
  },
  "Logging": {
    "LogLevel": {
      "Default": "Information",
      "Microsoft": "Warning",
      "Microsoft.Hosting.Lifetime": "Information"
    }
  }
}
```

## Building and Running the Application

### 1. Restore NuGet Packages
```bash
dotnet restore
```

### 2. Build the Application
```bash
dotnet build
```

### 3. Run the Application
```bash
# With default settings
dotnet run

# With custom parameters
dotnet run -- --site-url "https://your-site.com" --page-url "SitePages/page.aspx" --output-folder "C:\Output"

# View help
dotnet run -- --help
```

## Troubleshooting Authentication Issues

### Common Issues and Solutions

#### 1. "Authentication failed" or "Access denied"
**Solution:** 
- Verify the user has read access to the SharePoint site
- Check if the page exists and is accessible
- For MFA sites, ensure Azure AD app is properly configured

#### 2. "Could not find SharePointOnlineCredentials"
**Solution:**
This indicates a package compatibility issue. Update the service to use PnP.Framework authentication:

```csharp
// In SharePointMetadataService.cs, replace credential authentication with:
var authManager = new AuthenticationManager();
context = authManager.GetACSAppOnlyContext(siteUrl, clientId, clientSecret);
```

#### 3. MFA Prompt Not Appearing
**Solution:**
- Ensure `--use-modern-auth true` is set
- Verify Azure AD app registration includes correct redirect URI
- Check browser settings allow popups

#### 4. "Page not found" errors
**Solution:**
- Verify the page URL is relative to the site (e.g., "SitePages/page.aspx")
- Check the page exists in the Site Pages library
- Ensure the page is published and not in draft mode

## Advanced Configuration

### Using Certificate-based Authentication
For enhanced security in production environments:

1. Create a self-signed certificate or use a CA-issued certificate
2. Upload the certificate to your Azure AD app registration
3. Use certificate-based authentication:

```csharp
var authManager = new AuthenticationManager();
context = authManager.GetACSAppOnlyContext(siteUrl, clientId, certificatePath, certificatePassword);
```

### Configuring for On-Premises SharePoint 2016
For on-premises SharePoint 2016 with ADFS:

```csharp
var authManager = new AuthenticationManager();
context = authManager.GetNetworkCredentialAuthenticatedContext(siteUrl, username, password, domain);
```

## Sample Output

The application generates XML files with timestamps:
```
SharePoint_Page_Metadata_20250131_143022.xml
```

Sample XML structure:
```xml
<?xml version="1.0" encoding="UTF-8"?>
<SharePointPageMetadata>
  <SiteInformation>
    <SiteUrl>https://your-site.com</SiteUrl>
    <WebTitle>Your Site Title</WebTitle>
    <!-- Additional site metadata -->
  </SiteInformation>
  <PageInformation>
    <PageUrl>SitePages/page.aspx</PageUrl>
    <PageTitle>Your Page Title</PageTitle>
    <!-- Comprehensive page metadata -->
  </PageInformation>
</SharePointPageMetadata>
```

## Performance Tips

1. **Use app-only authentication** for better performance in automated scenarios
2. **Cache authentication tokens** when making multiple requests
3. **Limit field retrieval** to only necessary fields for better performance
4. **Use parallel processing** when extracting metadata from multiple pages

## Security Best Practices

1. **Store credentials securely** using Azure Key Vault or similar services
2. **Use certificate-based authentication** instead of client secrets when possible
3. **Implement proper access controls** in your Azure AD app registration
4. **Regularly rotate client secrets** and certificates
5. **Monitor authentication logs** for suspicious activity

## Support and Troubleshooting

For additional support:

1. **Check Azure AD sign-in logs** for authentication failures
2. **Review SharePoint ULS logs** for server-side errors
3. **Enable debug logging** in the application for detailed troubleshooting
4. **Test connectivity** using tools like PnP PowerShell first

## Comparison with PowerShell Script

This C# application provides equivalent functionality to the PowerShell script with these advantages:

- **Cross-platform compatibility** (Windows, macOS, Linux)
- **Better performance** and memory management
- **Type safety** and compile-time error checking
- **Enhanced error handling** and logging
- **Easier integration** into larger .NET applications
- **Support for modern authentication** patterns

## Next Steps

1. Complete the authentication configuration based on your environment
2. Test with a simple page first to validate connectivity
3. Customize the field extraction based on your specific requirements
4. Consider adding automated testing for your specific SharePoint environment
5. Implement any additional security measures required by your organization