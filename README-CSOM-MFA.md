# SharePoint 2016 Client-Side Page Content Extractor (MFA Support)

A client-side PowerShell solution for extracting SharePoint 2016 page content from remote machines with **Multi-Factor Authentication (MFA)** support using CSOM.

## 🚀 Key Features for Client-Side Execution

✅ **Remote Client Execution** - No need to run on SharePoint server  
✅ **Multi-Factor Authentication (MFA) Support** - Works with modern authentication  
✅ **SharePoint Online & On-Premises** - Supports both environments  
✅ **Automatic CSOM Detection** - Finds assemblies automatically  
✅ **Interactive Authentication** - Browser-based authentication for MFA  
✅ **Complete Page Extraction** - All metadata, webparts, content, and structure

## 📋 Prerequisites

### Required Software
1. **SharePoint Client Components SDK**
   - Download: [SharePoint 2016 Client Components](https://www.microsoft.com/en-us/download/details.aspx?id=51679)
   - Or: [SharePoint 2013 Client Components](https://www.microsoft.com/en-us/download/details.aspx?id=35585) (compatible)

2. **PowerShell 3.0 or higher**

3. **Client Machine Requirements**
   - Windows 7/8/10/11 or Windows Server 2012+
   - .NET Framework 4.5 or higher

### Permissions Required
- **Read permissions** to SharePoint site and target page
- **Access to page libraries** (Pages, Site Pages, etc.)
- **Web permissions** for webpart information

## 🛠️ Installation Steps

### Step 1: Install SharePoint Client Components

**Option A: Download from Microsoft**
```
1. Go to: https://www.microsoft.com/en-us/download/details.aspx?id=51679
2. Download "SharePoint 2016 Client Components SDK"
3. Run the installer: sharepointclientcomponents_x64-en-us.exe
4. Follow installation wizard
```

**Option B: Install via Package Manager**
```powershell
# Using Chocolatey
choco install sharepoint-client-components

# Using PowerShell (if available)
Install-Package SharePoint.Client -Source nuget.org
```

### Step 2: Verify Installation
```powershell
# Check if assemblies are available
$csomPath = "${env:ProgramFiles}\SharePoint Client Components\16.0\Assemblies"
Test-Path "$csomPath\Microsoft.SharePoint.Client.dll"
# Should return True
```

### Step 3: Download the Script
Download `Extract-SharePointPageContent-CSOM.ps1` to your working directory.

## 🔐 Authentication Options

### For MFA Environments (Recommended)

**Interactive Authentication:**
```powershell
.\Extract-SharePointPageContent-CSOM.ps1 `
    -SiteUrl "https://contoso.sharepoint.com" `
    -PageUrl "/Pages/Home.aspx" `
    -OutputPath "C:\Temp\PageData.xml" `
    -UseInteractiveAuth
```

This will:
- Automatically detect SharePoint Online vs On-Premises
- Trigger browser-based authentication for MFA
- Handle modern authentication flows
- Support Azure AD authentication

### For Traditional Environments

**Username/Password (No MFA):**
```powershell
$creds = Get-Credential
.\Extract-SharePointPageContent-CSOM.ps1 `
    -SiteUrl "http://sharepoint.contoso.com" `
    -PageUrl "/Pages/Home.aspx" `
    -OutputPath "C:\Temp\PageData.xml" `
    -Credentials $creds
```

**Default Windows Authentication:**
```powershell
.\Extract-SharePointPageContent-CSOM.ps1 `
    -SiteUrl "http://sharepoint.contoso.com" `
    -PageUrl "/Pages/Home.aspx" `
    -OutputPath "C:\Temp\PageData.xml"
```

## 📝 Usage Examples

### Example 1: SharePoint Online with MFA
```powershell
# Most common scenario for SharePoint Online
.\Extract-SharePointPageContent-CSOM.ps1 `
    -SiteUrl "https://contoso.sharepoint.com/sites/intranet" `
    -PageUrl "/SitePages/Welcome.aspx" `
    -OutputPath "C:\Extractions\Welcome.xml" `
    -UseInteractiveAuth
```

### Example 2: On-Premises with Custom CSOM Path
```powershell
# If CSOM is installed in non-standard location
.\Extract-SharePointPageContent-CSOM.ps1 `
    -SiteUrl "http://sharepoint.contoso.com" `
    -PageUrl "/Pages/Home.aspx" `
    -OutputPath "C:\Temp\PageData.xml" `
    -CSOMPath "C:\CustomCSOM" `
    -UseInteractiveAuth
```

### Example 3: Batch Processing Multiple Pages
```powershell
# Process multiple pages with MFA authentication
$siteUrl = "https://contoso.sharepoint.com"
$pages = @(
    "/Pages/Home.aspx",
    "/Pages/About.aspx",
    "/SitePages/News.aspx"
)

foreach ($page in $pages) {
    $fileName = $page.Split('/')[-1].Replace('.aspx', '.xml')
    $outputPath = "C:\Extractions\$fileName"
    
    Write-Host "Processing: $page"
    .\Extract-SharePointPageContent-CSOM.ps1 `
        -SiteUrl $siteUrl `
        -PageUrl $page `
        -OutputPath $outputPath `
        -UseInteractiveAuth
}
```

## 🔧 Advanced Configuration

### Custom CSOM Assembly Path
If CSOM assemblies are in a custom location:
```powershell
.\Extract-SharePointPageContent-CSOM.ps1 `
    -SiteUrl "https://contoso.sharepoint.com" `
    -PageUrl "/Pages/Home.aspx" `
    -OutputPath "C:\Temp\PageData.xml" `
    -UseInteractiveAuth `
    -CSOMPath "C:\MyCustomPath\CSOM"
```

### Portable CSOM Setup
For environments where you can't install CSOM system-wide:

1. **Download CSOM assemblies manually:**
   - Microsoft.SharePoint.Client.dll
   - Microsoft.SharePoint.Client.Runtime.dll
   - Microsoft.SharePoint.Client.Publishing.dll (optional)

2. **Place in script directory:**
   ```
   C:\SharePointExtraction\
   ├── Extract-SharePointPageContent-CSOM.ps1
   ├── CSOM\
   │   ├── Microsoft.SharePoint.Client.dll
   │   ├── Microsoft.SharePoint.Client.Runtime.dll
   │   └── Microsoft.SharePoint.Client.Publishing.dll
   ```

3. **Run with custom path:**
   ```powershell
   .\Extract-SharePointPageContent-CSOM.ps1 `
       -SiteUrl "https://contoso.sharepoint.com" `
       -PageUrl "/Pages/Home.aspx" `
       -OutputPath "C:\Temp\PageData.xml" `
       -UseInteractiveAuth `
       -CSOMPath ".\CSOM"
   ```

## 🚨 Troubleshooting

### Common Issues and Solutions

#### 1. "CSOM assemblies not found"
**Error:** `CSOM assemblies not found. Please install SharePoint Client Components...`

**Solutions:**
```powershell
# Check installation paths
$paths = @(
    "${env:ProgramFiles}\SharePoint Client Components\16.0\Assemblies",
    "${env:ProgramFiles}\SharePoint Client Components\15.0\Assemblies",
    "${env:ProgramFiles(x86)}\SharePoint Client Components\16.0\Assemblies"
)

foreach ($path in $paths) {
    if (Test-Path $path) {
        Write-Host "Found CSOM at: $path"
        Get-ChildItem $path -Name "*.dll"
    }
}

# If not found, reinstall Client Components or specify custom path
```

#### 2. Authentication Failed (401 Unauthorized)
**Error:** `Authentication failed: (401) Unauthorized`

**For MFA Environments:**
```powershell
# Always use -UseInteractiveAuth for MFA
.\Extract-SharePointPageContent-CSOM.ps1 `
    -SiteUrl "https://contoso.sharepoint.com" `
    -PageUrl "/Pages/Home.aspx" `
    -OutputPath "C:\Temp\PageData.xml" `
    -UseInteractiveAuth
```

**For On-Premises:**
```powershell
# Check credentials and permissions
$creds = Get-Credential -Message "Enter domain\username and password"
.\Extract-SharePointPageContent-CSOM.ps1 `
    -SiteUrl "http://sharepoint.contoso.com" `
    -PageUrl "/Pages/Home.aspx" `
    -OutputPath "C:\Temp\PageData.xml" `
    -Credentials $creds
```

#### 3. Modern Authentication Issues
**Error:** `Interactive authentication failed...`

**Solutions:**
1. **Update Client Components:**
   - Download latest version from Microsoft
   - Uninstall old version first

2. **Check Browser Settings:**
   - Ensure default browser supports modern auth
   - Clear browser cache and cookies
   - Disable popup blockers

3. **Corporate Network Issues:**
   - Check proxy settings
   - Verify firewall allows authentication traffic
   - Test from different network if possible

#### 4. Page Not Found
**Error:** `File Not Found` or `404`

**Solutions:**
```powershell
# Verify page URL format (server-relative)
# Correct formats:
"/Pages/Home.aspx"                    # Pages library
"/SitePages/Welcome.aspx"             # Site Pages library  
"/sites/mysite/Pages/News.aspx"       # Subsite pages

# Test page accessibility first
$siteUrl = "https://contoso.sharepoint.com"
$pageUrl = "/Pages/Home.aspx"
$fullUrl = $siteUrl + $pageUrl
Write-Host "Testing: $fullUrl"
# Open in browser to verify
```

### 5. Assembly Loading Errors
**Error:** `Could not load file or assembly...`

**Solutions:**
```powershell
# Check .NET Framework version
[System.Environment]::Version

# Check PowerShell architecture (x64 vs x86)
[System.Environment]::Is64BitProcess

# Ensure matching architecture for CSOM assemblies
# Use x64 CSOM for x64 PowerShell, x86 for x86 PowerShell
```

## 🔍 Authentication Flow Details

### SharePoint Online with MFA
1. Script detects ".sharepoint.com" in URL
2. Initializes ClientContext with site URL
3. Attempts initial connection to trigger authentication
4. Browser window opens for modern authentication
5. User completes MFA process in browser
6. Authentication token cached for session
7. Script proceeds with extraction

### On-Premises SharePoint
1. Script detects non-SharePoint Online URL
2. First tries Windows authentication (current user)
3. If that fails, prompts for credentials
4. Uses traditional NTLM/Kerberos authentication
5. Script proceeds with extraction

## 📊 Performance Considerations

### For Large Sites
```powershell
# Process pages individually to avoid memory issues
$pages = Get-Content "pages-list.txt"
foreach ($page in $pages) {
    Write-Host "Processing: $page"
    .\Extract-SharePointPageContent-CSOM.ps1 -SiteUrl $siteUrl -PageUrl $page -OutputPath "Extractions\$($page.Replace('/','-')).xml" -UseInteractiveAuth
    
    # Optional: Add delay between pages
    Start-Sleep -Seconds 2
}
```

### For Multiple Sites
```powershell
# Authenticate once per site (reuse context would be ideal, but this script creates new context each time)
$sites = @(
    "https://contoso.sharepoint.com",
    "https://contoso.sharepoint.com/sites/hr",
    "https://contoso.sharepoint.com/sites/it"
)

foreach ($site in $sites) {
    Write-Host "Processing site: $site"
    # Process all pages for this site
    .\Extract-SharePointPageContent-CSOM.ps1 -SiteUrl $site -PageUrl "/Pages/Home.aspx" -OutputPath "Extractions\$(($site -split '/')[-1])-Home.xml" -UseInteractiveAuth
}
```

## 🔒 Security Best Practices

### 1. Credential Management
```powershell
# Never hardcode credentials
$creds = Get-Credential

# For automation, use secure storage
$secureString = ConvertTo-SecureString "Password" -AsPlainText -Force
$creds = New-Object System.Management.Automation.PSCredential("username", $secureString)
```

### 2. Output File Security
```powershell
# Ensure output directory is secure
$outputDir = "C:\SecureExtractions"
if (!(Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force
    # Set appropriate ACLs if needed
}
```

### 3. Logging and Auditing
The script provides comprehensive logging. For additional auditing:
```powershell
# Capture all output to log file
.\Extract-SharePointPageContent-CSOM.ps1 `
    -SiteUrl "https://contoso.sharepoint.com" `
    -PageUrl "/Pages/Home.aspx" `
    -OutputPath "C:\Temp\PageData.xml" `
    -UseInteractiveAuth | Tee-Object -FilePath "C:\Logs\extraction.log"
```

## 📦 Deployment Package

For easy deployment across multiple client machines:

**Package Structure:**
```
SharePointExtractor\
├── Extract-SharePointPageContent-CSOM.ps1
├── SharePointPageExtraction.xsd
├── Validate-ExtractedXML.ps1
├── README-CSOM-MFA.md
├── CSOM\
│   ├── Microsoft.SharePoint.Client.dll
│   ├── Microsoft.SharePoint.Client.Runtime.dll
│   └── Microsoft.SharePoint.Client.Publishing.dll
└── Examples\
    ├── batch-extract.ps1
    └── validate-all.ps1
```

**Deployment Script:**
```powershell
# Deploy-SharePointExtractor.ps1
param([string]$TargetPath = "C:\SharePointExtractor")

# Copy files
Copy-Item -Path "SharePointExtractor\*" -Destination $TargetPath -Recurse -Force

# Create shortcuts
$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$env:USERPROFILE\Desktop\SharePoint Extractor.lnk")
$Shortcut.TargetPath = "$TargetPath\Extract-SharePointPageContent-CSOM.ps1"
$Shortcut.Save()

Write-Host "SharePoint Extractor deployed to: $TargetPath"
```

## 🎯 Use Cases for Client-Side Extraction

### 1. Content Migration Planning
```powershell
# Extract all pages from a site for migration analysis
$siteUrl = "https://contoso.sharepoint.com/sites/oldsite"
$pages = @("/Pages/Home.aspx", "/Pages/About.aspx", "/SitePages/News.aspx")

foreach ($page in $pages) {
    $outputFile = "Migration\$(($page -split '/')[-1] -replace '\.aspx', '.xml')"
    .\Extract-SharePointPageContent-CSOM.ps1 -SiteUrl $siteUrl -PageUrl $page -OutputPath $outputFile -UseInteractiveAuth
}
```

### 2. Backup Before Major Changes
```powershell
# Backup critical pages before site changes
$criticalPages = @(
    "/Pages/Home.aspx",
    "/Pages/CompanyPolicies.aspx",
    "/SitePages/EmergencyContacts.aspx"
)

$backupDate = Get-Date -Format "yyyyMMdd"
foreach ($page in $criticalPages) {
    $backupFile = "Backup\$backupDate\$(($page -split '/')[-1] -replace '\.aspx', '.xml')"
    .\Extract-SharePointPageContent-CSOM.ps1 -SiteUrl $siteUrl -PageUrl $page -OutputPath $backupFile -UseInteractiveAuth
}
```

### 3. Compliance Documentation
```powershell
# Extract pages for compliance audit
$compliancePages = Get-Content "compliance-pages.txt"
foreach ($page in $compliancePages) {
    $auditFile = "Compliance\$(Get-Date -Format 'yyyy-MM')_$(($page -split '/')[-1] -replace '\.aspx', '.xml')"
    .\Extract-SharePointPageContent-CSOM.ps1 -SiteUrl $siteUrl -PageUrl $page -OutputPath $auditFile -UseInteractiveAuth
}
```

---

## ✅ Summary

The client-side CSOM version provides:

- ✅ **MFA Support** - Works with modern authentication
- ✅ **Remote Execution** - No server access required  
- ✅ **Automatic Detection** - Finds CSOM assemblies automatically
- ✅ **Flexible Authentication** - Multiple authentication methods
- ✅ **Complete Extraction** - All page data, metadata, and structure
- ✅ **Production Ready** - Error handling and comprehensive logging

Perfect for environments where:
- Multi-Factor Authentication is required
- Direct server access is not available
- Client-side execution is preferred
- Modern authentication flows are needed

**Next Steps:**
1. Install SharePoint Client Components
2. Download the CSOM script
3. Test with `-UseInteractiveAuth` parameter
4. Validate output with validation script
5. Deploy to production environments