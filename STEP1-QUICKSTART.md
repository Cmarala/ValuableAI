# Step 1: Connect to SharePoint 2016 and Get Page Properties

This is the first step in building our SharePoint page extractor. We'll start simple by connecting to SharePoint and extracting basic page properties to XML.

## 🎯 What Step 1 Does

✅ **Connects to SharePoint 2016** (Online or On-Premises)  
✅ **Supports MFA authentication** with `-UseInteractiveAuth`  
✅ **Extracts basic page properties** (title, dates, authors, etc.)  
✅ **Gets all list item fields** from the page  
✅ **Outputs clean XML** with proper escaping  
✅ **Validates XML** before saving  

## 📋 Prerequisites

### Install SharePoint Client Components
```powershell
# Download and install from:
# https://www.microsoft.com/en-us/download/details.aspx?id=35585

# Or check if already installed:
Test-Path "${env:ProgramFiles}\SharePoint Client Components\16.0\Assemblies\Microsoft.SharePoint.Client.dll"
```

## 🚀 Quick Start

### For SharePoint Online with MFA:
```powershell
.\Step1-ConnectAndGetPageProperties.ps1 `
    -SiteUrl "https://contoso.sharepoint.com" `
    -PageUrl "/Pages/Home.aspx" `
    -OutputPath "step1-output.xml" `
    -UseInteractiveAuth
```

### For On-Premises SharePoint:
```powershell
$creds = Get-Credential
.\Step1-ConnectAndGetPageProperties.ps1 `
    -SiteUrl "http://sharepoint.contoso.com" `
    -PageUrl "/Pages/Home.aspx" `
    -OutputPath "step1-output.xml" `
    -Credentials $creds
```

### Using Default Windows Authentication:
```powershell
.\Step1-ConnectAndGetPageProperties.ps1 `
    -SiteUrl "http://sharepoint.contoso.com" `
    -PageUrl "/Pages/Home.aspx" `
    -OutputPath "step1-output.xml"
```

## 📝 Expected Output

The script will create an XML file like this:

```xml
<?xml version="1.0" encoding="utf-8"?>
<SharePointPageProperties>
  <ExtractionInfo>
    <ExtractionDate>2024-01-15T14:30:22</ExtractionDate>
    <ExtractedBy>YourUsername</ExtractedBy>
    <ScriptVersion>Step1-Basic</ScriptVersion>
  </ExtractionInfo>
  
  <BasicProperties>
    <PageId>1</PageId>
    <PageName>Home.aspx</PageName>
    <PageTitle>Welcome to Our Site</PageTitle>
    <ServerRelativeUrl>/Pages/Home.aspx</ServerRelativeUrl>
    <FileSize>12345</FileSize>
    <Version>2.0</Version>
    <CheckOutType>None</CheckOutType>
    <Created>2023-12-01T10:00:00</Created>
    <Modified>2024-01-15T14:00:00</Modified>
    <ContentType>Wiki Page</ContentType>
    <ContentTypeId>0x010108</ContentTypeId>
  </BasicProperties>
  
  <CreatedBy>
    <Id>1</Id>
    <LoginName>DOMAIN\user</LoginName>
    <Name>John Doe</Name>
    <Email>john.doe@contoso.com</Email>
  </CreatedBy>
  
  <ModifiedBy>
    <Id>5</Id>
    <LoginName>DOMAIN\admin</LoginName>
    <Name>Site Admin</Name>
    <Email>admin@contoso.com</Email>
  </ModifiedBy>
  
  <AllFields>
    <Field name="Title">Welcome to Our Site</Field>
    <Field name="WikiField">Page content goes here...</Field>
    <Field name="Created">2023-12-01T10:00:00Z</Field>
    <Field name="Modified">2024-01-15T14:00:00Z</Field>
    <!-- All other SharePoint fields -->
  </AllFields>
</SharePointPageProperties>
```

## 🎮 Console Output

When you run the script, you'll see:

```
[2024-01-15 14:30:15] [Info] SharePoint 2016 Page Properties Extractor - Step 1
[2024-01-15 14:30:15] [Info] =================================================
[2024-01-15 14:30:16] [Info] Loading CSOM assemblies...
[2024-01-15 14:30:16] [Success] Found CSOM assemblies at: C:\Program Files\SharePoint Client Components\16.0\Assemblies
[2024-01-15 14:30:16] [Success] Loaded: Microsoft.SharePoint.Client.dll
[2024-01-15 14:30:16] [Success] Loaded: Microsoft.SharePoint.Client.Runtime.dll
[2024-01-15 14:30:17] [Info] Connecting to SharePoint: https://contoso.sharepoint.com
[2024-01-15 14:30:17] [Info] Using interactive authentication...
[2024-01-15 14:30:20] [Success] Successfully connected to: Contoso Intranet
[2024-01-15 14:30:20] [Info] Getting page properties for: /Pages/Home.aspx
[2024-01-15 14:30:21] [Success] Found page: Home.aspx
[2024-01-15 14:30:22] [Info] Extracted 25 field values
[2024-01-15 14:30:22] [Info] Converting properties to XML...
[2024-01-15 14:30:22] [Success] XML validation successful
[2024-01-15 14:30:22] [Success] XML saved to: step1-output.xml
[2024-01-15 14:30:22] [Info] =================================================
[2024-01-15 14:30:22] [Info] EXTRACTION SUMMARY:
[2024-01-15 14:30:22] [Info] Page Name: Home.aspx
[2024-01-15 14:30:22] [Info] Page Title: Welcome to Our Site
[2024-01-15 14:30:22] [Info] Content Type: Wiki Page
[2024-01-15 14:30:22] [Info] File Size: 12345 bytes
[2024-01-15 14:30:22] [Info] Fields Extracted: 25
[2024-01-15 14:30:22] [Info] Output File: step1-output.xml
[2024-01-15 14:30:22] [Info] =================================================
[2024-01-15 14:30:22] [Success] Step 1 completed successfully!
```

## 🚨 Common Issues and Solutions

### 1. CSOM Not Found
**Error:** `CSOM assemblies not found`

**Solution:**
```powershell
# Install SharePoint Client Components
# Download from: https://www.microsoft.com/en-us/download/details.aspx?id=35585

# Or check common locations:
dir "${env:ProgramFiles}\SharePoint Client Components" -Recurse -Name "*.dll"
dir "${env:ProgramFiles}\Common Files\microsoft shared\Web Server Extensions" -Recurse -Name "*.dll"
```

### 2. Authentication Failed
**Error:** `(401) Unauthorized`

**For MFA (SharePoint Online):**
```powershell
# Always use -UseInteractiveAuth for MFA
.\Step1-ConnectAndGetPageProperties.ps1 -SiteUrl "https://contoso.sharepoint.com" -PageUrl "/Pages/Home.aspx" -OutputPath "output.xml" -UseInteractiveAuth
```

**For On-Premises:**
```powershell
# Use credentials
$creds = Get-Credential -Message "Enter DOMAIN\username and password"
.\Step1-ConnectAndGetPageProperties.ps1 -SiteUrl "http://sharepoint.local" -PageUrl "/Pages/Home.aspx" -OutputPath "output.xml" -Credentials $creds
```

### 3. Page Not Found
**Error:** `File Not Found`

**Solution:**
```powershell
# Check page URL format - must be server-relative:
"/Pages/Home.aspx"           # ✅ Correct
"/SitePages/Welcome.aspx"    # ✅ Correct  
"/sites/mysite/Pages/Test.aspx"  # ✅ Correct

"https://site.com/Pages/Home.aspx"  # ❌ Wrong - don't include domain
"Pages/Home.aspx"                   # ❌ Wrong - must start with /
```

## 📊 What You'll Get

After running Step 1, you'll have:

1. **✅ Verified CSOM connection works**
2. **✅ Confirmed authentication method**
3. **✅ Basic page metadata in XML format**
4. **✅ All SharePoint field values**
5. **✅ Properly formatted and validated XML output**

## 🔄 Next Steps

Once Step 1 works successfully:

- **Step 2**: Add web part enumeration
- **Step 3**: Extract web part properties  
- **Step 4**: Get web part content
- **Step 5**: Extract page content (Wiki, Publishing)
- **Step 6**: Preserve content order and structure

## 📁 Test Cases

Try these different page types to test Step 1:

```powershell
# Wiki page
.\Step1-ConnectAndGetPageProperties.ps1 -SiteUrl $site -PageUrl "/SitePages/Home.aspx" -OutputPath "wiki-test.xml" -UseInteractiveAuth

# Publishing page  
.\Step1-ConnectAndGetPageProperties.ps1 -SiteUrl $site -PageUrl "/Pages/Welcome.aspx" -OutputPath "publishing-test.xml" -UseInteractiveAuth

# Web Part page
.\Step1-ConnectAndGetPageProperties.ps1 -SiteUrl $site -PageUrl "/Pages/Dashboard.aspx" -OutputPath "webpart-test.xml" -UseInteractiveAuth
```

## 💡 Tips

1. **Start with a simple page** - test with a basic wiki page first
2. **Verify in browser** - make sure you can access the page manually
3. **Check permissions** - ensure you have read access to the page
4. **Test authentication** - verify your auth method works before running
5. **Use full paths** - always use complete server-relative URLs

---

**Ready to test Step 1?** Run the script with your SharePoint page and see the basic extraction in action!