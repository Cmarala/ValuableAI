<#
.SYNOPSIS
    Step 1: Connect to SharePoint 2016 and Extract Basic Page Properties
    
.DESCRIPTION
    This is the first step in building the SharePoint page extractor.
    It connects to SharePoint 2016 using CSOM and extracts basic page properties to XML.
    
.PARAMETER SiteUrl
    The URL of the SharePoint site
    
.PARAMETER PageUrl
    The server-relative URL of the page (e.g., "/Pages/Home.aspx")
    
.PARAMETER OutputPath
    Path where the XML file will be saved
    
.PARAMETER UseInteractiveAuth
    Use interactive authentication for MFA
    
.PARAMETER Credentials
    Username/password credentials
    
.EXAMPLE
    .\Step1-ConnectAndGetPageProperties.ps1 -SiteUrl "https://contoso.sharepoint.com" -PageUrl "/Pages/Home.aspx" -OutputPath "page-properties.xml" -UseInteractiveAuth
    
.EXAMPLE
    $creds = Get-Credential
    .\Step1-ConnectAndGetPageProperties.ps1 -SiteUrl "http://sharepoint.local" -PageUrl "/Pages/Home.aspx" -OutputPath "page-properties.xml" -Credentials $creds
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$PageUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$OutputPath,
    
    [Parameter(Mandatory=$false)]
    [switch]$UseInteractiveAuth,
    
    [Parameter(Mandatory=$false)]
    [System.Management.Automation.PSCredential]$Credentials
)

#region Helper Functions

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "Info"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        "Error" { "Red" }
        "Warning" { "Yellow" }
        "Success" { "Green" }
        default { "White" }
    }
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
}

function Load-CSOMAssemblies {
    Write-Log "Loading CSOM assemblies..." "Info"
    
    # Possible CSOM assembly locations
    $possiblePaths = @(
        "${env:ProgramFiles}\Common Files\microsoft shared\Web Server Extensions\16\ISAPI",
        "${env:ProgramFiles}\Common Files\microsoft shared\Web Server Extensions\15\ISAPI",
        "${env:ProgramFiles(x86)}\Common Files\microsoft shared\Web Server Extensions\16\ISAPI",
        "${env:ProgramFiles(x86)}\Common Files\microsoft shared\Web Server Extensions\15\ISAPI",
        "${env:ProgramFiles}\SharePoint Client Components\16.0\Assemblies",
        "${env:ProgramFiles}\SharePoint Client Components\15.0\Assemblies",
        "${env:ProgramFiles(x86)}\SharePoint Client Components\16.0\Assemblies",
        "${env:ProgramFiles(x86)}\SharePoint Client Components\15.0\Assemblies",
        ".\CSOM"
    )
    
    $requiredAssemblies = @(
        "Microsoft.SharePoint.Client.dll",
        "Microsoft.SharePoint.Client.Runtime.dll"
    )
    
    $assemblyPath = ""
    foreach ($path in $possiblePaths) {
        if (Test-Path $path) {
            $allFound = $true
            foreach ($assembly in $requiredAssemblies) {
                if (-not (Test-Path (Join-Path $path $assembly))) {
                    $allFound = $false
                    break
                }
            }
            if ($allFound) {
                $assemblyPath = $path
                break
            }
        }
    }
    
    if (-not $assemblyPath) {
        throw "CSOM assemblies not found. Please install SharePoint Client Components."
    }
    
    Write-Log "Found CSOM assemblies at: $assemblyPath" "Success"
    
    # Load assemblies
    foreach ($assembly in $requiredAssemblies) {
        $assemblyFile = Join-Path $assemblyPath $assembly
        Add-Type -Path $assemblyFile
        Write-Log "Loaded: $assembly" "Success"
    }
}

function Connect-SharePoint {
    param(
        [string]$SiteUrl,
        [bool]$UseInteractive,
        [System.Management.Automation.PSCredential]$Creds
    )
    
    Write-Log "Connecting to SharePoint: $SiteUrl" "Info"
    
    $context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
    
    if ($UseInteractive) {
        Write-Log "Using interactive authentication..." "Info"
        # For SharePoint Online, this will trigger modern auth
        # For on-premises, will use Windows auth or prompt for credentials
    }
    elseif ($Creds) {
        Write-Log "Using provided credentials..." "Info"
        if ($SiteUrl.Contains(".sharepoint.com")) {
            # SharePoint Online
            $context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Creds.UserName, $Creds.Password)
        } else {
            # On-premises
            $context.Credentials = $Creds.GetNetworkCredential()
        }
    }
    else {
        Write-Log "Using default credentials..." "Info"
        $context.Credentials = [System.Net.CredentialCache]::DefaultCredentials
    }
    
    # Test connection
    try {
        $context.Load($context.Web)
        $context.ExecuteQuery()
        Write-Log "Successfully connected to: $($context.Web.Title)" "Success"
        return $context
    }
    catch {
        Write-Log "Connection failed: $($_.Exception.Message)" "Error"
        throw
    }
}

function Get-PageProperties {
    param(
        $Context,
        [string]$PageUrl
    )
    
    Write-Log "Getting page properties for: $PageUrl" "Info"
    
    try {
        # Get the page file
        $file = $Context.Web.GetFileByServerRelativeUrl($PageUrl)
        $Context.Load($file)
        $Context.ExecuteQuery()
        
        Write-Log "Found page: $($file.Name)" "Success"
        
        # Get the list item
        $listItem = $file.ListItemAllFields
        $Context.Load($listItem)
        $Context.Load($file.Author)
        $Context.Load($file.ModifiedBy)
        $Context.ExecuteQuery()
        
        # Get content type if available
        $contentType = ""
        $contentTypeId = ""
        try {
            if ($listItem.ContentType) {
                $Context.Load($listItem.ContentType)
                $Context.ExecuteQuery()
                $contentType = $listItem.ContentType.Name
                $contentTypeId = $listItem.ContentType.Id.ToString()
            }
        }
        catch {
            Write-Log "Could not load content type" "Warning"
        }
        
        # Build properties object
        $properties = @{
            # Basic file properties
            PageId = $listItem.Id
            PageName = $file.Name
            PageTitle = if ($listItem["Title"]) { $listItem["Title"] } else { $file.Name }
            ServerRelativeUrl = $file.ServerRelativeUrl
            
            # File metadata
            Created = $listItem["Created"]
            Modified = $listItem["Modified"]
            FileSize = $file.Length
            Version = if ($listItem["_UIVersionString"]) { $listItem["_UIVersionString"] } else { "1.0" }
            CheckOutType = $file.CheckOutType.ToString()
            
            # Content type
            ContentType = $contentType
            ContentTypeId = $contentTypeId
            
            # Author information
            CreatedBy = @{
                Id = if ($file.Author.Id) { $file.Author.Id } else { 0 }
                LoginName = if ($file.Author.LoginName) { $file.Author.LoginName } else { "" }
                Name = if ($file.Author.Name) { $file.Author.Name } else { "" }
                Email = if ($file.Author.Email) { $file.Author.Email } else { "" }
            }
            
            ModifiedBy = @{
                Id = if ($file.ModifiedBy.Id) { $file.ModifiedBy.Id } else { 0 }
                LoginName = if ($file.ModifiedBy.LoginName) { $file.ModifiedBy.LoginName } else { "" }
                Name = if ($file.ModifiedBy.Name) { $file.ModifiedBy.Name } else { "" }
                Email = if ($file.ModifiedBy.Email) { $file.ModifiedBy.Email } else { "" }
            }
            
            # All list item fields
            AllFields = @{}
        }
        
        # Extract all field values
        foreach ($fieldName in $listItem.FieldValues.Keys) {
            try {
                $fieldValue = $listItem[$fieldName]
                if ($fieldValue -ne $null) {
                    $properties.AllFields[$fieldName] = $fieldValue.ToString()
                }
            }
            catch {
                Write-Log "Could not extract field: $fieldName" "Warning"
            }
        }
        
        Write-Log "Extracted $($properties.AllFields.Count) field values" "Info"
        return $properties
        
    }
    catch {
        Write-Log "Error getting page properties: $($_.Exception.Message)" "Error"
        throw
    }
}

function Convert-ToXml {
    param($Properties)
    
    Write-Log "Converting properties to XML..." "Info"
    
    # Helper function to escape XML content
    function Escape-Xml {
        param([string]$Text)
        if (-not $Text) { return "" }
        return [System.Security.SecurityElement]::Escape($Text)
    }
    
    $xml = @"
<?xml version="1.0" encoding="utf-8"?>
<SharePointPageProperties>
  <ExtractionInfo>
    <ExtractionDate>$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ss')</ExtractionDate>
    <ExtractedBy>$([Environment]::UserName)</ExtractedBy>
    <ScriptVersion>Step1-Basic</ScriptVersion>
  </ExtractionInfo>
  
  <BasicProperties>
    <PageId>$($Properties.PageId)</PageId>
    <PageName>$(Escape-Xml $Properties.PageName)</PageName>
    <PageTitle>$(Escape-Xml $Properties.PageTitle)</PageTitle>
    <ServerRelativeUrl>$(Escape-Xml $Properties.ServerRelativeUrl)</ServerRelativeUrl>
    <FileSize>$($Properties.FileSize)</FileSize>
    <Version>$(Escape-Xml $Properties.Version)</Version>
    <CheckOutType>$(Escape-Xml $Properties.CheckOutType)</CheckOutType>
    <Created>$($Properties.Created.ToString('yyyy-MM-ddTHH:mm:ss'))</Created>
    <Modified>$($Properties.Modified.ToString('yyyy-MM-ddTHH:mm:ss'))</Modified>
    <ContentType>$(Escape-Xml $Properties.ContentType)</ContentType>
    <ContentTypeId>$(Escape-Xml $Properties.ContentTypeId)</ContentTypeId>
  </BasicProperties>
  
  <CreatedBy>
    <Id>$($Properties.CreatedBy.Id)</Id>
    <LoginName>$(Escape-Xml $Properties.CreatedBy.LoginName)</LoginName>
    <Name>$(Escape-Xml $Properties.CreatedBy.Name)</Name>
    <Email>$(Escape-Xml $Properties.CreatedBy.Email)</Email>
  </CreatedBy>
  
  <ModifiedBy>
    <Id>$($Properties.ModifiedBy.Id)</Id>
    <LoginName>$(Escape-Xml $Properties.ModifiedBy.LoginName)</LoginName>
    <Name>$(Escape-Xml $Properties.ModifiedBy.Name)</Name>
    <Email>$(Escape-Xml $Properties.ModifiedBy.Email)</Email>
  </ModifiedBy>
  
  <AllFields>
"@

    # Add all field values
    foreach ($fieldName in $Properties.AllFields.Keys) {
        $fieldValue = Escape-Xml $Properties.AllFields[$fieldName]
        $xml += @"

    <Field name="$(Escape-Xml $fieldName)">$fieldValue</Field>
"@
    }

    $xml += @"

  </AllFields>
</SharePointPageProperties>
"@

    return $xml
}

#endregion

#region Main Execution

try {
    Write-Log "SharePoint 2016 Page Properties Extractor - Step 1" "Info"
    Write-Log "=================================================" "Info"
    
    # Step 1: Load CSOM assemblies
    Load-CSOMAssemblies
    
    # Step 2: Connect to SharePoint
    $context = Connect-SharePoint -SiteUrl $SiteUrl -UseInteractive $UseInteractiveAuth.IsPresent -Creds $Credentials
    
    # Step 3: Get page properties
    $properties = Get-PageProperties -Context $context -PageUrl $PageUrl
    
    # Step 4: Convert to XML
    $xmlContent = Convert-ToXml -Properties $properties
    
    # Step 5: Validate XML
    try {
        $xmlDoc = [xml]$xmlContent
        Write-Log "XML validation successful" "Success"
    }
    catch {
        Write-Log "XML validation failed: $($_.Exception.Message)" "Error"
        throw "Generated XML is not well-formed"
    }
    
    # Step 6: Save to file
    $outputDir = Split-Path $OutputPath -Parent
    if ($outputDir -and !(Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        Write-Log "Created output directory: $outputDir" "Info"
    }
    
    $xmlContent | Out-File -FilePath $OutputPath -Encoding UTF8
    Write-Log "XML saved to: $OutputPath" "Success"
    
    # Step 7: Display summary
    Write-Log "=================================================" "Info"
    Write-Log "EXTRACTION SUMMARY:" "Info"
    Write-Log "Page Name: $($properties.PageName)" "Info"
    Write-Log "Page Title: $($properties.PageTitle)" "Info"
    Write-Log "Content Type: $($properties.ContentType)" "Info"
    Write-Log "File Size: $($properties.FileSize) bytes" "Info"
    Write-Log "Fields Extracted: $($properties.AllFields.Count)" "Info"
    Write-Log "Output File: $OutputPath" "Info"
    Write-Log "=================================================" "Info"
    Write-Log "Step 1 completed successfully!" "Success"
    
}
catch {
    Write-Log "Step 1 failed: $($_.Exception.Message)" "Error"
    
    # Provide helpful troubleshooting
    if ($_.Exception.Message -like "*401*" -or $_.Exception.Message -like "*Unauthorized*") {
        Write-Log "" "Error"
        Write-Log "AUTHENTICATION HELP:" "Error"
        Write-Log "- For MFA: Use -UseInteractiveAuth" "Error"
        Write-Log "- For credentials: Use -Credentials parameter" "Error"
        Write-Log "- Check permissions to the site and page" "Error"
    }
    elseif ($_.Exception.Message -like "*CSOM*") {
        Write-Log "" "Error"
        Write-Log "CSOM HELP:" "Error"
        Write-Log "- Install SharePoint Client Components" "Error"
        Write-Log "- Download: https://www.microsoft.com/en-us/download/details.aspx?id=35585" "Error"
    }
    elseif ($_.Exception.Message -like "*File Not Found*") {
        Write-Log "" "Error"
        Write-Log "PAGE NOT FOUND HELP:" "Error"
        Write-Log "- Check page URL format: /Pages/Home.aspx" "Error"
        Write-Log "- Verify page exists in browser first" "Error"
        Write-Log "- Ensure you have read access to the page" "Error"
    }
    
    exit 1
}
finally {
    if ($context) {
        $context.Dispose()
        Write-Log "SharePoint context disposed" "Info"
    }
}

#endregion