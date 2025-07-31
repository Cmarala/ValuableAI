using Microsoft.SharePoint.Client;
using PnP.Framework;
using System.Xml;

class Program
{
    // Configuration - Update these for your environment
    private static readonly string SiteUrl = "https://consulting.global.deloitteonline.com/sites/Aflac/POC";
    private static readonly string PageUrl = "SitePages/stars.aspx";
    private static readonly string OutputFolder = @"C:\Temp\SharePointOutput";

    static async Task Main(string[] args)
    {
        Console.WriteLine("SharePoint 2016 Page Metadata Extractor (POC)");
        Console.WriteLine("==============================================");
        Console.WriteLine($"Site: {SiteUrl}");
        Console.WriteLine($"Page: {PageUrl}");
        Console.WriteLine();

        try
        {
            // Create output directory
            Directory.CreateDirectory(OutputFolder);

            // Authenticate and extract metadata
            var metadata = await ExtractPageMetadata();

            // Generate XML output
            var outputPath = GenerateXmlOutput(metadata);

            Console.WriteLine("SUCCESS!");
            Console.WriteLine($"XML file saved: {outputPath}");
            Console.WriteLine();
            Console.WriteLine("=== METADATA SUMMARY ===");
            DisplaySummary(metadata);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ERROR: {ex.Message}");
        }

        Console.WriteLine("\nPress any key to exit...");
        Console.ReadKey();
    }

    static async Task<SharePointPageMetadata> ExtractPageMetadata()
    {
        Console.WriteLine("Connecting to SharePoint...");
        Console.WriteLine("A browser window will open for authentication.");
        Console.WriteLine("Please complete the login process...");

        var authManager = new AuthenticationManager();
        
        // Use interactive authentication - this will open browser for login
        using var context = authManager.GetACSAppOnlyContext(SiteUrl, "", ""); // This won't work, need proper method
        
        // Actually, let's use a simpler approach for POC
        using var ctx = new ClientContext(SiteUrl);
        
        // For SharePoint 2016 with browser auth, we'll need to prompt user
        Console.WriteLine("For this POC, please provide your credentials:");
        Console.Write("Username: ");
        var username = Console.ReadLine();
        Console.Write("Password: ");
        var password = ReadPassword();
        
        if (!string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(password))
        {
            var securePassword = new System.Security.SecureString();
            foreach (char c in password)
                securePassword.AppendChar(c);
            securePassword.MakeReadOnly();
            
            ctx.Credentials = new SharePointOnlineCredentials(username, securePassword);
        }

        var metadata = new SharePointPageMetadata();

        // Get site information
        Console.WriteLine("Getting site information...");
        var web = ctx.Web;
        ctx.Load(web, w => w.Title, w => w.Description, w => w.Id, w => w.ServerRelativeUrl, w => w.Created);
        ctx.ExecuteQuery();

        metadata.SiteTitle = web.Title;
        metadata.SiteDescription = web.Description;
        metadata.SiteId = web.Id.ToString();
        metadata.SiteUrl = SiteUrl;
        metadata.SiteServerRelativeUrl = web.ServerRelativeUrl;
        metadata.SiteCreated = web.Created;

        // Get page information
        Console.WriteLine("Getting page information...");
        var pageServerRelativeUrl = web.ServerRelativeUrl.TrimEnd('/') + "/" + PageUrl.TrimStart('/');
        var pageFile = web.GetFileByServerRelativeUrl(pageServerRelativeUrl);
        ctx.Load(pageFile, f => f.Name, f => f.ServerRelativeUrl, f => f.Length, f => f.TimeCreated, f => f.TimeLastModified);

        var pageListItem = pageFile.ListItemAllFields;
        ctx.Load(pageListItem);
        ctx.ExecuteQuery();

        metadata.PageName = pageFile.Name;
        metadata.PageServerRelativeUrl = pageFile.ServerRelativeUrl;
        metadata.PageFileSize = pageFile.Length;
        metadata.PageCreated = pageFile.TimeCreated;
        metadata.PageModified = pageFile.TimeLastModified;

        // Get page metadata from list item
        if (pageListItem.FieldValues.ContainsKey("Title"))
            metadata.PageTitle = pageListItem.FieldValues["Title"]?.ToString();

        if (pageListItem.FieldValues.ContainsKey("ID"))
            metadata.PageId = Convert.ToInt32(pageListItem.FieldValues["ID"]);

        if (pageListItem.FieldValues.ContainsKey("Created"))
            metadata.ListItemCreated = Convert.ToDateTime(pageListItem.FieldValues["Created"]);

        if (pageListItem.FieldValues.ContainsKey("Modified"))
            metadata.ListItemModified = Convert.ToDateTime(pageListItem.FieldValues["Modified"]);

        // Get author information
        if (pageListItem.FieldValues.ContainsKey("Author"))
        {
            var author = pageListItem.FieldValues["Author"] as FieldUserValue;
            metadata.CreatedBy = author?.LookupValue;
        }

        if (pageListItem.FieldValues.ContainsKey("Editor"))
        {
            var editor = pageListItem.FieldValues["Editor"] as FieldUserValue;
            metadata.ModifiedBy = editor?.LookupValue;
        }

        // Get content type
        if (pageListItem.FieldValues.ContainsKey("ContentType"))
            metadata.ContentType = pageListItem.FieldValues["ContentType"]?.ToString();

        // Get custom fields (non-system fields)
        var systemFields = new HashSet<string> { "ID", "Title", "Created", "Modified", "Author", "Editor", "ContentType", "UniqueId", "GUID", "_UIVersionString", "FileRef", "FileDirRef", "FileLeafRef" };
        metadata.CustomFields = new Dictionary<string, string>();

        foreach (var field in pageListItem.FieldValues)
        {
            if (!systemFields.Contains(field.Key) && field.Value != null && !string.IsNullOrEmpty(field.Value.ToString()))
            {
                metadata.CustomFields[field.Key] = field.Value.ToString()!;
            }
        }

        Console.WriteLine("Metadata extraction completed!");
        return metadata;
    }

    static string GenerateXmlOutput(SharePointPageMetadata metadata)
    {
        var xmlDoc = new XmlDocument();
        var xmlDeclaration = xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", null);
        xmlDoc.AppendChild(xmlDeclaration);

        var rootElement = xmlDoc.CreateElement("SharePointPageMetadata");
        xmlDoc.AppendChild(rootElement);

        // Site Information
        var siteInfo = xmlDoc.CreateElement("SiteInformation");
        rootElement.AppendChild(siteInfo);
        AddElement(xmlDoc, siteInfo, "SiteUrl", metadata.SiteUrl);
        AddElement(xmlDoc, siteInfo, "SiteTitle", metadata.SiteTitle);
        AddElement(xmlDoc, siteInfo, "SiteDescription", metadata.SiteDescription);
        AddElement(xmlDoc, siteInfo, "SiteId", metadata.SiteId);
        AddElement(xmlDoc, siteInfo, "SiteServerRelativeUrl", metadata.SiteServerRelativeUrl);
        if (metadata.SiteCreated.HasValue)
            AddElement(xmlDoc, siteInfo, "SiteCreated", metadata.SiteCreated.Value.ToString("yyyy-MM-ddTHH:mm:ssZ"));

        // Page Information
        var pageInfo = xmlDoc.CreateElement("PageInformation");
        rootElement.AppendChild(pageInfo);
        AddElement(xmlDoc, pageInfo, "PageUrl", PageUrl);
        AddElement(xmlDoc, pageInfo, "PageName", metadata.PageName);
        AddElement(xmlDoc, pageInfo, "PageTitle", metadata.PageTitle);
        if (metadata.PageId > 0)
            AddElement(xmlDoc, pageInfo, "PageId", metadata.PageId.ToString());
        AddElement(xmlDoc, pageInfo, "PageServerRelativeUrl", metadata.PageServerRelativeUrl);
        if (metadata.PageFileSize > 0)
            AddElement(xmlDoc, pageInfo, "PageFileSize", metadata.PageFileSize.ToString());

        // Dates
        var dates = xmlDoc.CreateElement("Dates");
        pageInfo.AppendChild(dates);
        if (metadata.ListItemCreated.HasValue)
            AddElement(xmlDoc, dates, "Created", metadata.ListItemCreated.Value.ToString("yyyy-MM-ddTHH:mm:ssZ"));
        if (metadata.ListItemModified.HasValue)
            AddElement(xmlDoc, dates, "Modified", metadata.ListItemModified.Value.ToString("yyyy-MM-ddTHH:mm:ssZ"));
        if (metadata.PageCreated.HasValue)
            AddElement(xmlDoc, dates, "FileCreated", metadata.PageCreated.Value.ToString("yyyy-MM-ddTHH:mm:ssZ"));
        if (metadata.PageModified.HasValue)
            AddElement(xmlDoc, dates, "FileModified", metadata.PageModified.Value.ToString("yyyy-MM-ddTHH:mm:ssZ"));

        // Authors
        var authors = xmlDoc.CreateElement("Authors");
        pageInfo.AppendChild(authors);
        if (!string.IsNullOrEmpty(metadata.CreatedBy))
        {
            var createdBy = xmlDoc.CreateElement("CreatedBy");
            authors.AppendChild(createdBy);
            AddElement(xmlDoc, createdBy, "LoginName", metadata.CreatedBy);
        }
        if (!string.IsNullOrEmpty(metadata.ModifiedBy))
        {
            var modifiedBy = xmlDoc.CreateElement("ModifiedBy");
            authors.AppendChild(modifiedBy);
            AddElement(xmlDoc, modifiedBy, "LoginName", metadata.ModifiedBy);
        }

        // Content Type
        var contentType = xmlDoc.CreateElement("ContentType");
        pageInfo.AppendChild(contentType);
        AddElement(xmlDoc, contentType, "Name", metadata.ContentType);

        // Custom Fields
        var customFields = xmlDoc.CreateElement("CustomFields");
        pageInfo.AppendChild(customFields);
        foreach (var field in metadata.CustomFields)
        {
            var fieldElement = xmlDoc.CreateElement("Field");
            customFields.AppendChild(fieldElement);
            fieldElement.SetAttribute("Name", field.Key);
            fieldElement.InnerText = field.Value;
        }

        // Save to file
        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var outputFileName = $"SharePoint_Page_Metadata_{timestamp}.xml";
        var outputPath = Path.Combine(OutputFolder, outputFileName);
        xmlDoc.Save(outputPath);

        return outputPath;
    }

    static void AddElement(XmlDocument doc, XmlElement parent, string name, string? value)
    {
        if (!string.IsNullOrEmpty(value))
        {
            var element = doc.CreateElement(name);
            element.InnerText = value;
            parent.AppendChild(element);
        }
    }

    static void DisplaySummary(SharePointPageMetadata metadata)
    {
        Console.WriteLine($"Site Title: {metadata.SiteTitle}");
        Console.WriteLine($"Page Title: {metadata.PageTitle}");
        Console.WriteLine($"Page ID: {metadata.PageId}");
        Console.WriteLine($"File Size: {metadata.PageFileSize} bytes");
        Console.WriteLine($"Created By: {metadata.CreatedBy}");
        Console.WriteLine($"Modified By: {metadata.ModifiedBy}");
        Console.WriteLine($"Content Type: {metadata.ContentType}");
        Console.WriteLine($"Custom Fields: {metadata.CustomFields.Count}");
        if (metadata.ListItemCreated.HasValue)
            Console.WriteLine($"Created: {metadata.ListItemCreated.Value}");
        if (metadata.ListItemModified.HasValue)
            Console.WriteLine($"Modified: {metadata.ListItemModified.Value}");
    }

    static string ReadPassword()
    {
        var password = "";
        ConsoleKeyInfo key;
        do
        {
            key = Console.ReadKey(true);
            if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter)
            {
                password += key.KeyChar;
                Console.Write("*");
            }
            else if (key.Key == ConsoleKey.Backspace && password.Length > 0)
            {
                password = password[0..^1];
                Console.Write("\b \b");
            }
        }
        while (key.Key != ConsoleKey.Enter);
        Console.WriteLine();
        return password;
    }
}

// Simple data model for POC
public class SharePointPageMetadata
{
    public string? SiteUrl { get; set; }
    public string? SiteTitle { get; set; }
    public string? SiteDescription { get; set; }
    public string? SiteId { get; set; }
    public string? SiteServerRelativeUrl { get; set; }
    public DateTime? SiteCreated { get; set; }

    public string? PageName { get; set; }
    public string? PageTitle { get; set; }
    public int PageId { get; set; }
    public string? PageServerRelativeUrl { get; set; }
    public long PageFileSize { get; set; }
    public DateTime? PageCreated { get; set; }
    public DateTime? PageModified { get; set; }

    public DateTime? ListItemCreated { get; set; }
    public DateTime? ListItemModified { get; set; }

    public string? CreatedBy { get; set; }
    public string? ModifiedBy { get; set; }
    public string? ContentType { get; set; }

    public Dictionary<string, string> CustomFields { get; set; } = new();
}