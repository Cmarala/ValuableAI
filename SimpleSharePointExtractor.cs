using Microsoft.SharePoint.Client;
using System.Security;
using System.Xml;

class Program
{
    // Configuration - Update these for your environment
    private static readonly string SiteUrl = "https://consulting.global.deloitteonline.com/sites/Aflac/POC";
    private static readonly string PageUrl = "SitePages/stars.aspx";
    private static readonly string OutputFolder = @"C:\Temp\SharePointOutput";

    static void Main(string[] args)
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

            // Get credentials
            Console.WriteLine("Please provide your SharePoint credentials:");
            Console.Write("Username: ");
            var username = Console.ReadLine();
            Console.Write("Password: ");
            var password = ReadPassword();

            if (string.IsNullOrEmpty(username) || string.IsNullOrEmpty(password))
            {
                Console.WriteLine("Username and password are required!");
                return;
            }

            // Extract metadata
            var metadata = ExtractPageMetadata(username, password);

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
            if (ex.InnerException != null)
                Console.WriteLine($"Inner Exception: {ex.InnerException.Message}");
        }

        Console.WriteLine("\nPress any key to exit...");
        Console.ReadKey();
    }

    static SharePointPageMetadata ExtractPageMetadata(string username, string password)
    {
        Console.WriteLine("Connecting to SharePoint...");

        using var context = new ClientContext(SiteUrl);

        // Set up credentials
        var securePassword = new SecureString();
        foreach (char c in password)
            securePassword.AppendChar(c);
        securePassword.MakeReadOnly();

        // For SharePoint 2016, try different credential approaches
        try
        {
            // Try SharePoint Online credentials first
            context.Credentials = new SharePointOnlineCredentials(username, securePassword);
        }
        catch
        {
            try
            {
                // Try Network credentials for on-premises
                context.Credentials = new System.Net.NetworkCredential(username, securePassword);
            }
            catch
            {
                // Try default credentials
                context.Credentials = System.Net.CredentialCache.DefaultCredentials;
            }
        }

        var metadata = new SharePointPageMetadata();

        // Get site information
        Console.WriteLine("Getting site information...");
        var web = context.Web;
        context.Load(web, w => w.Title, w => w.Description, w => w.Id, w => w.ServerRelativeUrl, w => w.Created);
        context.ExecuteQuery();

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
        context.Load(pageFile, f => f.Name, f => f.ServerRelativeUrl, f => f.Length, f => f.TimeCreated, f => f.TimeLastModified);

        var pageListItem = pageFile.ListItemAllFields;
        context.Load(pageListItem);
        context.ExecuteQuery();

        metadata.PageName = pageFile.Name;
        metadata.PageServerRelativeUrl = pageFile.ServerRelativeUrl;
        metadata.PageFileSize = pageFile.Length;
        metadata.PageCreated = pageFile.TimeCreated;
        metadata.PageModified = pageFile.TimeLastModified;

        // Get page metadata from list item
        var fieldValues = pageListItem.FieldValues;

        if (fieldValues.ContainsKey("Title"))
            metadata.PageTitle = fieldValues["Title"]?.ToString();

        if (fieldValues.ContainsKey("ID"))
            metadata.PageId = Convert.ToInt32(fieldValues["ID"]);

        if (fieldValues.ContainsKey("Created"))
            metadata.ListItemCreated = Convert.ToDateTime(fieldValues["Created"]);

        if (fieldValues.ContainsKey("Modified"))
            metadata.ListItemModified = Convert.ToDateTime(fieldValues["Modified"]);

        // Get author information
        if (fieldValues.ContainsKey("Author"))
        {
            var author = fieldValues["Author"] as FieldUserValue;
            metadata.CreatedBy = author?.LookupValue;
        }

        if (fieldValues.ContainsKey("Editor"))
        {
            var editor = fieldValues["Editor"] as FieldUserValue;
            metadata.ModifiedBy = editor?.LookupValue;
        }

        // Get content type
        if (fieldValues.ContainsKey("ContentType"))
            metadata.ContentType = fieldValues["ContentType"]?.ToString();

        // Get custom fields (exclude system fields)
        var systemFields = new HashSet<string> 
        { 
            "ID", "Title", "Created", "Modified", "Author", "Editor", "ContentType", 
            "UniqueId", "GUID", "_UIVersionString", "FileRef", "FileDirRef", "FileLeafRef",
            "_Level", "_IsCurrentVersion", "ItemChildCount", "FolderChildCount",
            "Attachments", "_ModerationStatus", "File_x0020_Type", "HTML_x0020_File_x0020_Type",
            "Edit", "LinkTitleNoMenu", "LinkTitle", "DocIcon", "FileSizeDisplay", 
            "ServerUrl", "EncodedAbsUrl", "BaseName", "MetaInfo", "SelectTitle"
        };

        metadata.CustomFields = new Dictionary<string, string>();

        foreach (var field in fieldValues)
        {
            if (!systemFields.Contains(field.Key) && field.Value != null && !string.IsNullOrEmpty(field.Value.ToString()))
            {
                try
                {
                    metadata.CustomFields[field.Key] = field.Value.ToString()!;
                }
                catch
                {
                    // Skip fields that can't be converted to string
                }
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
        AddElement(xmlDoc, siteInfo, "WebTitle", metadata.SiteTitle);
        AddElement(xmlDoc, siteInfo, "WebDescription", metadata.SiteDescription);
        AddElement(xmlDoc, siteInfo, "WebId", metadata.SiteId);
        AddElement(xmlDoc, siteInfo, "WebServerRelativeUrl", metadata.SiteServerRelativeUrl);
        if (metadata.SiteCreated.HasValue)
            AddElement(xmlDoc, siteInfo, "WebCreated", metadata.SiteCreated.Value.ToString("yyyy-MM-ddTHH:mm:ssZ"));

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