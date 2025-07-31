using Microsoft.Extensions.Logging;
using SharePointPageMetadata.Models;
using System.Xml;

namespace SharePointPageMetadata.Services;

public class XmlMetadataGenerator
{
    private readonly ILogger _logger;

    public XmlMetadataGenerator(ILogger logger)
    {
        _logger = logger;
    }

    public async Task<string> GenerateXmlAsync(Models.SharePointPageMetadata metadata, string outputFolder)
    {
        _logger.LogInformation("Generating XML metadata output...");

        try
        {
            var xmlDoc = new XmlDocument();
            
            // Create XML declaration
            var xmlDeclaration = xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", null);
            xmlDoc.AppendChild(xmlDeclaration);

            // Root element
            var rootElement = xmlDoc.CreateElement("SharePointPageMetadata");
            xmlDoc.AppendChild(rootElement);

            // Add site information
            AddSiteInformation(xmlDoc, rootElement, metadata.SiteInformation);

            // Add page information
            AddPageInformation(xmlDoc, rootElement, metadata.PageInformation);

            // Generate output filename with timestamp
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var outputFileName = $"SharePoint_Page_Metadata_{timestamp}.xml";
            var outputPath = Path.Combine(outputFolder, outputFileName);

            // Save XML to file
            await Task.Run(() => xmlDoc.Save(outputPath));

            _logger.LogInformation("XML metadata saved to: {OutputPath}", outputPath);
            return outputPath;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to generate XML metadata");
            throw;
        }
    }

    private void AddSiteInformation(XmlDocument xmlDoc, XmlElement rootElement, SiteInformation? siteInfo)
    {
        if (siteInfo == null) return;

        var siteElement = xmlDoc.CreateElement("SiteInformation");
        rootElement.AppendChild(siteElement);

        AddElementIfNotNull(xmlDoc, siteElement, "SiteUrl", siteInfo.SiteUrl);
        AddElementIfNotNull(xmlDoc, siteElement, "WebTitle", siteInfo.WebTitle);
        AddElementIfNotNull(xmlDoc, siteElement, "WebDescription", siteInfo.WebDescription);
        AddElementIfNotNull(xmlDoc, siteElement, "WebId", siteInfo.WebId);
        AddElementIfNotNull(xmlDoc, siteElement, "WebServerRelativeUrl", siteInfo.WebServerRelativeUrl);
        
        if (siteInfo.WebCreated.HasValue)
            AddElementIfNotNull(xmlDoc, siteElement, "WebCreated", siteInfo.WebCreated.Value.ToString("yyyy-MM-ddTHH:mm:ssZ"));
        
        if (siteInfo.WebLastModified.HasValue)
            AddElementIfNotNull(xmlDoc, siteElement, "WebLastModified", siteInfo.WebLastModified.Value.ToString("yyyy-MM-ddTHH:mm:ssZ"));
    }

    private void AddPageInformation(XmlDocument xmlDoc, XmlElement rootElement, PageInformation? pageInfo)
    {
        if (pageInfo == null) return;

        var pageElement = xmlDoc.CreateElement("PageInformation");
        rootElement.AppendChild(pageElement);

        // Basic page information
        AddElementIfNotNull(xmlDoc, pageElement, "PageUrl", pageInfo.PageUrl);
        AddElementIfNotNull(xmlDoc, pageElement, "PageName", pageInfo.PageName);
        AddElementIfNotNull(xmlDoc, pageElement, "PageTitle", pageInfo.PageTitle);
        
        if (pageInfo.PageId > 0)
            AddElementIfNotNull(xmlDoc, pageElement, "PageId", pageInfo.PageId.ToString());
        
        AddElementIfNotNull(xmlDoc, pageElement, "PageUniqueId", pageInfo.PageUniqueId);
        AddElementIfNotNull(xmlDoc, pageElement, "PageServerRelativeUrl", pageInfo.PageServerRelativeUrl);
        
        if (pageInfo.PageFileSize > 0)
            AddElementIfNotNull(xmlDoc, pageElement, "PageFileSize", pageInfo.PageFileSize.ToString());
        
        AddElementIfNotNull(xmlDoc, pageElement, "PageCheckOutType", pageInfo.PageCheckOutType);
        AddElementIfNotNull(xmlDoc, pageElement, "PageLevel", pageInfo.PageLevel);

        // Add dates
        AddDates(xmlDoc, pageElement, pageInfo.Dates);

        // Add authors
        AddAuthors(xmlDoc, pageElement, pageInfo.Authors);

        // Add content type
        AddContentType(xmlDoc, pageElement, pageInfo.ContentType);

        // Add publishing information
        AddPublishingInformation(xmlDoc, pageElement, pageInfo.PublishingInformation);

        // Add custom fields
        AddCustomFields(xmlDoc, pageElement, pageInfo.CustomFields);

        // Add version information
        AddVersionInformation(xmlDoc, pageElement, pageInfo.VersionInformation);
    }

    private void AddDates(XmlDocument xmlDoc, XmlElement pageElement, DateInformation? dates)
    {
        if (dates == null) return;

        var datesElement = xmlDoc.CreateElement("Dates");
        pageElement.AppendChild(datesElement);

        if (dates.Created.HasValue)
            AddElementIfNotNull(xmlDoc, datesElement, "Created", dates.Created.Value.ToString("yyyy-MM-ddTHH:mm:ssZ"));

        if (dates.Modified.HasValue)
            AddElementIfNotNull(xmlDoc, datesElement, "Modified", dates.Modified.Value.ToString("yyyy-MM-ddTHH:mm:ssZ"));

        if (dates.FileCreated.HasValue)
            AddElementIfNotNull(xmlDoc, datesElement, "FileCreated", dates.FileCreated.Value.ToString("yyyy-MM-ddTHH:mm:ssZ"));

        if (dates.FileModified.HasValue)
            AddElementIfNotNull(xmlDoc, datesElement, "FileModified", dates.FileModified.Value.ToString("yyyy-MM-ddTHH:mm:ssZ"));
    }

    private void AddAuthors(XmlDocument xmlDoc, XmlElement pageElement, AuthorInformation? authors)
    {
        if (authors == null) return;

        var authorsElement = xmlDoc.CreateElement("Authors");
        pageElement.AppendChild(authorsElement);

        if (authors.CreatedBy != null)
        {
            var createdByElement = xmlDoc.CreateElement("CreatedBy");
            authorsElement.AppendChild(createdByElement);
            
            AddElementIfNotNull(xmlDoc, createdByElement, "LoginName", authors.CreatedBy.LoginName);
            AddElementIfNotNull(xmlDoc, createdByElement, "Email", authors.CreatedBy.Email);
        }

        if (authors.ModifiedBy != null)
        {
            var modifiedByElement = xmlDoc.CreateElement("ModifiedBy");
            authorsElement.AppendChild(modifiedByElement);
            
            AddElementIfNotNull(xmlDoc, modifiedByElement, "LoginName", authors.ModifiedBy.LoginName);
            AddElementIfNotNull(xmlDoc, modifiedByElement, "Email", authors.ModifiedBy.Email);
        }
    }

    private void AddContentType(XmlDocument xmlDoc, XmlElement pageElement, ContentTypeInformation? contentType)
    {
        if (contentType == null) return;

        var contentTypeElement = xmlDoc.CreateElement("ContentType");
        pageElement.AppendChild(contentTypeElement);

        AddElementIfNotNull(xmlDoc, contentTypeElement, "Name", contentType.Name);
        AddElementIfNotNull(xmlDoc, contentTypeElement, "Id", contentType.Id);
    }

    private void AddPublishingInformation(XmlDocument xmlDoc, XmlElement pageElement, Models.PublishingInformation? publishingInfo)
    {
        var publishingElement = xmlDoc.CreateElement("PublishingInformation");
        pageElement.AppendChild(publishingElement);

        if (publishingInfo == null) return;

        AddElementIfNotNull(xmlDoc, publishingElement, "PublishingPageLayout", publishingInfo.PublishingPageLayout);
        AddElementIfNotNull(xmlDoc, publishingElement, "PublishingPageContent", publishingInfo.PublishingPageContent);
        AddElementIfNotNull(xmlDoc, publishingElement, "PublishingContact", publishingInfo.PublishingContact);
        AddElementIfNotNull(xmlDoc, publishingElement, "PublishingPageImage", publishingInfo.PublishingPageImage);
        AddElementIfNotNull(xmlDoc, publishingElement, "PublishingRollupImage", publishingInfo.PublishingRollupImage);

        if (publishingInfo.PublishingStartDate.HasValue)
            AddElementIfNotNull(xmlDoc, publishingElement, "PublishingStartDate", publishingInfo.PublishingStartDate.Value.ToString("yyyy-MM-ddTHH:mm:ssZ"));

        if (publishingInfo.PublishingExpirationDate.HasValue)
            AddElementIfNotNull(xmlDoc, publishingElement, "PublishingExpirationDate", publishingInfo.PublishingExpirationDate.Value.ToString("yyyy-MM-ddTHH:mm:ssZ"));

        if (publishingInfo.PublishingHidden.HasValue)
            AddElementIfNotNull(xmlDoc, publishingElement, "PublishingHidden", publishingInfo.PublishingHidden.Value.ToString().ToLower());
    }

    private void AddCustomFields(XmlDocument xmlDoc, XmlElement pageElement, Dictionary<string, string>? customFields)
    {
        var customFieldsElement = xmlDoc.CreateElement("CustomFields");
        pageElement.AppendChild(customFieldsElement);

        if (customFields == null || customFields.Count == 0) return;

        foreach (var field in customFields)
        {
            try
            {
                var fieldElement = xmlDoc.CreateElement("Field");
                customFieldsElement.AppendChild(fieldElement);
                fieldElement.SetAttribute("Name", field.Key);
                fieldElement.InnerText = field.Value;
            }
            catch (Exception ex)
            {
                _logger.LogWarning("Could not process custom field '{FieldName}': {Error}", field.Key, ex.Message);
            }
        }
    }

    private void AddVersionInformation(XmlDocument xmlDoc, XmlElement pageElement, VersionInformation? versionInfo)
    {
        var versionElement = xmlDoc.CreateElement("VersionInformation");
        pageElement.AppendChild(versionElement);

        if (versionInfo == null) return;

        AddElementIfNotNull(xmlDoc, versionElement, "UIVersion", versionInfo.UIVersion);
        
        if (versionInfo.VersionNumber.HasValue)
            AddElementIfNotNull(xmlDoc, versionElement, "VersionNumber", versionInfo.VersionNumber.Value.ToString());

        // Add version history
        if (versionInfo.VersionHistory != null && versionInfo.VersionHistory.Count > 0)
        {
            var versionHistoryElement = xmlDoc.CreateElement("VersionHistory");
            versionElement.AppendChild(versionHistoryElement);

            foreach (var version in versionInfo.VersionHistory)
            {
                try
                {
                    var versionItemElement = xmlDoc.CreateElement("Version");
                    versionHistoryElement.AppendChild(versionItemElement);

                    if (!string.IsNullOrEmpty(version.VersionLabel))
                        versionItemElement.SetAttribute("VersionLabel", version.VersionLabel);

                    if (version.Size > 0)
                        versionItemElement.SetAttribute("Size", version.Size.ToString());

                    if (version.Created.HasValue)
                        versionItemElement.SetAttribute("Created", version.Created.Value.ToString("yyyy-MM-ddTHH:mm:ssZ"));

                    if (!string.IsNullOrEmpty(version.CreatedBy))
                        versionItemElement.SetAttribute("CreatedBy", version.CreatedBy);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning("Could not process version: {Error}", ex.Message);
                }
            }
        }
    }

    private void AddElementIfNotNull(XmlDocument xmlDoc, XmlElement parent, string elementName, string? value)
    {
        if (!string.IsNullOrEmpty(value))
        {
            var element = xmlDoc.CreateElement(elementName);
            element.InnerText = value;
            parent.AppendChild(element);
        }
    }
}