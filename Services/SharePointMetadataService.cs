using Microsoft.SharePoint.Client;
using Microsoft.Extensions.Logging;
using SharePointPageMetadata.Models;
using System.Security;

namespace SharePointPageMetadata.Services;

public class SharePointMetadataService
{
    private readonly ILogger _logger;
    private readonly HashSet<string> _systemFields;

    public SharePointMetadataService(ILogger logger)
    {
        _logger = logger;
        _systemFields = new HashSet<string>
        {
            "ID", "Title", "Created", "Modified", "Author", "Editor", "ContentType", "ContentTypeId",
            "UniqueId", "GUID", "_UIVersionString", "FileRef", "FileDirRef", "FileLeafRef",
            "_Level", "_IsCurrentVersion", "ItemChildCount", "FolderChildCount", "_HasCopyDestinations",
            "_CopySource", "owshiddenversion", "WorkflowVersion", "_UIVersion", "Attachments",
            "_ModerationStatus", "_ModerationComments", "File_x0020_Type", "HTML_x0020_File_x0020_Type",
            "Edit", "LinkTitleNoMenu", "LinkTitle", "DocIcon", "FileSizeDisplay", "ServerUrl",
            "EncodedAbsUrl", "BaseName", "MetaInfo", "_EditMenuTableStart", "_EditMenuTableEnd",
            "LinkFilenameNoMenu", "LinkFilename", "SelectTitle", "SelectFilename", "IndentLevel",
            "PermMask", "CheckedOutUserId", "IsCheckedoutToLocal", "CheckedOutTitle", "CheckoutUser",
            "ScopeId", "VirusStatus", "_CheckinComment", "LinkCheckedOutTitle", "Modified_x0020_By",
            "Created_x0020_By", "File_x0020_Size", "InstanceID", "Order", "WorkflowInstanceID"
        };
    }

    public async Task<Models.SharePointPageMetadata> ExtractPageMetadataAsync(
        string siteUrl, string pageUrl, string? username, string? password, bool useModernAuth)
    {
        _logger.LogInformation("Connecting to SharePoint site: {SiteUrl}", siteUrl);

        using var context = new ClientContext(siteUrl);

        // Configure authentication
        if (useModernAuth)
        {
            _logger.LogInformation("Using modern authentication (interactive browser login)");
            // For modern auth, you would typically use Microsoft.Graph or other modern auth libraries
            // For this example, we'll use basic auth as a fallback
            _logger.LogWarning("Modern auth not fully implemented in this sample. Using credentials if provided.");
        }

        if (!string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(password))
        {
            _logger.LogInformation("Using credential-based authentication for user: {Username}", username);
            var securePassword = new SecureString();
            foreach (char c in password)
                securePassword.AppendChar(c);
            securePassword.MakeReadOnly();

            context.Credentials = new SharePointOnlineCredentials(username, securePassword);
        }

        try
        {
            var metadata = new Models.SharePointPageMetadata();

            // Get site information
            _logger.LogInformation("Retrieving web information...");
            var web = context.Web;
            context.Load(web, w => w.Title, w => w.Description, w => w.Id, w => w.ServerRelativeUrl, w => w.Created, w => w.LastItemModifiedDate);
            await context.ExecuteQueryAsync();

            metadata.SiteInformation = new SiteInformation
            {
                SiteUrl = siteUrl,
                WebTitle = web.Title,
                WebDescription = web.Description,
                WebId = web.Id.ToString(),
                WebServerRelativeUrl = web.ServerRelativeUrl,
                WebCreated = web.Created,
                WebLastModified = web.LastItemModifiedDate
            };

            _logger.LogInformation("Web retrieved successfully: {WebTitle}", web.Title);

            // Get page file information
            _logger.LogInformation("Retrieving page file: {PageUrl}", pageUrl);
            var pageFile = web.GetFileByServerRelativeUrl(pageUrl);
            context.Load(pageFile, f => f.Name, f => f.ServerRelativeUrl, f => f.Length, f => f.CheckOutType, f => f.Level, f => f.TimeCreated, f => f.TimeLastModified);

            // Get page as list item
            _logger.LogInformation("Getting page as list item...");
            var pageListItem = pageFile.ListItemAllFields;
            context.Load(pageListItem);

            await context.ExecuteQueryAsync();

            _logger.LogInformation("Page file retrieved successfully: {PageName}", pageFile.Name);

            // Get the Site Pages list
            List? sitePagesList = null;
            try
            {
                _logger.LogInformation("Getting Site Pages list...");
                sitePagesList = web.Lists.GetByTitle("Site Pages");
                context.Load(sitePagesList, l => l.Title, l => l.Id);
                await context.ExecuteQueryAsync();
                _logger.LogInformation("Site Pages list retrieved successfully: {ListTitle}", sitePagesList.Title);
            }
            catch (Exception ex)
            {
                _logger.LogWarning("Could not retrieve Site Pages list: {Error}", ex.Message);
                try
                {
                    sitePagesList = web.Lists.GetByTitle("Pages");
                    context.Load(sitePagesList, l => l.Title, l => l.Id);
                    await context.ExecuteQueryAsync();
                    _logger.LogInformation("Pages list retrieved successfully: {ListTitle}", sitePagesList.Title);
                }
                catch
                {
                    _logger.LogWarning("Could not find Site Pages or Pages list");
                }
            }

            // Build page information
            metadata.PageInformation = await BuildPageInformationAsync(context, pageFile, pageListItem, pageUrl);

            // Get version history
            await GetVersionHistoryAsync(context, pageFile, metadata.PageInformation);

            _logger.LogInformation("Metadata extraction completed successfully");
            return metadata;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to extract SharePoint page metadata");
            throw;
        }
    }

    private async Task<PageInformation> BuildPageInformationAsync(ClientContext context, Microsoft.SharePoint.Client.File pageFile, ListItem pageListItem, string pageUrl)
    {
        var pageInfo = new PageInformation
        {
            PageUrl = pageUrl,
            PageName = pageFile.Name,
            PageServerRelativeUrl = pageFile.ServerRelativeUrl,
            PageFileSize = pageFile.Length,
            PageCheckOutType = pageFile.CheckOutType.ToString(),
            PageLevel = pageFile.Level.ToString()
        };

        // Get additional list item properties
        if (pageListItem != null)
        {
            try
            {
                // Reload with all field values
                context.Load(pageListItem, item => item.FieldValues);
                await context.ExecuteQueryAsync();

                var fieldValues = pageListItem.FieldValues;

                // Basic properties
                if (fieldValues.ContainsKey("Title") && fieldValues["Title"] != null)
                    pageInfo.PageTitle = fieldValues["Title"].ToString();

                if (fieldValues.ContainsKey("ID") && fieldValues["ID"] != null)
                    pageInfo.PageId = Convert.ToInt32(fieldValues["ID"]);

                if (fieldValues.ContainsKey("UniqueId") && fieldValues["UniqueId"] != null)
                    pageInfo.PageUniqueId = fieldValues["UniqueId"].ToString();

                // Dates
                pageInfo.Dates = new DateInformation
                {
                    FileCreated = pageFile.TimeCreated,
                    FileModified = pageFile.TimeLastModified
                };

                if (fieldValues.ContainsKey("Created") && fieldValues["Created"] != null)
                    pageInfo.Dates.Created = Convert.ToDateTime(fieldValues["Created"]);

                if (fieldValues.ContainsKey("Modified") && fieldValues["Modified"] != null)
                    pageInfo.Dates.Modified = Convert.ToDateTime(fieldValues["Modified"]);

                // Authors
                pageInfo.Authors = new AuthorInformation();

                if (fieldValues.ContainsKey("Author") && fieldValues["Author"] != null)
                {
                    var authorValue = fieldValues["Author"] as FieldUserValue;
                    if (authorValue != null)
                    {
                        pageInfo.Authors.CreatedBy = new UserInfo
                        {
                            LoginName = authorValue.LookupValue,
                            Email = authorValue.Email
                        };
                    }
                }

                if (fieldValues.ContainsKey("Editor") && fieldValues["Editor"] != null)
                {
                    var editorValue = fieldValues["Editor"] as FieldUserValue;
                    if (editorValue != null)
                    {
                        pageInfo.Authors.ModifiedBy = new UserInfo
                        {
                            LoginName = editorValue.LookupValue,
                            Email = editorValue.Email
                        };
                    }
                }

                // Content Type
                pageInfo.ContentType = new ContentTypeInformation();
                if (fieldValues.ContainsKey("ContentType") && fieldValues["ContentType"] != null)
                    pageInfo.ContentType.Name = fieldValues["ContentType"].ToString();

                if (fieldValues.ContainsKey("ContentTypeId") && fieldValues["ContentTypeId"] != null)
                    pageInfo.ContentType.Id = fieldValues["ContentTypeId"].ToString();

                // Publishing Information
                pageInfo.PublishingInformation = ExtractPublishingInformation(fieldValues);

                // Version Information
                pageInfo.VersionInformation = new VersionInformation();
                if (fieldValues.ContainsKey("_UIVersionString") && fieldValues["_UIVersionString"] != null)
                    pageInfo.VersionInformation.UIVersion = fieldValues["_UIVersionString"].ToString();

                if (fieldValues.ContainsKey("_UIVersion") && fieldValues["_UIVersion"] != null)
                    pageInfo.VersionInformation.VersionNumber = Convert.ToInt32(fieldValues["_UIVersion"]);

                // Custom Fields
                pageInfo.CustomFields = ExtractCustomFields(fieldValues);
            }
            catch (Exception ex)
            {
                _logger.LogWarning("Could not load all field values: {Error}", ex.Message);
            }
        }

        return pageInfo;
    }

    private Models.PublishingInformation ExtractPublishingInformation(IDictionary<string, object> fieldValues)
    {
        var publishingInfo = new Models.PublishingInformation();
        var publishingFields = new Dictionary<string, Action<string>>
        {
            ["PublishingPageLayout"] = value => publishingInfo.PublishingPageLayout = value,
            ["PublishingPageContent"] = value => publishingInfo.PublishingPageContent = value,
            ["PublishingContact"] = value => publishingInfo.PublishingContact = value,
            ["PublishingPageImage"] = value => publishingInfo.PublishingPageImage = value,
            ["PublishingRollupImage"] = value => publishingInfo.PublishingRollupImage = value,
        };

        foreach (var field in publishingFields)
        {
            if (fieldValues.ContainsKey(field.Key) && fieldValues[field.Key] != null)
                field.Value(fieldValues[field.Key].ToString()!);
        }

        // Date fields
        if (fieldValues.ContainsKey("PublishingStartDate") && fieldValues["PublishingStartDate"] != null)
            publishingInfo.PublishingStartDate = Convert.ToDateTime(fieldValues["PublishingStartDate"]);

        if (fieldValues.ContainsKey("PublishingExpirationDate") && fieldValues["PublishingExpirationDate"] != null)
            publishingInfo.PublishingExpirationDate = Convert.ToDateTime(fieldValues["PublishingExpirationDate"]);

        if (fieldValues.ContainsKey("PublishingHidden") && fieldValues["PublishingHidden"] != null)
            publishingInfo.PublishingHidden = Convert.ToBoolean(fieldValues["PublishingHidden"]);

        return publishingInfo;
    }

    private Dictionary<string, string> ExtractCustomFields(IDictionary<string, object> fieldValues)
    {
        var customFields = new Dictionary<string, string>();

        foreach (var field in fieldValues)
        {
            if (!_systemFields.Contains(field.Key) && field.Value != null && !string.IsNullOrEmpty(field.Value.ToString()))
            {
                try
                {
                    customFields[field.Key] = field.Value.ToString()!;
                }
                catch (Exception ex)
                {
                    _logger.LogWarning("Could not process custom field '{FieldName}': {Error}", field.Key, ex.Message);
                }
            }
        }

        return customFields;
    }

    private async Task GetVersionHistoryAsync(ClientContext context, Microsoft.SharePoint.Client.File pageFile, PageInformation pageInfo)
    {
        try
        {
            _logger.LogInformation("Retrieving version history...");
            var versions = pageFile.Versions;
            context.Load(versions, v => v.Include(ver => ver.VersionLabel, ver => ver.Size, ver => ver.Created, ver => ver.CreatedBy));
            await context.ExecuteQueryAsync();

            pageInfo.VersionInformation ??= new VersionInformation();
            pageInfo.VersionInformation.VersionHistory = new List<VersionHistoryItem>();

            foreach (var version in versions)
            {
                var versionItem = new VersionHistoryItem
                {
                    VersionLabel = version.VersionLabel,
                    Size = version.Size,
                    Created = version.Created,
                    CreatedBy = version.CreatedBy?.LookupValue
                };
                pageInfo.VersionInformation.VersionHistory.Add(versionItem);
            }

            _logger.LogInformation("Version history retrieved: {VersionCount} versions", versions.Count);
        }
        catch (Exception ex)
        {
            _logger.LogWarning("Could not retrieve version history: {Error}", ex.Message);
        }
    }
}