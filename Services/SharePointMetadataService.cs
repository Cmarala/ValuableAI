using Microsoft.SharePoint.Client;
using Microsoft.Extensions.Logging;
using SharePointPageMetadata.Models;
using PnP.Framework;
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
        return await Task.Run(() => ExtractPageMetadata(siteUrl, pageUrl, username, password, useModernAuth));
    }

    private Models.SharePointPageMetadata ExtractPageMetadata(
        string siteUrl, string pageUrl, string? username, string? password, bool useModernAuth)
    {
        _logger.LogInformation("Connecting to SharePoint site: {SiteUrl}", siteUrl);

        ClientContext? context = null;
        
        try
        {
            // Create AuthenticationManager
            var authManager = new AuthenticationManager();
            
            if (useModernAuth || string.IsNullOrEmpty(username) || string.IsNullOrEmpty(password))
            {
                _logger.LogInformation("Modern authentication requested");
                
                // For modern auth with MFA, user needs to configure app registration
                // This is a simplified approach - in real scenarios you'd need proper app registration
                _logger.LogWarning("Modern authentication requires app registration. Using direct context creation for demo.");
                _logger.LogWarning("For MFA support, please register an Azure AD app and use GetACSAppOnlyContext with client credentials.");
                
                // Create basic context (this will need credentials or will fail)
                context = new ClientContext(siteUrl);
                
                if (!string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(password))
                {
                    var securePassword = new SecureString();
                    foreach (char c in password)
                        securePassword.AppendChar(c);
                    securePassword.MakeReadOnly();
                    
                    context.Credentials = new SharePointOnlineCredentials(username, securePassword);
                }
            }
            else
            {
                _logger.LogInformation("Using credential-based authentication for user: {Username}", username);
                
                // Use basic SharePointOnlineCredentials
                context = new ClientContext(siteUrl);
                var securePassword = new SecureString();
                foreach (char c in password)
                    securePassword.AppendChar(c);
                securePassword.MakeReadOnly();

                context.Credentials = new SharePointOnlineCredentials(username, securePassword);
            }

            if (context == null)
            {
                throw new InvalidOperationException("Failed to create SharePoint context. Authentication may have failed.");
            }

            _logger.LogInformation("SharePoint context created successfully");

            var metadata = new Models.SharePointPageMetadata();

            // Get site information
            _logger.LogInformation("Retrieving web information...");
            var web = context.Web;
            context.Load(web, w => w.Title, w => w.Description, w => w.Id, w => w.ServerRelativeUrl, w => w.Created, w => w.LastItemModifiedDate);
            context.ExecuteQuery();

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
            
            // Construct the full server relative URL for the page
            var pageServerRelativeUrl = pageUrl;
            if (!pageUrl.StartsWith("/"))
            {
                pageServerRelativeUrl = web.ServerRelativeUrl.TrimEnd('/') + "/" + pageUrl.TrimStart('/');
            }

            var pageFile = web.GetFileByServerRelativeUrl(pageServerRelativeUrl);
            context.Load(pageFile, f => f.Name, f => f.ServerRelativeUrl, f => f.Length, f => f.CheckOutType, f => f.Level, f => f.TimeCreated, f => f.TimeLastModified);

            // Get page as list item
            _logger.LogInformation("Getting page as list item...");
            var pageListItem = pageFile.ListItemAllFields;
            context.Load(pageListItem);

            context.ExecuteQuery();

            _logger.LogInformation("Page file retrieved successfully: {PageName}", pageFile.Name);

            // Get the Site Pages list (optional, for additional context)
            List? sitePagesList = null;
            try
            {
                _logger.LogInformation("Getting Site Pages list...");
                sitePagesList = web.Lists.GetByTitle("Site Pages");
                context.Load(sitePagesList, l => l.Title, l => l.Id);
                context.ExecuteQuery();
                _logger.LogInformation("Site Pages list retrieved successfully: {ListTitle}", sitePagesList.Title);
            }
            catch (Exception ex)
            {
                _logger.LogWarning("Could not retrieve Site Pages list: {Error}", ex.Message);
                try
                {
                    sitePagesList = web.Lists.GetByTitle("Pages");
                    context.Load(sitePagesList, l => l.Title, l => l.Id);
                    context.ExecuteQuery();
                    _logger.LogInformation("Pages list retrieved successfully: {ListTitle}", sitePagesList.Title);
                }
                catch
                {
                    _logger.LogWarning("Could not find Site Pages or Pages list");
                }
            }

            // Build page information
            metadata.PageInformation = BuildPageInformation(context, pageFile, pageListItem, pageUrl);

            // Get version history
            GetVersionHistory(context, pageFile, metadata.PageInformation);

            _logger.LogInformation("Metadata extraction completed successfully");
            return metadata;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to extract SharePoint page metadata");
            throw;
        }
        finally
        {
            // Clean up context
            context?.Dispose();
        }
    }

    private PageInformation BuildPageInformation(ClientContext context, Microsoft.SharePoint.Client.File pageFile, ListItem pageListItem, string pageUrl)
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
                context.Load(pageListItem, item => item.FieldValues, item => item["Author"], item => item["Editor"]);
                context.ExecuteQuery();

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

                // Handle Author field
                try
                {
                    var authorField = pageListItem["Author"];
                    if (authorField != null)
                    {
                        if (authorField is FieldUserValue authorUserValue)
                        {
                            pageInfo.Authors.CreatedBy = new UserInfo
                            {
                                LoginName = authorUserValue.LookupValue,
                                Email = authorUserValue.Email
                            };
                        }
                        else if (authorField is User authorUser)
                        {
                            context.Load(authorUser, u => u.Title, u => u.Email, u => u.LoginName);
                            context.ExecuteQuery();
                            pageInfo.Authors.CreatedBy = new UserInfo
                            {
                                LoginName = authorUser.Title,
                                Email = authorUser.Email
                            };
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning("Could not process Author field: {Error}", ex.Message);
                }

                // Handle Editor field
                try
                {
                    var editorField = pageListItem["Editor"];
                    if (editorField != null)
                    {
                        if (editorField is FieldUserValue editorUserValue)
                        {
                            pageInfo.Authors.ModifiedBy = new UserInfo
                            {
                                LoginName = editorUserValue.LookupValue,
                                Email = editorUserValue.Email
                            };
                        }
                        else if (editorField is User editorUser)
                        {
                            context.Load(editorUser, u => u.Title, u => u.Email, u => u.LoginName);
                            context.ExecuteQuery();
                            pageInfo.Authors.ModifiedBy = new UserInfo
                            {
                                LoginName = editorUser.Title,
                                Email = editorUser.Email
                            };
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning("Could not process Editor field: {Error}", ex.Message);
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
        {
            if (DateTime.TryParse(fieldValues["PublishingStartDate"].ToString(), out var startDate))
                publishingInfo.PublishingStartDate = startDate;
        }

        if (fieldValues.ContainsKey("PublishingExpirationDate") && fieldValues["PublishingExpirationDate"] != null)
        {
            if (DateTime.TryParse(fieldValues["PublishingExpirationDate"].ToString(), out var expDate))
                publishingInfo.PublishingExpirationDate = expDate;
        }

        if (fieldValues.ContainsKey("PublishingHidden") && fieldValues["PublishingHidden"] != null)
        {
            if (bool.TryParse(fieldValues["PublishingHidden"].ToString(), out var hidden))
                publishingInfo.PublishingHidden = hidden;
        }

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

    private void GetVersionHistory(ClientContext context, Microsoft.SharePoint.Client.File pageFile, PageInformation pageInfo)
    {
        try
        {
            _logger.LogInformation("Retrieving version history...");
            var versions = pageFile.Versions;
            context.Load(versions, v => v.Include(ver => ver.VersionLabel, ver => ver.Size, ver => ver.Created, ver => ver.CreatedBy));
            context.ExecuteQuery();

            pageInfo.VersionInformation ??= new VersionInformation();
            pageInfo.VersionInformation.VersionHistory = new List<VersionHistoryItem>();

            foreach (var version in versions)
            {
                try
                {
                    var versionItem = new VersionHistoryItem
                    {
                        VersionLabel = version.VersionLabel,
                        Size = version.Size,
                        Created = version.Created
                    };

                    // Try to get the created by user info
                    if (version.CreatedBy != null)
                    {
                        try
                        {
                            context.Load(version.CreatedBy, u => u.Title, u => u.Email);
                            context.ExecuteQuery();
                            versionItem.CreatedBy = version.CreatedBy.Title;
                        }
                        catch
                        {
                            // If we can't load the user, just use what we have
                            versionItem.CreatedBy = "Unknown";
                        }
                    }

                    pageInfo.VersionInformation.VersionHistory.Add(versionItem);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning("Could not process version {VersionLabel}: {Error}", version.VersionLabel, ex.Message);
                }
            }

            _logger.LogInformation("Version history retrieved: {VersionCount} versions", versions.Count);
        }
        catch (Exception ex)
        {
            _logger.LogWarning("Could not retrieve version history: {Error}", ex.Message);
        }
    }
}