namespace SharePointPageMetadata.Models;

public class SharePointPageMetadata
{
    public SiteInformation? SiteInformation { get; set; }
    public PageInformation? PageInformation { get; set; }
}

public class SiteInformation
{
    public string? SiteUrl { get; set; }
    public string? WebTitle { get; set; }
    public string? WebDescription { get; set; }
    public string? WebId { get; set; }
    public string? WebServerRelativeUrl { get; set; }
    public DateTime? WebCreated { get; set; }
    public DateTime? WebLastModified { get; set; }
}

public class PageInformation
{
    public string? PageUrl { get; set; }
    public string? PageName { get; set; }
    public string? PageTitle { get; set; }
    public int PageId { get; set; }
    public string? PageUniqueId { get; set; }
    public string? PageServerRelativeUrl { get; set; }
    public long PageFileSize { get; set; }
    public string? PageCheckOutType { get; set; }
    public string? PageLevel { get; set; }
    public DateInformation? Dates { get; set; }
    public AuthorInformation? Authors { get; set; }
    public ContentTypeInformation? ContentType { get; set; }
    public PublishingInformation? PublishingInformation { get; set; }
    public Dictionary<string, string>? CustomFields { get; set; }
    public VersionInformation? VersionInformation { get; set; }
}

public class DateInformation
{
    public DateTime? Created { get; set; }
    public DateTime? Modified { get; set; }
    public DateTime? FileCreated { get; set; }
    public DateTime? FileModified { get; set; }
}

public class AuthorInformation
{
    public UserInfo? CreatedBy { get; set; }
    public UserInfo? ModifiedBy { get; set; }
}

public class UserInfo
{
    public string? LoginName { get; set; }
    public string? Email { get; set; }
}

public class ContentTypeInformation
{
    public string? Name { get; set; }
    public string? Id { get; set; }
}

public class PublishingInformation
{
    public string? PublishingPageLayout { get; set; }
    public string? PublishingPageContent { get; set; }
    public DateTime? PublishingStartDate { get; set; }
    public DateTime? PublishingExpirationDate { get; set; }
    public string? PublishingContact { get; set; }
    public string? PublishingPageImage { get; set; }
    public string? PublishingRollupImage { get; set; }
    public bool? PublishingHidden { get; set; }
}

public class VersionInformation
{
    public string? UIVersion { get; set; }
    public int? VersionNumber { get; set; }
    public List<VersionHistoryItem>? VersionHistory { get; set; }
}

public class VersionHistoryItem
{
    public string? VersionLabel { get; set; }
    public long Size { get; set; }
    public DateTime? Created { get; set; }
    public string? CreatedBy { get; set; }
}