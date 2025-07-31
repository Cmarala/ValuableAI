using System.CommandLine;
using Microsoft.Extensions.Logging;
using SharePointPageMetadata.Services;
using SharePointPageMetadata.Models;

namespace SharePointPageMetadata;

class Program
{
    static async Task<int> Main(string[] args)
    {
        // Create command line options
        var siteUrlOption = new Option<string>(
            name: "--site-url",
            description: "The SharePoint site URL",
            getDefaultValue: () => "https://consulting.global.deloitteonline.com/sites/Aflac/POC");

        var pageUrlOption = new Option<string>(
            name: "--page-url", 
            description: "The relative URL of the page to extract metadata from",
            getDefaultValue: () => "SitePages/stars.aspx");

        var outputFolderOption = new Option<string>(
            name: "--output-folder",
            description: "The folder where the XML output will be saved",
            getDefaultValue: () => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "SharePointMetadataOutput"));

        var usernameOption = new Option<string?>(
            name: "--username",
            description: "Username for SharePoint authentication (optional for Modern Auth)");

        var passwordOption = new Option<string?>(
            name: "--password",
            description: "Password for SharePoint authentication (optional for Modern Auth)");

        var useModernAuthOption = new Option<bool>(
            name: "--use-modern-auth",
            description: "Use modern authentication (interactive browser login) for MFA support",
            getDefaultValue: () => true);

        // Create root command
        var rootCommand = new RootCommand("SharePoint 2016 Page Metadata Extractor")
        {
            siteUrlOption,
            pageUrlOption,
            outputFolderOption,
            usernameOption,
            passwordOption,
            useModernAuthOption
        };

        rootCommand.SetHandler(async (siteUrl, pageUrl, outputFolder, username, password, useModernAuth) =>
        {
            // Setup logging
            using var loggerFactory = LoggerFactory.Create(builder =>
            {
                builder.AddConsole().SetMinimumLevel(LogLevel.Information);
            });
            var logger = loggerFactory.CreateLogger<Program>();

            try
            {
                logger.LogInformation("SharePoint 2016 Page Metadata Extractor Starting...");
                logger.LogInformation("Site URL: {SiteUrl}", siteUrl);
                logger.LogInformation("Page URL: {PageUrl}", pageUrl);
                logger.LogInformation("Output Folder: {OutputFolder}", outputFolder);

                // Create output directory if it doesn't exist
                if (!Directory.Exists(outputFolder))
                {
                    Directory.CreateDirectory(outputFolder);
                    logger.LogInformation("Created output directory: {OutputFolder}", outputFolder);
                }

                // Create SharePoint service
                var sharePointService = new SharePointMetadataService(logger);

                // Extract metadata
                var metadata = await sharePointService.ExtractPageMetadataAsync(
                    siteUrl, pageUrl, username, password, useModernAuth);

                // Generate XML output
                var xmlGenerator = new XmlMetadataGenerator(logger);
                var outputPath = await xmlGenerator.GenerateXmlAsync(metadata, outputFolder);

                logger.LogInformation("Metadata extraction completed successfully!");
                logger.LogInformation("Output file: {OutputPath}", outputPath);

                // Display summary
                DisplayMetadataSummary(metadata, logger);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "An error occurred during metadata extraction");
                return;
            }
        }, siteUrlOption, pageUrlOption, outputFolderOption, usernameOption, passwordOption, useModernAuthOption);

        return await rootCommand.InvokeAsync(args);
    }

    private static void DisplayMetadataSummary(Models.SharePointPageMetadata metadata, ILogger logger)
    {
        logger.LogInformation("=== METADATA SUMMARY ===");
        
        if (!string.IsNullOrEmpty(metadata.PageInformation?.PageTitle))
            logger.LogInformation("Page Title: {PageTitle}", metadata.PageInformation.PageTitle);
        
        if (metadata.PageInformation?.PageId > 0)
            logger.LogInformation("Page ID: {PageId}", metadata.PageInformation.PageId);
        
        if (metadata.PageInformation?.Dates?.Created.HasValue == true)
            logger.LogInformation("Created: {Created}", metadata.PageInformation.Dates.Created.Value);
        
        if (metadata.PageInformation?.Dates?.Modified.HasValue == true)
            logger.LogInformation("Modified: {Modified}", metadata.PageInformation.Dates.Modified.Value);
        
        if (metadata.PageInformation?.PageFileSize > 0)
            logger.LogInformation("File Size: {FileSize} bytes", metadata.PageInformation.PageFileSize);
        
        if (metadata.PageInformation?.CustomFields?.Count > 0)
            logger.LogInformation("Custom Fields Found: {CustomFieldsCount}", metadata.PageInformation.CustomFields.Count);
        
        if (metadata.PageInformation?.VersionInformation?.VersionHistory?.Count > 0)
            logger.LogInformation("Version History: {VersionCount} versions", metadata.PageInformation.VersionInformation.VersionHistory.Count);
    }
}