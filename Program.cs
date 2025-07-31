using System;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using OfficeDevPnP.Core;

namespace SharePointWebPartReader
{
    class Program
    {
        static void Main(string[] args)
        {
            // SharePoint site URL and page URL
            string siteUrl = "https://your-sharepoint-site.com";  // Update with your SharePoint site URL
            string pageUrl = "/sites/dev/SitePages/TestPage.aspx";
            
            Console.WriteLine("SharePoint Web Part Reader");
            Console.WriteLine("=============================");
            Console.WriteLine();
            
            try
            {
                Console.WriteLine("Connecting to SharePoint with MFA-enabled WebLogin...");
                Console.WriteLine("A browser window will open for authentication.");
                Console.WriteLine();
                
                // Create ClientContext using PnP AuthenticationManager for WebLogin with MFA support
                // This will open a browser window for interactive login including MFA
                AuthenticationManager authManager = new AuthenticationManager();
                using (ClientContext context = authManager.GetWebLoginClientContext(siteUrl))
                {
                    // Load web and site information
                    Web web = context.Web;
                    context.Load(web, w => w.Title, w => w.Url);
                    context.ExecuteQuery();
                    
                    Console.WriteLine($"Connected to: {web.Title}");
                    Console.WriteLine($"Site URL: {web.Url}");
                    Console.WriteLine();
                    
                    // Get the file (page)
                    Microsoft.SharePoint.Client.File file = web.GetFileByServerRelativeUrl(pageUrl);
                    context.Load(file, f => f.Exists, f => f.Name);
                    context.ExecuteQuery();
                    
                    if (!file.Exists)
                    {
                        Console.WriteLine($"ERROR: Page not found: {pageUrl}");
                        return;
                    }
                    
                    Console.WriteLine($"Loading page: {file.Name}");
                    Console.WriteLine();
                    
                    // Get the LimitedWebPartManager
                    LimitedWebPartManager webPartManager = file.GetLimitedWebPartManager(PersonalizationScope.Shared);
                    
                    // Load web parts
                    var webParts = webPartManager.WebParts;
                    context.Load(webParts, wps => wps.Include(
                        wp => wp.Id,
                        wp => wp.ZoneId,
                        wp => wp.WebPart.Title,
                        wp => wp.WebPart.ZoneIndex
                    ));
                    context.ExecuteQuery();
                    
                    Console.WriteLine($"Found {webParts.Count} web parts on the page:");
                    Console.WriteLine("==========================================");
                    Console.WriteLine();
                    
                    // Loop through all web parts and print information
                    for (int i = 0; i < webParts.Count; i++)
                    {
                        WebPartDefinition wpDef = webParts[i];
                        WebPart webPart = wpDef.WebPart;
                        
                        // Calculate rendered div ID
                        string renderedDivId = $"MSOZoneCell_WebPartWPQ{webPart.ZoneIndex}";
                        
                        Console.WriteLine($"Web Part #{i + 1}:");
                        Console.WriteLine($"  Title: {webPart.Title}");
                        Console.WriteLine($"  Instance ID: {wpDef.Id}");
                        Console.WriteLine($"  Zone ID: {wpDef.ZoneId}");
                        Console.WriteLine($"  Zone Index: {webPart.ZoneIndex}");
                        Console.WriteLine($"  Rendered Div ID: {renderedDivId}");
                        Console.WriteLine();
                    }
                    
                    if (webParts.Count == 0)
                    {
                        Console.WriteLine("No web parts found on this page.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERROR: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner Exception: {ex.InnerException.Message}");
                }
            }
            
            Console.WriteLine();
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
        

    }
}