# SharePoint 2016 Authentication Script
# Run this once to authenticate, then use the main script multiple times

param(
    [Parameter(Mandatory=$false)]
    [string]$SiteUrl = "https://consulting.global.deloitteonline.com/sites/Aflac/POC"
)

# Import required module
try {
    Import-Module SharePointPnPPowerShell2016 -ErrorAction Stop
    Write-Host "SharePointPnPPowerShell2016 module loaded successfully" -ForegroundColor Green
}
catch {
    Write-Error "Failed to load SharePointPnPPowerShell2016 module. Please ensure it's installed."
    Write-Host "Install with: Install-Module SharePointPnPPowerShell2016" -ForegroundColor Yellow
    exit 1
}

try {
    # Connect to SharePoint with MFA support
    Write-Host "Connecting to SharePoint site: $SiteUrl" -ForegroundColor Yellow
    Write-Host "Please complete authentication in the browser window..." -ForegroundColor Yellow
    Connect-PnPOnline -Url $SiteUrl -UseWebLogin
    Write-Host "Successfully connected to SharePoint" -ForegroundColor Green
    
    # Test the connection
    $web = Get-PnPWeb
    Write-Host "Connection verified. Connected to: $($web.Title)" -ForegroundColor Green
    Write-Host "Server Relative URL: $($web.ServerRelativeUrl)" -ForegroundColor Green
    
    Write-Host "`n=== AUTHENTICATION SUCCESS ===" -ForegroundColor Cyan
    Write-Host "You are now authenticated to SharePoint." -ForegroundColor White
    Write-Host "You can now run the main script multiple times without re-authenticating." -ForegroundColor White
    Write-Host "Connection will remain active for this PowerShell session." -ForegroundColor White
    Write-Host "`nTo run the metadata extraction script:" -ForegroundColor Yellow
    Write-Host ".\Get-SharePointPageMetadata-NoAuth.ps1" -ForegroundColor Green
}
catch {
    Write-Error "Authentication failed: $($_.Exception.Message)"
    Write-Host "Please check your credentials and try again." -ForegroundColor Red
    exit 1
}