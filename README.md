# SharePoint Site Follower PowerShell Scripts

These PowerShell scripts help you make Azure AD users follow specific SharePoint sites automatically.

## Scripts Overview

1. **Follow-SharePointSite.ps1** - Full-featured script with comprehensive error handling and detailed output
2. **Follow-SharePointSite-Simple.ps1** - Simplified version for quick usage
3. **Get-UserFollowedSites.ps1** - Retrieves and displays SharePoint sites that a specific user is following
4. **Example-Usage.ps1** - Example commands showing how to use the main script

## Prerequisites

- PowerShell 5.1 or higher
- Azure AD App Registration with the following API permissions:
  - Microsoft Graph API > Application permissions:
    - `Group.Read.All` (to read group members)
    - `User.Read.All` (to read user information)
    - `Sites.ReadWrite.All` (to modify followed sites)

## Usage Instructions

### Setting Up Azure AD App Registration

1. Go to the Azure Portal and navigate to Azure Active Directory > App registrations
2. Create a new registration
3. Add the required API permissions listed above
4. Create a client secret
5. Note down the Application (client) ID, Directory (tenant) ID, and client secret

### Basic Usage

#### Making Users Follow Sites

```powershell
.\Follow-SharePointSite.ps1 `
    -TenantID "yourtenant.onmicrosoft.com" `
    -ApplicationId "your-application-id" `
    -ApplicationSecret "your-application-secret" `
    -SiteIds @("sites/yourtenant.sharepoint.com:/sites/YourSiteName") `
    -UserIds @("user1@yourtenant.com", "user2@yourtenant.com")
```

#### Getting Sites a User is Following

```powershell
# Basic usage with default table output
.\Get-UserFollowedSites.ps1 -UserId "user@yourtenant.com"

# With full configuration and CSV export
.\Get-UserFollowedSites.ps1 `
    -TenantID "yourtenant.onmicrosoft.com" `
    -ApplicationId "your-application-id" `
    -ApplicationSecret "your-application-secret" `
    -UserId "user@yourtenant.com" `
    -OutputFormat "CSV" `
    -OutputFile "user-followed-sites.csv" `
    -IncludeDetails

# JSON output with detailed information
.\Get-UserFollowedSites.ps1 `
    -UserId "user@yourtenant.com" `
    -OutputFormat "JSON" `
    -IncludeDetails
```

### Getting SharePoint Site IDs

To get the ID of a SharePoint site, you can use the following PowerShell command:

```powershell
$token = "your-access-token"
$headers = @{
    "Authorization" = "Bearer $token"
    "Content-Type" = "application/json"
}

# For a specific site
$siteUrl = "https://yourtenant.sharepoint.com/sites/YourSiteName"
$encodedUrl = [System.Web.HttpUtility]::UrlEncode($siteUrl)
Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites?search=$encodedUrl" -Headers $headers -Method Get
```

## Feature Comparison

| Feature | Follow-SharePointSite.ps1 | Follow-SharePointSite-Simple.ps1 | Get-UserFollowedSites.ps1 |
|---------|---------------------------|--------------------------------|---------------------------|
| Detailed logging | ✅ | ❌ | ✅ |
| Error handling | ✅ | ✅ (basic) | ✅ |
| Parameter validation | ✅ | ❌ | ✅ |
| Summary statistics | ✅ | ❌ | ✅ |
| Configuration via parameters | ✅ | ❌ | ✅ |
| Configuration via script variables | ❌ | ✅ | ❌ |
| Multiple output formats | ❌ | ❌ | ✅ (Table, List, CSV, JSON) |
| Export to file | ❌ | ❌ | ✅ |
| Detailed site information | ❌ | ❌ | ✅ (optional) |
| User validation | ❌ | ❌ | ✅ |

## Credit

This script was inspired by Kelvin Tegelaar's blog post: [Automating with PowerShell: Automatically following all SharePoint Sites or Teams for all users](https://www.cyberdrain.com/automating-with-powershell-automatically-following-all-sharepoint-sites-or-teas-for-all-users/)
