# SharePoint Site Follower PowerShell Script

This PowerShell script allows you to automatically make Azure AD users follow a specific SharePoint site. Users can be specified directly via user IDs or through an Azure AD group.

## Script Versions

- **`Follow-SharePointSite.ps1`** - Full-featured version with command-line parameters and stored values
- **`Follow-SharePointSite-Simple.ps1`** - Simplified version with values stored directly in the script (recommended for easy setup)

## Prerequisites

1. **PowerShell 5.1 or later**
2. **Azure AD Application** with the following Microsoft Graph API permissions:
   - `Sites.Read.All` - To read SharePoint sites
   - `User.Read.All` - To read user information
   - `Group.Read.All` - To read group members (if using GroupId parameter)

## Setup Instructions

### 1. Create an Azure AD Application

1. Go to the [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations**
3. Click **New registration**
4. Enter a name for your application
5. Select **Accounts in this organizational directory only**
6. Click **Register**

### 2. Create a Client Secret

1. In your app registration, go to **Certificates & secrets**
2. Click **New client secret**
3. Add a description and select an expiration period
4. Click **Add**
5. **Copy the secret value immediately** (you won't be able to see it again)

### 3. Grant API Permissions

1. In your app registration, go to **API permissions**
2. Click **Add a permission**
3. Select **Microsoft Graph**
4. Choose **Application permissions**
5. Add the following permissions:
   - `Sites.Read.All`
   - `User.Read.All`
   - `Group.Read.All` (if using GroupId parameter)
6. Click **Grant admin consent** (requires admin privileges)

### 4. Get Required Information

- **Tenant ID**: Found in Azure AD > Overview
- **Application ID**: Found in your app registration > Overview
- **Application Secret**: The secret you created in step 2
- **Site ID**: The SharePoint site ID you want users to follow
- **Group ID** (optional): The Azure AD group ID if using group-based assignment

## Usage

### Simple Version (Recommended)

1. Open `Follow-SharePointSite-Simple.ps1`
2. Modify the configuration section at the top of the script:
   ```powershell
   # Azure AD Configuration
   $TenantID = "your-tenant.onmicrosoft.com"
   $ApplicationId = "your-application-id"
   $ApplicationSecret = "your-application-secret"
   
   # SharePoint Site Configuration
   $SiteId = "your-sharepoint-site-id"
   $SiteUrl = "https://your-tenant.sharepoint.com/sites/YourSite"
   
   # User Configuration
   $UserIds = @(
       "user1@yourdomain.com",
       "user2@yourdomain.com",
       "user3@yourdomain.com"
   )
   ```
3. Run the script: `.\Follow-SharePointSite-Simple.ps1`

### Full Version with Parameters

#### Basic Usage with User IDs

```powershell
.\Follow-SharePointSite.ps1 -TenantID "contoso.onmicrosoft.com" -ApplicationId "12345678-1234-1234-1234-123456789012" -ApplicationSecret "your-secret" -SiteId "12345678-1234-1234-1234-123456789012" -UserIds @("user1@contoso.com", "user2@contoso.com")
```

#### Using an Azure AD Group

```powershell
.\Follow-SharePointSite.ps1 -TenantID "contoso.onmicrosoft.com" -ApplicationId "12345678-1234-1234-1234-123456789012" -ApplicationSecret "your-secret" -SiteId "12345678-1234-1234-1234-123456789012" -GroupId "87654321-4321-4321-4321-210987654321"
```

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `TenantID` | String | Yes | Azure AD tenant ID |
| `ApplicationId` | String | Yes | Azure application ID |
| `ApplicationSecret` | String | Yes | Azure application secret |
| `SiteId` | String | Yes | SharePoint site ID to follow |
| `UserIds` | String[] | No* | Array of user IDs to make follow the site |
| `GroupId` | String | No* | Azure AD group ID to get users from |
| `SiteUrl` | String | No | SharePoint site URL (for logging purposes) |

*Either `UserIds` or `GroupId` must be provided.

## Features

- **Flexible User Selection**: Use either specific user IDs or an Azure AD group
- **Error Handling**: Comprehensive error handling with detailed logging
- **Rate Limiting**: Built-in delays to avoid API rate limits
- **Progress Tracking**: Real-time progress updates and summary statistics
- **Validation**: Input validation to ensure required parameters are provided

## Troubleshooting

### Common Issues

1. **Authentication Failed**: Verify your Tenant ID, Application ID, and Application Secret
2. **Insufficient Permissions**: Ensure the application has the required Microsoft Graph permissions
3. **User Not Found**: Verify user IDs are correct and users exist in the tenant
4. **Group Not Found**: Verify the Group ID is correct and the group exists

### Getting Site ID

To find a SharePoint site ID:
1. Go to the SharePoint site
2. Open browser developer tools (F12)
3. Look for the site ID in the page source or network requests
4. Or use the Microsoft Graph API: `https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{sitename}`

## Security Notes

- Store application secrets securely
- Consider using Azure Key Vault for production environments
- Regularly rotate application secrets
- Use least-privilege permissions

## License

This script is provided as-is for educational and administrative purposes.
