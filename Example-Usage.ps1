# Example usage of Follow-SharePointSite.ps1

# Example 1: Make specific users follow a specific SharePoint site
.\Follow-SharePointSite.ps1 `
    -TenantID "contoso.onmicrosoft.com" `
    -ApplicationId "00000000-0000-0000-0000-000000000000" `
    -ApplicationSecret "YourAppSecretHere" `
    -SiteIds @("sites/contoso.sharepoint.com:/sites/Marketing") `
    -UserIds @("user1@contoso.com", "user2@contoso.com")

# Example 2: Make all users in a group follow multiple SharePoint sites
.\Follow-SharePointSite.ps1 `
    -TenantID "contoso.onmicrosoft.com" `
    -ApplicationId "00000000-0000-0000-0000-000000000000" `
    -ApplicationSecret "YourAppSecretHere" `
    -SiteIds @("sites/contoso.sharepoint.com:/sites/Marketing", "sites/contoso.sharepoint.com:/sites/HR") `
    -GroupId "11111111-1111-1111-1111-111111111111"

# Example 3: Make both specific users and users in a group follow a SharePoint site
.\Follow-SharePointSite.ps1 `
    -TenantID "contoso.onmicrosoft.com" `
    -ApplicationId "00000000-0000-0000-0000-000000000000" `
    -ApplicationSecret "YourAppSecretHere" `
    -SiteIds @("sites/contoso.sharepoint.com:/sites/IT") `
    -UserIds @("admin@contoso.com") `
    -GroupId "22222222-2222-2222-2222-222222222222"
