# Example usage of Follow-SharePointSite.ps1

# Example 1: Using specific user IDs
$params = @{
    TenantID = "contoso.onmicrosoft.com"
    ApplicationId = "12345678-1234-1234-1234-123456789012"
    ApplicationSecret = "your-application-secret-here"
    SiteId = "12345678-1234-1234-1234-123456789012"
    UserIds = @(
        "user1@contoso.com",
        "user2@contoso.com",
        "user3@contoso.com"
    )
    SiteUrl = "https://contoso.sharepoint.com/sites/MySite"
}

# Execute the script
& ".\Follow-SharePointSite.ps1" @params

# Example 2: Using an AD group
$params2 = @{
    TenantID = "contoso.onmicrosoft.com"
    ApplicationId = "12345678-1234-1234-1234-123456789012"
    ApplicationSecret = "your-application-secret-here"
    SiteId = "12345678-1234-1234-1234-123456789012"
    GroupId = "87654321-4321-4321-4321-210987654321"
    SiteUrl = "https://contoso.sharepoint.com/sites/MySite"
}

# Execute the script
& ".\Follow-SharePointSite.ps1" @params2
