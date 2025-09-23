<#
.SYNOPSIS
Makes Azure AD users follow specific SharePoint sites

.DESCRIPTION
This script takes in a list of Azure AD users, an optional AD group, and SharePoint sites and makes those users follow the specified sites.

.PARAMETER TenantID
Your Microsoft 365 tenant ID (e.g. "contoso.onmicrosoft.com")

.PARAMETER ApplicationId
The Application (client) ID of your registered Azure AD application

.PARAMETER ApplicationSecret
The client secret of your registered Azure AD application

.PARAMETER UserId
(Optional) A single Azure AD user ID or UPN who should follow the SharePoint sites

.PARAMETER GroupId
(Optional) The ID of an Azure AD group whose members should follow the SharePoint sites

.PARAMETER SiteIds
An array of SharePoint site IDs that users should follow

.PARAMETER UserIds
An array of Azure AD user IDs who should follow the SharePoint sites

.EXAMPLE
.\Follow-SharePointSite.ps1 -SiteIds @("sites/contoso.sharepoint.com/sites/Marketing") -UserIds @("user1@contoso.com", "user2@contoso.com")

.EXAMPLE
.\Follow-SharePointSite.ps1 -SiteIds @("sites/contoso.sharepoint.com/sites/Marketing") -UserId "user1@contoso.com"

.EXAMPLE
.\Follow-SharePointSite.ps1 -TenantID "contoso.onmicrosoft.com" -ApplicationId "1234abcd-1234-abcd-1234-1234abcd1234" -ApplicationSecret "YourAppSecret" -SiteIds @("sites/contoso.sharepoint.com/sites/Marketing", "sites/contoso.sharepoint.com/sites/HR") -GroupId "5678efgh-5678-efgh-5678-5678efgh5678"
#>

[string]$DEFAULT_TENANT_ID = "contoso.onmicrosoft.com"
[string]$DEFAULT_APPLICATION_ID = "00000000-0000-0000-0000-000000000000"
[string]$DEFAULT_APPLICATION_SECRET = "REPLACE_WITH_APP_SECRET"

# Defaults for users and sites
[string[]]$DEFAULT_SITE_IDS = @(
    "sites/contoso.sharepoint.com/sites/Marketing"
)
[string[]]$DEFAULT_USER_IDS = @(
    "user1@contoso.com",
    "user2@contoso.com"
)
[string]$DEFAULT_USER_ID = "user1@contoso.com"

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$TenantID = $DEFAULT_TENANT_ID,
    
    [Parameter(Mandatory = $false)]
    [string]$ApplicationId = $DEFAULT_APPLICATION_ID,
    
    [Parameter(Mandatory = $false)]
    [string]$ApplicationSecret = $DEFAULT_APPLICATION_SECRET,
    
    [Parameter(Mandatory = $false)]
    [string]$GroupId,
    
    [Parameter(Mandatory = $false)]
    [array]$SiteIds = $DEFAULT_SITE_IDS,
    
    [Parameter(Mandatory = $false)]
    [array]$UserIds = $DEFAULT_USER_IDS,

    [Parameter(Mandatory = $false)]
    [string]$UserId = $DEFAULT_USER_ID
)

# Check if at least one of GroupId, UserIds, or UserId is provided
if (-not $GroupId -and -not $UserIds -and -not $UserId) {
    Write-Error "Either GroupId, UserIds, or UserId must be specified."
    exit 1
}

# Function to get authentication token
function Get-AuthToken {
    $body = @{
        'resource'      = 'https://graph.microsoft.com'
        'client_id'     = $ApplicationId
        'client_secret' = $ApplicationSecret
        'grant_type'    = "client_credentials"
        'scope'         = "openid"
    }

    try {
        $TokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantID/oauth2/token" -Body $body -ErrorAction Stop
        return $TokenResponse.access_token
    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Error "Failed to obtain authentication token: $errorMessage"
        exit 1
    }
}

# Function to add a SharePoint site to user's followed sites
function Add-SiteToUserFollowed {
    param (
        [string]$UserId,
        [string]$SiteId,
        [string]$Token
    )

    $headers = @{
        "Authorization" = "Bearer $Token"
        "Content-Type"  = "application/json"
    }

    $body = @{
        "value" = @(
            @{
                "id" = $SiteId
            }
        )
    } | ConvertTo-Json

    try {
        Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$UserId/followedSites/add" -Headers $headers -Method Post -Body $body -ContentType "application/json"
        Write-Host "Successfully added site $SiteId to followed sites for user $UserId" -ForegroundColor Green
        return $true
    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Host "Failed to add site $SiteId to followed sites for user $UserId: $errorMessage" -ForegroundColor Red
        return $false
    }
}

# Main script
Write-Host "Starting SharePoint site following process..." -ForegroundColor Cyan

# Get authentication token
$token = Get-AuthToken
$headers = @{
    "Authorization" = "Bearer $token"
    "Content-Type"  = "application/json"
}

# Initialize an array to store all user IDs
$allUserIds = @()

# If UserIds is provided, add them to the array
if ($UserIds) {
    $allUserIds += $UserIds
}

# If single UserId is provided, add it to the array
if ($UserId) {
    $allUserIds += $UserId
}

# If GroupId is provided, get users from the group
if ($GroupId) {
    try {
        Write-Host "Fetching users from group $GroupId..." -ForegroundColor Cyan
        $groupUsers = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/members" -Headers $headers -Method Get).value
        
        if ($groupUsers) {
            $groupUserIds = $groupUsers | Select-Object -ExpandProperty id
            $allUserIds += $groupUserIds
            Write-Host "Found $($groupUserIds.Count) users in the specified group." -ForegroundColor Cyan
        }
        else {
            Write-Host "No users found in the specified group." -ForegroundColor Yellow
        }
    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Host "Error fetching users from group: $errorMessage" -ForegroundColor Red
    }
}

# Remove any duplicate user IDs
$allUserIds = $allUserIds | Select-Object -Unique

if ($allUserIds.Count -eq 0) {
    Write-Error "No users found to process."
    exit 1
}

# Process each user and make them follow each site
$successCount = 0
$failureCount = 0

foreach ($userId in $allUserIds) {
    foreach ($siteId in $SiteIds) {
        Write-Host "Processing user $userId for site $siteId..." -ForegroundColor Cyan
        
        $result = Add-SiteToUserFollowed -UserId $userId -SiteId $siteId -Token $token
        
        if ($result) {
            $successCount++
        }
        else {
            $failureCount++
        }
    }
}

# Summary
Write-Host "`nProcess completed!" -ForegroundColor Cyan
Write-Host "Successful operations: $successCount" -ForegroundColor Green
$failureColor = if ($failureCount -gt 0) { "Red" } else { "Green" }
Write-Host "Failed operations: $failureCount" -ForegroundColor $failureColor

