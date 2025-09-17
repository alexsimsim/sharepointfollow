#Requires -Version 5.1

<#
.SYNOPSIS
    Makes specified Azure users follow a SharePoint site (Simple Version).

.DESCRIPTION
    This script makes Azure users follow a SharePoint site using values stored directly in the script.
    Simply modify the configuration section below and run the script.

.NOTES
    Author: Generated PowerShell Script
    Requires: PowerShell 5.1 or later
    Requires: Microsoft Graph API permissions:
        - Sites.Read.All
        - Group.Read.All (if using GroupId)
        - User.Read.All
#>

# =============================================================================
# CONFIGURATION SECTION - MODIFY THESE VALUES AS NEEDED
# =============================================================================

# Azure AD Configuration
$TenantID = "contoso.onmicrosoft.com"  # Your Azure AD tenant ID
$ApplicationId = "12345678-1234-1234-1234-123456789012"  # Your Azure application ID
$ApplicationSecret = "your-application-secret-here"  # Your Azure application secret

# SharePoint Site Configuration
$SiteId = "12345678-1234-1234-1234-123456789012"  # The SharePoint site ID to follow
$SiteUrl = "https://contoso.sharepoint.com/sites/MySite"  # Optional: Site URL for logging

# User Configuration - Choose ONE of the following options:

# Option 1: Specify individual users by their IDs/emails
$UserIds = @(
    "user1@contoso.com",
    "user2@contoso.com",
    "user3@contoso.com"
)

# Option 2: Use an Azure AD group (comment out UserIds above and uncomment below)
# $GroupId = "87654321-4321-4321-4321-210987654321"  # Your Azure AD group ID

# =============================================================================
# END CONFIGURATION SECTION - DO NOT MODIFY BELOW THIS LINE
# =============================================================================

# Validate configuration
if (-not $UserIds -and -not $GroupId) {
    Write-Error "Either UserIds or GroupId must be provided in the configuration section above."
    exit 1
}

# Function to get access token
function Get-AccessToken {
    param(
        [string]$TenantId,
        [string]$AppId,
        [string]$AppSecret
    )
    
    try {
        $body = @{
            'resource'      = 'https://graph.microsoft.com'
            'client_id'     = $AppId
            'client_secret' = $AppSecret
            'grant_type'    = "client_credentials"
            'scope'         = "openid"
        }

        $tokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$($TenantId)/oauth2/token" -Body $body -ErrorAction Stop
        return $tokenResponse.access_token
    }
    catch {
        Write-Error "Failed to get access token: $($_.Exception.Message)"
        throw
    }
}

# Function to get users from a group
function Get-UsersFromGroup {
    param(
        [string]$GroupId,
        [hashtable]$Headers
    )
    
    try {
        $groupMembersUri = "https://graph.microsoft.com/v1.0/groups/$GroupId/members"
        $response = Invoke-RestMethod -Uri $groupMembersUri -Headers $Headers -Method Get -ContentType "application/json"
        
        $users = @()
        foreach ($member in $response.value) {
            if ($member.'@odata.type' -eq '#microsoft.graph.user') {
                $users += $member.id
            }
        }
        
        # Handle pagination if needed
        while ($response.'@odata.nextLink') {
            $response = Invoke-RestMethod -Uri $response.'@odata.nextLink' -Headers $Headers -Method Get -ContentType "application/json"
            foreach ($member in $response.value) {
                if ($member.'@odata.type' -eq '#microsoft.graph.user') {
                    $users += $member.id
                }
            }
        }
        
        return $users
    }
    catch {
        Write-Error "Failed to get users from group $GroupId : $($_.Exception.Message)"
        throw
    }
}

# Function to make a user follow a site
function Add-UserFollowSite {
    param(
        [string]$UserId,
        [string]$SiteId,
        [hashtable]$Headers
    )
    
    try {
        $addSitesBody = @{
            value = @(
                @{
                    "id" = $SiteId
                }
            )
        } | ConvertTo-Json -Depth 3

        $followSiteUri = "https://graph.microsoft.com/v1.0/users/$UserId/followedSites/add"
        $result = Invoke-RestMethod -Uri $followSiteUri -Headers $Headers -Method POST -Body $addSitesBody -ContentType "application/json"
        
        return $true
    }
    catch {
        Write-Warning "Failed to make user $UserId follow site $SiteId : $($_.Exception.Message)"
        return $false
    }
}

# Main execution
try {
    Write-Host "Starting SharePoint site following process..." -ForegroundColor Green
    Write-Host "Configuration:" -ForegroundColor Yellow
    Write-Host "  Tenant ID: $TenantID" -ForegroundColor Gray
    Write-Host "  Application ID: $ApplicationId" -ForegroundColor Gray
    Write-Host "  Site ID: $SiteId" -ForegroundColor Gray
    if ($SiteUrl) { Write-Host "  Site URL: $SiteUrl" -ForegroundColor Gray }
    
    # Get access token
    Write-Host "`nGetting access token..." -ForegroundColor Yellow
    $accessToken = Get-AccessToken -TenantId $TenantID -AppId $ApplicationId -AppSecret $ApplicationSecret
    $headers = @{ "Authorization" = "Bearer $accessToken" }
    
    # Get list of users to process
    $usersToProcess = @()
    
    if ($UserIds) {
        Write-Host "Using provided UserIds..." -ForegroundColor Yellow
        $usersToProcess = $UserIds
    }
    elseif ($GroupId) {
        Write-Host "Getting users from group $GroupId..." -ForegroundColor Yellow
        $usersToProcess = Get-UsersFromGroup -GroupId $GroupId -Headers $headers
    }
    
    if ($usersToProcess.Count -eq 0) {
        Write-Warning "No users found to process."
        exit 0
    }
    
    Write-Host "Found $($usersToProcess.Count) users to process." -ForegroundColor Green
    
    # Process each user
    $successCount = 0
    $failureCount = 0
    
    foreach ($userId in $usersToProcess) {
        Write-Host "Processing user: $userId" -ForegroundColor Cyan
        
        $result = Add-UserFollowSite -UserId $userId -SiteId $SiteId -Headers $headers
        
        if ($result) {
            $successCount++
            Write-Host "✓ Successfully made user $userId follow the site" -ForegroundColor Green
        }
        else {
            $failureCount++
            Write-Host "✗ Failed to make user $userId follow the site" -ForegroundColor Red
        }
        
        # Add a small delay to avoid rate limiting
        Start-Sleep -Milliseconds 100
    }
    
    # Summary
    Write-Host "`n=== Summary ===" -ForegroundColor Yellow
    Write-Host "Total users processed: $($usersToProcess.Count)" -ForegroundColor White
    Write-Host "Successful: $successCount" -ForegroundColor Green
    Write-Host "Failed: $failureCount" -ForegroundColor Red
    
    if ($SiteUrl) {
        Write-Host "Site URL: $SiteUrl" -ForegroundColor Cyan
    }
    
    Write-Host "Process completed!" -ForegroundColor Green
}
catch {
    Write-Error "Script execution failed: $($_.Exception.Message)"
    exit 1
}
