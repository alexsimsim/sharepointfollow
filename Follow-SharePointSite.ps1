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

# Set default values if not provided
if (-not $TenantID) { $TenantID = "contoso.onmicrosoft.com" }
if (-not $ApplicationId) { $ApplicationId = "00000000-0000-0000-0000-000000000000" }
if (-not $ApplicationSecret) { $ApplicationSecret = "REPLACE_WITH_APP_SECRET" }
if (-not $SiteIds) { $SiteIds = @("sites/contoso.sharepoint.com/sites/Marketing") }
if (-not $UserIds) { $UserIds = @("user1@contoso.com", "user2@contoso.com") }
if (-not $UserId) { $UserId = "user1@contoso.com" }

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
        Write-LogMessage "Requesting authentication token from Microsoft Graph..." "DEBUG" "Gray"
        $TokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantID/oauth2/token" -Body $body -ErrorAction Stop
        Write-LogMessage "Authentication token received successfully" "DEBUG" "Gray"
        return $TokenResponse.access_token
    }
    catch {
        $statusCode = ""
        if ($_.Exception.Response) {
            $statusCode = " (HTTP $($_.Exception.Response.StatusCode.value__))"
        }
        Write-LogMessage "Failed to obtain authentication token$statusCode`: $($_.Exception.Message)" "ERROR" "Red"
        Write-LogMessage "Check your TenantID, ApplicationId, and ApplicationSecret values" "INFO" "Yellow"
        exit 1
    }
}

# Function to write timestamped log messages
function Write-LogMessage {
    param (
        [string]$Message,
        [string]$Level = "INFO",
        [string]$Color = "White"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Write-Host $logMessage -ForegroundColor $Color
}

# Function to get user's currently followed sites
function Get-UserFollowedSites {
    param (
        [string]$UserId,
        [string]$Token
    )

    $headers = @{
        "Authorization" = "Bearer $Token"
        "Content-Type"  = "application/json"
    }

    try {
        Write-LogMessage "Retrieving followed sites for user $UserId" "DEBUG" "Gray"
        $response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$UserId/followedSites" -Headers $headers -Method Get -ErrorAction Stop
        return $response.value
    }
    catch {
        Write-LogMessage "Failed to retrieve followed sites for user $UserId`: $($_.Exception.Message)" "ERROR" "Red"
        return $null
    }
}

# Function to get site information for troubleshooting
function Get-SiteInfo {
    param (
        [string]$SiteId,
        [string]$Token
    )

    $headers = @{
        "Authorization" = "Bearer $Token"
        "Content-Type"  = "application/json"
    }

    try {
        Write-LogMessage "Retrieving site information for $SiteId..." "DEBUG" "Gray"
        $response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/$SiteId" -Headers $headers -Method Get -ErrorAction Stop
        Write-LogMessage "Site found: $($response.displayName) ($($response.webUrl))" "DEBUG" "Gray"
        return $response
    }
    catch {
        Write-LogMessage "Failed to retrieve site information for $SiteId`: $($_.Exception.Message)" "WARNING" "Yellow"
        return $null
    }
}

# Function to verify if a site is followed by a user
function Test-SiteIsFollowed {
    param (
        [string]$UserId,
        [string]$SiteId,
        [string]$Token
    )

    $followedSites = Get-UserFollowedSites -UserId $UserId -Token $Token
    
    if ($null -eq $followedSites) {
        Write-LogMessage "Could not verify if site $SiteId is followed by user $UserId (failed to get followed sites)" "WARNING" "Yellow"
        return $false
    }

    # Try multiple matching strategies
    $isFollowed = $followedSites | Where-Object { 
        $_.id -eq $SiteId -or 
        $_.webUrl -contains $SiteId -or
        $_.name -eq $SiteId -or
        $_.displayName -eq $SiteId
    }
    
    if ($isFollowed) {
        Write-LogMessage "✓ Verified: Site $SiteId is followed by user $UserId" "SUCCESS" "Green"
        Write-LogMessage "  Matched site: $($isFollowed.displayName) ($($isFollowed.webUrl))" "DEBUG" "Gray"
        return $true
    }
    else {
        Write-LogMessage "✗ Verification failed: Site $SiteId is NOT followed by user $UserId" "WARNING" "Yellow"
        
        # Try to get site info for troubleshooting
        $siteInfo = Get-SiteInfo -SiteId $SiteId -Token $Token
        if ($siteInfo) {
            Write-LogMessage "  Site exists: $($siteInfo.displayName) ($($siteInfo.webUrl))" "DEBUG" "Gray"
        }
        
        # Show currently followed sites for debugging
        if ($followedSites.Count -gt 0) {
            Write-LogMessage "  User currently follows $($followedSites.Count) sites:" "DEBUG" "Gray"
            foreach ($site in $followedSites) {
                Write-LogMessage "    - $($site.displayName): $($site.id)" "DEBUG" "Gray"
            }
        } else {
            Write-LogMessage "  User is not following any sites currently" "DEBUG" "Gray"
        }
        
        return $false
    }
}

# Function to add a SharePoint site to user's followed sites with verification
function Add-SiteToUserFollowed {
    param (
        [string]$UserId,
        [string]$SiteId,
        [string]$Token,
        [int]$MaxRetries = 3,
        [int]$RetryDelaySeconds = 2
    )

    $headers = @{
        "Authorization" = "Bearer $Token"
        "Content-Type"  = "application/json"
    }

    # Check if site is already followed
    Write-LogMessage "Checking if site $SiteId is already followed by user $UserId..." "INFO" "Cyan"
    if (Test-SiteIsFollowed -UserId $UserId -SiteId $SiteId -Token $Token) {
        Write-LogMessage "Site $SiteId is already followed by user $UserId - skipping" "INFO" "Green"
        return $true
    }

    $body = @{
        "value" = @(
            @{
                "id" = $SiteId
            }
        )
    } | ConvertTo-Json

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            Write-LogMessage "Attempt $attempt of $MaxRetries`: Adding site $SiteId to followed sites for user $UserId" "INFO" "Cyan"
            
            Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$UserId/followedSites/add" -Headers $headers -Method Post -Body $body -ContentType "application/json" -ErrorAction Stop | Out-Null
            
            Write-LogMessage "API call successful - waiting $RetryDelaySeconds seconds before verification..." "INFO" "Yellow"
            Start-Sleep -Seconds $RetryDelaySeconds
            
            # Verify the site was actually added
            if (Test-SiteIsFollowed -UserId $UserId -SiteId $SiteId -Token $Token) {
                Write-LogMessage "Successfully added and verified site $SiteId for user $UserId" "SUCCESS" "Green"
                return $true
            }
            else {
                Write-LogMessage "API call succeeded but verification failed for site $SiteId and user $UserId" "WARNING" "Yellow"
                if ($attempt -lt $MaxRetries) {
                    Write-LogMessage "Retrying in $RetryDelaySeconds seconds..." "INFO" "Yellow"
                    Start-Sleep -Seconds $RetryDelaySeconds
                }
            }
        }
        catch {
            $errorMessage = $_.Exception.Message
            $statusCode = ""
            
            if ($_.Exception.Response) {
                $statusCode = " (HTTP $($_.Exception.Response.StatusCode.value__))"
            }
            
            Write-LogMessage "Attempt $attempt failed to add site $SiteId for user $UserId$statusCode`: $errorMessage" "ERROR" "Red"
            
            if ($attempt -lt $MaxRetries) {
                Write-LogMessage "Retrying in $RetryDelaySeconds seconds..." "INFO" "Yellow"
                Start-Sleep -Seconds $RetryDelaySeconds
            }
        }
    }
    
    Write-LogMessage "Failed to add site $SiteId to followed sites for user $UserId after $MaxRetries attempts" "ERROR" "Red"
    return $false
}

# Main script
Write-LogMessage "Starting SharePoint site following process..." "INFO" "Cyan"
Write-LogMessage "Script Parameters:" "INFO" "Cyan"
Write-LogMessage "  - TenantID: $TenantID" "INFO" "Gray"
Write-LogMessage "  - ApplicationId: $ApplicationId" "INFO" "Gray"
Write-LogMessage "  - SiteIds: $($SiteIds -join ', ')" "INFO" "Gray"
if ($GroupId) { Write-LogMessage "  - GroupId: $GroupId" "INFO" "Gray" }
if ($UserId) { Write-LogMessage "  - UserId: $UserId" "INFO" "Gray" }
if ($UserIds) { Write-LogMessage "  - UserIds: $($UserIds -join ', ')" "INFO" "Gray" }

# Get authentication token
Write-LogMessage "Obtaining authentication token..." "INFO" "Cyan"
$token = Get-AuthToken
Write-LogMessage "Authentication token obtained successfully" "SUCCESS" "Green"

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
        Write-LogMessage "Fetching users from group $GroupId..." "INFO" "Cyan"
        $groupUsers = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/members" -Headers $headers -Method Get -ErrorAction Stop).value
        
        if ($groupUsers) {
            $groupUserIds = $groupUsers | Select-Object -ExpandProperty id
            $allUserIds += $groupUserIds
            Write-LogMessage "Found $($groupUserIds.Count) users in the specified group." "SUCCESS" "Green"
            Write-LogMessage "Group users: $($groupUserIds -join ', ')" "DEBUG" "Gray"
        }
        else {
            Write-LogMessage "No users found in the specified group." "WARNING" "Yellow"
        }
    }
    catch {
        Write-LogMessage "Error fetching users from group $GroupId`: $($_.Exception.Message)" "ERROR" "Red"
    }
}

# Remove any duplicate user IDs
$allUserIds = $allUserIds | Select-Object -Unique

Write-LogMessage "Processing $($allUserIds.Count) unique users for $($SiteIds.Count) sites" "INFO" "Cyan"
Write-LogMessage "Users to process: $($allUserIds -join ', ')" "DEBUG" "Gray"

if ($allUserIds.Count -eq 0) {
    Write-LogMessage "No users found to process." "ERROR" "Red"
    exit 1
}

# Process each user and make them follow each site
$successCount = 0
$failureCount = 0
$verificationFailures = @()

Write-LogMessage "Starting follow operations..." "INFO" "Cyan"

foreach ($userId in $allUserIds) {
    foreach ($siteId in $SiteIds) {
        Write-LogMessage "Processing user $userId for site $siteId..." "INFO" "Cyan"
        
        $result = Add-SiteToUserFollowed -UserId $userId -SiteId $siteId -Token $token
        
        if ($result) {
            $successCount++
        }
        else {
            $failureCount++
            $verificationFailures += @{
                UserId = $userId
                SiteId = $siteId
            }
        }
        
        # Add a small delay between operations to avoid rate limiting
        Start-Sleep -Milliseconds 500
    }
}

# Final verification of all operations
Write-LogMessage "`nPerforming final verification of all follow operations..." "INFO" "Cyan"

$finalVerificationResults = @()
foreach ($userId in $allUserIds) {
    foreach ($siteId in $SiteIds) {
        $isFollowed = Test-SiteIsFollowed -UserId $userId -SiteId $siteId -Token $token
        $finalVerificationResults += @{
            UserId = $userId
            SiteId = $siteId
            IsFollowed = $isFollowed
        }
    }
}

# Summary
Write-LogMessage "`nProcess completed!" "INFO" "Cyan"
Write-LogMessage "================== SUMMARY ==================" "INFO" "White"
Write-LogMessage "Successful operations: $successCount" "SUCCESS" "Green"
$failureColor = if ($failureCount -gt 0) { "Red" } else { "Green" }
Write-LogMessage "Failed operations: $failureCount" "INFO" $failureColor

if ($verificationFailures.Count -gt 0) {
    Write-LogMessage "`nFailed operations details:" "WARNING" "Yellow"
    foreach ($failure in $verificationFailures) {
        Write-LogMessage "  - User: $($failure.UserId), Site: $($failure.SiteId)" "WARNING" "Yellow"
    }
}

# Final verification summary
$verifiedCount = ($finalVerificationResults | Where-Object { $_.IsFollowed }).Count
$totalExpected = $finalVerificationResults.Count
$unverifiedCount = $totalExpected - $verifiedCount

Write-LogMessage "`n============= FINAL VERIFICATION ==============" "INFO" "White"
Write-LogMessage "Expected follow relationships: $totalExpected" "INFO" "Cyan"
Write-LogMessage "Verified as following: $verifiedCount" "SUCCESS" "Green"
Write-LogMessage "Not following (verification failed): $unverifiedCount" "WARNING" $(if ($unverifiedCount -gt 0) { "Red" } else { "Green" })

if ($unverifiedCount -gt 0) {
    Write-LogMessage "`nUnverified follow relationships:" "WARNING" "Yellow"
    $unverified = $finalVerificationResults | Where-Object { -not $_.IsFollowed }
    foreach ($item in $unverified) {
        Write-LogMessage "  - User: $($item.UserId), Site: $($item.SiteId)" "WARNING" "Yellow"
    }
    Write-LogMessage "`nRecommendation: Check SharePoint permissions and site accessibility for the above relationships." "INFO" "Yellow"
}

if ($verifiedCount -eq $totalExpected) {
    Write-LogMessage "`n✓ SUCCESS: All expected follow relationships have been verified!" "SUCCESS" "Green"
} else {
    Write-LogMessage "`n⚠ WARNING: Some follow relationships could not be verified. Check the details above." "WARNING" "Yellow"
}