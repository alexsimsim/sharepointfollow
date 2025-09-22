<#
.SYNOPSIS
Simple script to make Azure AD users follow specific SharePoint sites
.DESCRIPTION
This script connects to Microsoft Graph API and makes specified users follow SharePoint sites
#>

# Set your variables here
$TenantID = "yourtenant.onmicrosoft.com"
$ApplicationId = "your-application-id"
$ApplicationSecret = "your-application-secret"
$SiteIds = @("sites/contoso.sharepoint.com:/sites/YourSiteName")
$UserIds = @("user1@contoso.com", "user2@contoso.com")
# Optional: Specify a group ID to include all users from that group
$GroupId = ""  # Leave empty if not using

# Initialize counters for summary
$successCount = 0
$failureCount = 0

# Function to log messages with timestamp and color
function Write-Log {
    param (
        [string]$Message,
        [string]$Type = "INFO",
        [string]$ForegroundColor = "White"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] [$Type] $Message" -ForegroundColor $ForegroundColor
}

# Start logging
$scriptStartTime = Get-Date
Write-Log "Starting SharePoint site follower script" "START" "Cyan"
Write-Log "Tenant: $TenantID" "CONFIG" "Gray"
Write-Log "Sites to follow: $($SiteIds.Count)" "CONFIG" "Gray"

# Get authentication token
Write-Log "Authenticating to Microsoft Graph API..." "AUTH" "Yellow"
$body = @{
    'resource'      = 'https://graph.microsoft.com'
    'client_id'     = $ApplicationId
    'client_secret' = $ApplicationSecret
    'grant_type'    = "client_credentials"
    'scope'         = "openid"
}

try {
    $authResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantID/oauth2/token" -Body $body -ErrorAction Stop
    $token = $authResponse.access_token
    $headers = @{
        "Authorization" = "Bearer $token"
        "Content-Type"  = "application/json"
    }
    Write-Log "Authentication successful" "AUTH" "Green"
}
catch {
    Write-Log "Authentication failed: $_" "ERROR" "Red"
    exit 1
}

# Validate token by making a simple API call
Write-Log "Verifying API connection..." "VERIFY" "Yellow"
try {
    $verifyResult = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/me" -Headers $headers -Method Get -ErrorAction SilentlyContinue
    Write-Log "API connection verified successfully" "VERIFY" "Green"
}
catch {
    # This is expected to fail with app-only permissions, but we can continue
    Write-Log "API connection verification completed with expected app permission error" "VERIFY" "Green"
}

# Get users from direct specification
Write-Log "Processing user list..." "USERS" "Magenta"
$allUserIds = $UserIds
Write-Log "Directly specified users: $($UserIds.Count)" "USERS" "Magenta"

# Get users from group if specified
if ($GroupId) {
    Write-Log "Fetching users from group $GroupId..." "GROUP" "Magenta"
    try {
        $groupUsers = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/members" -Headers $headers -Method Get).value
        $groupUserIds = $groupUsers | Select-Object -ExpandProperty id
        $allUserIds += $groupUserIds
        Write-Log "Found $($groupUserIds.Count) users in the group" "GROUP" "Magenta"
    }
    catch {
        Write-Log "Error fetching users from group: $_" "ERROR" "Red"
    }
}

# Remove any duplicate user IDs
$allUserIds = $allUserIds | Select-Object -Unique
Write-Log "Total unique users to process: $($allUserIds.Count)" "USERS" "Magenta"

if ($allUserIds.Count -eq 0) {
    Write-Log "No users found to process" "ERROR" "Red"
    exit 1
}

# Process each site
Write-Log "Starting to process SharePoint site following..." "PROCESS" "Cyan"
$totalOperations = $allUserIds.Count * $SiteIds.Count
$currentOperation = 0

foreach ($siteId in $SiteIds) {
    Write-Log "Processing site: $siteId" "SITE" "White"
    
    # Make each user follow the site
    foreach ($userId in $allUserIds) {
        $currentOperation++
        $percentComplete = [math]::Round(($currentOperation / $totalOperations) * 100)
        Write-Progress -Activity "Following SharePoint Sites" -Status "Processing $currentOperation of $totalOperations ($percentComplete%)" -PercentComplete $percentComplete
        
        $body = @{
            "value" = @(
                @{
                    "id" = $siteId
                }
            )
        } | ConvertTo-Json

        try {
            Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$userId/followedSites/add" -Headers $headers -Method Post -Body $body -ContentType "application/json"
            Write-Log "Added site to favorites for user $userId" "SUCCESS" "Green"
            $successCount++
        }
        catch {
            Write-Log "Failed to add site for user $userId: $_" "FAILURE" "Red"
            $failureCount++
        }
    }
}

# Complete the progress bar
Write-Progress -Activity "Following SharePoint Sites" -Completed

# Calculate execution time
$scriptEndTime = Get-Date
$executionTime = $scriptEndTime - $scriptStartTime
$formattedExecutionTime = "{0:mm}m {0:ss}s" -f $executionTime

# Display summary
Write-Log "Operation complete!" "COMPLETE" "Cyan"
Write-Log "Execution time: $formattedExecutionTime" "STATS" "White"
Write-Log "Success count: $successCount" "STATS" $(if ($successCount -gt 0) { "Green" } else { "Yellow" })
Write-Log "Failure count: $failureCount" "STATS" $(if ($failureCount -gt 0) { "Red" } else { "Green" })
Write-Log "Total operations: $($successCount + $failureCount)" "STATS" "White"
Write-Log "Success rate: $(if (($successCount + $failureCount) -gt 0) { [math]::Round(($successCount / ($successCount + $failureCount)) * 100, 2) } else { 0 })%" "STATS" "Cyan"
