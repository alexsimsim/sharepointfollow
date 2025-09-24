<#
.SYNOPSIS
Retrieves a list of SharePoint sites that a specific user is following

.DESCRIPTION
This script connects to Microsoft Graph API and retrieves all SharePoint sites that a specified user is following.
It provides detailed information about each site and offers various output formats.

.PARAMETER TenantID
Your Microsoft 365 tenant ID (e.g. "contoso.onmicrosoft.com")

.PARAMETER ApplicationId
The Application (client) ID of your registered Azure AD application

.PARAMETER ApplicationSecret
The client secret of your registered Azure AD application

.PARAMETER UserId
The Azure AD user ID or UPN whose followed sites you want to retrieve

.PARAMETER OutputFormat
The format for the output. Options: "Table", "List", "CSV", "JSON"
Default: "Table"

.PARAMETER OutputFile
Optional file path to save the results. The file extension should match the OutputFormat

.PARAMETER IncludeDetails
Include additional details like site description, last modified date, etc.

.EXAMPLE
.\Get-UserFollowedSites.ps1 -UserId "user@contoso.com"

.EXAMPLE
.\Get-UserFollowedSites.ps1 -TenantID "contoso.onmicrosoft.com" -ApplicationId "1234abcd-1234-abcd-1234-1234abcd1234" -ApplicationSecret "YourAppSecret" -UserId "user@contoso.com" -OutputFormat "CSV" -OutputFile "followed-sites.csv"

.EXAMPLE
.\Get-UserFollowedSites.ps1 -UserId "user@contoso.com" -IncludeDetails -OutputFormat "JSON"
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$TenantID,
    
    [Parameter(Mandatory=$false)]
    [string]$ApplicationId,
    
    [Parameter(Mandatory=$false)]
    [string]$ApplicationSecret,
    
    [Parameter(Mandatory=$true)]
    [string]$UserId,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("Table", "List", "CSV", "JSON")]
    [string]$OutputFormat = "Table",
    
    [Parameter(Mandatory=$false)]
    [string]$OutputFile,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeDetails
)

# Set default values if not provided
if (-not $TenantID) { $TenantID = "contoso.onmicrosoft.com" }
if (-not $ApplicationId) { $ApplicationId = "00000000-0000-0000-0000-000000000000" }
if (-not $ApplicationSecret) { $ApplicationSecret = "REPLACE_WITH_APP_SECRET" }

# Load required assemblies
Add-Type -AssemblyName System.Web

# Validate required parameters
if (-not $UserId) {
    Write-Error "UserId parameter is required."
    exit 1
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

# Function to get authentication token
function Get-AuthToken {
    $body = @{
        'client_id'     = $ApplicationId
        'client_secret' = $ApplicationSecret
        'scope'         = 'https://graph.microsoft.com/.default'
        'grant_type'    = "client_credentials"
    }

    try {
        Write-LogMessage "Requesting authentication token from Microsoft Graph..." "DEBUG" "Gray"
        $TokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token" -Body $body -ErrorAction Stop
        Write-LogMessage "Authentication token received successfully" "DEBUG" "Gray"
        return $TokenResponse.access_token
    }
    catch {
        $statusCode = ""
        $errorDetails = ""
        if ($_.Exception.Response) {
            $statusCode = " (HTTP $($_.Exception.Response.StatusCode.value__))"
            try {
                $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                $errorDetails = $reader.ReadToEnd()
                $reader.Close()
                if ($errorDetails) {
                    $errorDetails = " | Response: $errorDetails"
                }
            }
            catch {
                # If we can't read the response stream, continue without it
            }
        }
        Write-LogMessage "Failed to obtain authentication token$statusCode`: $($_.Exception.Message)$errorDetails" "ERROR" "Red"
        Write-LogMessage "Check your TenantID, ApplicationId, and ApplicationSecret values" "INFO" "Yellow"
        Write-LogMessage "Ensure your app has the following Microsoft Graph permissions: Sites.Read.All, Sites.ReadWrite.All" "INFO" "Yellow"
        exit 1
    }
}

# Function to validate user exists
function Test-UserExists {
    param (
        [string]$UserId,
        [string]$Token
    )

    $headers = @{
        "Authorization" = "Bearer $Token"
        "Content-Type"  = "application/json"
    }

    try {
        Write-LogMessage "Validating user $UserId exists..." "DEBUG" "Gray"
        $response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$UserId" -Headers $headers -Method Get -ErrorAction Stop
        Write-LogMessage "User found: $($response.displayName) ($($response.userPrincipalName))" "SUCCESS" "Green"
        return $response
    }
    catch {
        $statusCode = ""
        if ($_.Exception.Response) {
            $statusCode = " (HTTP $($_.Exception.Response.StatusCode.value__))"
        }
        Write-LogMessage "User $UserId not found$statusCode`: $($_.Exception.Message)" "ERROR" "Red"
        return $null
    }
}

# Function to test API connectivity and permissions
function Test-APIConnectivity {
    param (
        [string]$Token
    )

    $headers = @{
        "Authorization" = "Bearer $Token"
        "Content-Type"  = "application/json"
    }

    Write-LogMessage "Testing API connectivity and permissions..." "INFO" "Cyan"
    
    # Test 1: Basic Graph API connectivity
    try {
        Write-LogMessage "Test 1: Basic Microsoft Graph connectivity..." "DEBUG" "Gray"
        $response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/me" -Headers $headers -Method Get -ErrorAction Stop
        Write-LogMessage "✓ Basic Graph API connectivity successful" "SUCCESS" "Green"
    }
    catch {
        Write-LogMessage "✗ Basic Graph API connectivity failed: $($_.Exception.Message)" "ERROR" "Red"
        return $false
    }
    
    # Test 2: Sites endpoint access
    try {
        Write-LogMessage "Test 2: Sites endpoint access..." "DEBUG" "Gray"
        $response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites" -Headers $headers -Method Get -ErrorAction Stop
        Write-LogMessage "✓ Sites endpoint accessible" "SUCCESS" "Green"
    }
    catch {
        Write-LogMessage "✗ Sites endpoint access failed: $($_.Exception.Message)" "ERROR" "Red"
        if ($_.Exception.Response) {
            $statusCode = $_.Exception.Response.StatusCode.value__
            if ($statusCode -eq 403) {
                Write-LogMessage "This indicates missing Sites permissions. Ensure your app has Sites.Read.All or Sites.ReadWrite.All" "INFO" "Yellow"
            }
        }
        return $false
    }
    
    Write-LogMessage "✓ API connectivity tests passed" "SUCCESS" "Green"
    return $true
}

# Function to get user's followed sites with detailed information
function Get-UserFollowedSites {
    param (
        [string]$UserId,
        [string]$Token,
        [bool]$IncludeDetails = $false
    )

    $headers = @{
        "Authorization" = "Bearer $Token"
        "Content-Type"  = "application/json"
    }

    try {
        Write-LogMessage "Retrieving followed sites for user $UserId..." "INFO" "Cyan"
        Write-LogMessage "API Endpoint: https://graph.microsoft.com/v1.0/users/$UserId/followedSites" "DEBUG" "Gray"
        $response = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$UserId/followedSites" -Headers $headers -Method Get -ErrorAction Stop
        
        $sites = $response.value
        Write-LogMessage "Found $($sites.Count) followed sites" "SUCCESS" "Green"
        
        # Debug: Show site information structure
        if ($sites.Count -gt 0) {
            Write-LogMessage "Sample site structure for debugging:" "DEBUG" "Gray"
            $sampleSite = $sites[0]
            Write-LogMessage "  - ID: $($sampleSite.id)" "DEBUG" "Gray"
            Write-LogMessage "  - Display Name: $($sampleSite.displayName)" "DEBUG" "Gray"
            Write-LogMessage "  - Web URL: $($sampleSite.webUrl)" "DEBUG" "Gray"
        }

        if ($IncludeDetails -and $sites.Count -gt 0) {
            Write-LogMessage "Fetching detailed information for each site..." "INFO" "Cyan"
            Write-LogMessage "Note: Some detailed information may not be available due to API limitations" "INFO" "Yellow"
            
            # Enhance each site with additional details
            for ($i = 0; $i -lt $sites.Count; $i++) {
                $site = $sites[$i]
                try {
                    Write-Progress -Activity "Fetching site details" -Status "Processing site $($i + 1) of $($sites.Count): $($site.displayName)" -PercentComplete (($i + 1) / $sites.Count * 100)
                    
                    # Try multiple approaches to get site details
                    $detailedSite = $null
                    
                    # First try with the direct site ID
                    try {
                        $detailedSite = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)" -Headers $headers -Method Get -ErrorAction Stop
                    }
                    catch {
                        # If that fails, try with the webUrl
                        if ($site.webUrl) {
                            try {
                                $encodedUrl = [System.Web.HttpUtility]::UrlEncode($site.webUrl)
                                $searchResponse = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites?search=$encodedUrl" -Headers $headers -Method Get -ErrorAction Stop
                                if ($searchResponse.value -and $searchResponse.value.Count -gt 0) {
                                    $detailedSite = $searchResponse.value[0]
                                }
                            }
                            catch {
                                Write-LogMessage "Could not fetch detailed info for site $($site.displayName) using webUrl approach" "DEBUG" "Gray"
                            }
                        }
                    }
                    
                    if ($detailedSite) {
                        # Add additional properties
                        $site | Add-Member -NotePropertyName "Description" -NotePropertyValue $detailedSite.description -Force
                        $site | Add-Member -NotePropertyName "LastModifiedDateTime" -NotePropertyValue $detailedSite.lastModifiedDateTime -Force
                        $site | Add-Member -NotePropertyName "SiteCollection" -NotePropertyValue $detailedSite.siteCollection -Force
                        Write-LogMessage "Enhanced details for site: $($site.displayName)" "DEBUG" "Gray"
                    } else {
                        # Add placeholder values if we couldn't get details
                        $site | Add-Member -NotePropertyName "Description" -NotePropertyValue "Not available" -Force
                        $site | Add-Member -NotePropertyName "LastModifiedDateTime" -NotePropertyValue $null -Force
                        $site | Add-Member -NotePropertyName "SiteCollection" -NotePropertyValue $null -Force
                        Write-LogMessage "Could not fetch enhanced details for site: $($site.displayName)" "DEBUG" "Gray"
                    }
                    
                    Start-Sleep -Milliseconds 300  # Increased rate limiting
                }
                catch {
                    Write-LogMessage "Failed to get detailed info for site $($site.displayName): $($_.Exception.Message)" "WARNING" "Yellow"
                    # Add placeholder values on error
                    $site | Add-Member -NotePropertyName "Description" -NotePropertyValue "Error retrieving" -Force
                    $site | Add-Member -NotePropertyName "LastModifiedDateTime" -NotePropertyValue $null -Force
                    $site | Add-Member -NotePropertyName "SiteCollection" -NotePropertyValue $null -Force
                }
            }
            Write-Progress -Activity "Fetching site details" -Completed
        }

        return $sites
    }
    catch {
        $statusCode = ""
        $errorDetails = ""
        if ($_.Exception.Response) {
            $statusCode = " (HTTP $($_.Exception.Response.StatusCode.value__))"
            try {
                $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                $errorDetails = $reader.ReadToEnd()
                $reader.Close()
                if ($errorDetails) {
                    $errorDetails = " | Response: $errorDetails"
                }
            }
            catch {
                # If we can't read the response stream, continue without it
            }
        }
        Write-LogMessage "Failed to retrieve followed sites for user $UserId$statusCode`: $($_.Exception.Message)$errorDetails" "ERROR" "Red"
        
        # Provide specific guidance for common errors
        if ($statusCode -like "*500*") {
            Write-LogMessage "HTTP 500 errors can indicate:" "INFO" "Yellow"
            Write-LogMessage "  1. Insufficient API permissions (need Sites.Read.All or Sites.ReadWrite.All)" "INFO" "Yellow"
            Write-LogMessage "  2. User does not have access to followed sites feature" "INFO" "Yellow"
            Write-LogMessage "  3. Tenant configuration issue with SharePoint following" "INFO" "Yellow"
            Write-LogMessage "  4. Try using a different user ID or check user permissions" "INFO" "Yellow"
        }
        elseif ($statusCode -like "*403*") {
            Write-LogMessage "Access forbidden - check app permissions in Azure AD" "INFO" "Yellow"
        }
        elseif ($statusCode -like "*404*") {
            Write-LogMessage "User not found or followed sites endpoint not available" "INFO" "Yellow"
        }
        
        return $null
    }
}

# Function to format and display results
function Format-SiteResults {
    param (
        [array]$Sites,
        [string]$Format,
        [string]$OutputFile = "",
        [bool]$IncludeDetails = $false
    )

    if ($Sites.Count -eq 0) {
        Write-LogMessage "No followed sites found for the user." "INFO" "Yellow"
        return
    }

    # Prepare the data
    if ($IncludeDetails) {
        $siteData = $Sites | Select-Object @{Name="Site Name"; Expression={$_.displayName}}, 
                                          @{Name="Site URL"; Expression={$_.webUrl}}, 
                                          @{Name="Site ID"; Expression={$_.id}},
                                          @{Name="Description"; Expression={$_.description}},
                                          @{Name="Last Modified"; Expression={if ($_.lastModifiedDateTime) { [DateTime]::Parse($_.lastModifiedDateTime).ToString("yyyy-MM-dd HH:mm:ss") } else { "N/A" }}},
                                          @{Name="Created"; Expression={if ($_.createdDateTime) { [DateTime]::Parse($_.createdDateTime).ToString("yyyy-MM-dd HH:mm:ss") } else { "N/A" }}}
    }
    else {
        $siteData = $Sites | Select-Object @{Name="Site Name"; Expression={$_.displayName}}, 
                                          @{Name="Site URL"; Expression={$_.webUrl}}, 
                                          @{Name="Site ID"; Expression={$_.id}}
    }

    switch ($Format) {
        "Table" {
            Write-LogMessage "`nFollowed SharePoint Sites:" "INFO" "Cyan"
            Write-LogMessage "=========================" "INFO" "Cyan"
            $siteData | Format-Table -AutoSize
        }
        "List" {
            Write-LogMessage "`nFollowed SharePoint Sites:" "INFO" "Cyan"
            Write-LogMessage "=========================" "INFO" "Cyan"
            $siteData | Format-List
        }
        "CSV" {
            if ($OutputFile) {
                $siteData | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
                Write-LogMessage "Results exported to CSV: $OutputFile" "SUCCESS" "Green"
            }
            else {
                $siteData | ConvertTo-Csv -NoTypeInformation
            }
        }
        "JSON" {
            $jsonOutput = $siteData | ConvertTo-Json -Depth 3
            if ($OutputFile) {
                $jsonOutput | Out-File -FilePath $OutputFile -Encoding UTF8
                Write-LogMessage "Results exported to JSON: $OutputFile" "SUCCESS" "Green"
            }
            else {
                Write-LogMessage "`nFollowed SharePoint Sites (JSON):" "INFO" "Cyan"
                Write-LogMessage "=================================" "INFO" "Cyan"
                $jsonOutput
            }
        }
    }

    # Display summary
    Write-LogMessage "`nSummary:" "INFO" "Cyan"
    Write-LogMessage "========" "INFO" "Cyan"
    Write-LogMessage "Total followed sites: $($Sites.Count)" "INFO" "White"
    Write-LogMessage "User: $UserId" "INFO" "White"
    
    if ($OutputFile) {
        Write-LogMessage "Output file: $OutputFile" "INFO" "White"
    }
}

# Main script execution
Write-LogMessage "Starting SharePoint followed sites retrieval..." "INFO" "Cyan"
Write-LogMessage "Script Parameters:" "INFO" "Cyan"
Write-LogMessage "  - TenantID: $TenantID" "INFO" "Gray"
Write-LogMessage "  - ApplicationId: $ApplicationId" "INFO" "Gray"
Write-LogMessage "  - UserId: $UserId" "INFO" "Gray"
Write-LogMessage "  - OutputFormat: $OutputFormat" "INFO" "Gray"
Write-LogMessage "  - IncludeDetails: $IncludeDetails" "INFO" "Gray"
if ($OutputFile) { Write-LogMessage "  - OutputFile: $OutputFile" "INFO" "Gray" }

# Get authentication token
Write-LogMessage "Obtaining authentication token..." "INFO" "Cyan"
$token = Get-AuthToken
Write-LogMessage "Authentication token obtained successfully" "SUCCESS" "Green"

# Test API connectivity and permissions
$connectivityTest = Test-APIConnectivity -Token $token
if (-not $connectivityTest) {
    Write-LogMessage "API connectivity test failed. Please check your app permissions." "ERROR" "Red"
    Write-LogMessage "Required permissions: Sites.Read.All or Sites.ReadWrite.All" "INFO" "Yellow"
    exit 1
}

# Validate user exists
$userInfo = Test-UserExists -UserId $UserId -Token $token
if (-not $userInfo) {
    Write-LogMessage "Cannot proceed without valid user information." "ERROR" "Red"
    exit 1
}

# Get followed sites
$followedSites = Get-UserFollowedSites -UserId $UserId -Token $token -IncludeDetails $IncludeDetails

if ($null -eq $followedSites) {
    Write-LogMessage "Failed to retrieve followed sites. Exiting." "ERROR" "Red"
    exit 1
}

# Format and display results
Format-SiteResults -Sites $followedSites -Format $OutputFormat -OutputFile $OutputFile -IncludeDetails $IncludeDetails

Write-LogMessage "`nOperation completed successfully!" "SUCCESS" "Green"
