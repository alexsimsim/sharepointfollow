<#
.SYNOPSIS
Example usage of Get-UserFollowedSites.ps1

.DESCRIPTION
This file provides examples of how to use the Get-UserFollowedSites.ps1 script
to retrieve SharePoint sites that users are following.
#>

# Example 1: Basic usage with default table output
# This will show followed sites in a formatted table
Write-Host "Example 1: Basic table output" -ForegroundColor Green
.\Get-UserFollowedSites.ps1 -UserId "user@contoso.com"

Write-Host "`n" + "="*50 + "`n"

# Example 2: Detailed information with list format
# This includes additional site details and shows them in list format
Write-Host "Example 2: Detailed information in list format" -ForegroundColor Green
.\Get-UserFollowedSites.ps1 `
    -UserId "user@contoso.com" `
    -OutputFormat "List" `
    -IncludeDetails

Write-Host "`n" + "="*50 + "`n"

# Example 3: Export to CSV file
# This exports the results to a CSV file for further analysis
Write-Host "Example 3: Export to CSV" -ForegroundColor Green
.\Get-UserFollowedSites.ps1 `
    -TenantID "contoso.onmicrosoft.com" `
    -ApplicationId "your-application-id" `
    -ApplicationSecret "your-application-secret" `
    -UserId "user@contoso.com" `
    -OutputFormat "CSV" `
    -OutputFile "user-followed-sites.csv" `
    -IncludeDetails

Write-Host "`n" + "="*50 + "`n"

# Example 4: JSON output for programmatic use
# This outputs JSON format which can be used by other scripts or applications
Write-Host "Example 4: JSON output" -ForegroundColor Green
.\Get-UserFollowedSites.ps1 `
    -UserId "user@contoso.com" `
    -OutputFormat "JSON" `
    -IncludeDetails

Write-Host "`n" + "="*50 + "`n"

# Example 5: Multiple users in a loop
# This shows how to get followed sites for multiple users
Write-Host "Example 5: Multiple users processing" -ForegroundColor Green
$users = @("user1@contoso.com", "user2@contoso.com", "user3@contoso.com")

foreach ($user in $users) {
    Write-Host "Processing user: $user" -ForegroundColor Yellow
    .\Get-UserFollowedSites.ps1 `
        -UserId $user `
        -OutputFormat "Table"
    Write-Host ""
}

Write-Host "`n" + "="*50 + "`n"

# Example 6: Export each user's followed sites to separate files
# This creates individual CSV files for each user
Write-Host "Example 6: Individual CSV files for multiple users" -ForegroundColor Green
$users = @("user1@contoso.com", "user2@contoso.com", "user3@contoso.com")

foreach ($user in $users) {
    $fileName = "followed-sites-$(($user -split '@')[0]).csv"
    Write-Host "Exporting followed sites for $user to $fileName" -ForegroundColor Yellow
    
    .\Get-UserFollowedSites.ps1 `
        -UserId $user `
        -OutputFormat "CSV" `
        -OutputFile $fileName `
        -IncludeDetails
}
