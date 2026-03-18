# ============================================================
# Setup-ASTReportsSite.ps1
# Creates the new AST Reports site with required lists and libraries
#
# Run:
#   .\Setup-ASTReportsSite.ps1
# ============================================================

param(
    [string]$SiteUrl  = "https://international260.sharepoint.com/sites/ASTReports",
    [string]$ClientId = "66d5787b-f69b-468a-af0b-6791dee76928",
    [string]$TenantId = "3e34eb62-b83a-4aa8-d0959d15e612"
)

Write-Host "`nConnecting to SharePoint..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $TenantId -Interactive
Write-Host "Connected.`n" -ForegroundColor Green

# Step 1: Create "Reports" list
Write-Host "Creating 'Reports' list..." -ForegroundColor Cyan
try {
    New-PnPList -Title "Reports" -Template DocumentLibrary -Url "Reports" -ErrorAction SilentlyContinue
    Write-Host "Created 'Reports' list" -ForegroundColor Green
} catch {
    if ($_.Exception.Message -like "*already exists*") {
        Write-Host "'Reports' list already exists" -ForegroundColor Yellow
    } else {
        Write-Host "Failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Step 2: Create "Report Photos" document library
Write-Host "`nCreating 'Report Photos' document library..." -ForegroundColor Cyan
try {
    New-PnPList -Title "Report Photos" -Template DocumentLibrary -Url "ReportPhotos" -ErrorAction SilentlyContinue
    Write-Host "Created 'Report Photos' library" -ForegroundColor Green
} catch {
    if ($_.Exception.Message -like "*already exists*") {
        Write-Host "'Report Photos' library already exists" -ForegroundColor Yellow
    } else {
        Write-Host "Failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Step 3: Create Report Photos folder
Write-Host "`nCreating 'Report Photos' folder..." -ForegroundColor Cyan
try {
    New-PnPFolder -List "Report Photos" -Name "Report Photos" -ErrorAction SilentlyContinue
    Write-Host "Created 'Report Photos' folder" -ForegroundColor Green
} catch {
    if ($_.Exception.Message -like "*already exists*") {
        Write-Host "'Report Photos' folder already exists" -ForegroundColor Yellow
    } else {
        Write-Host "Failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Show what was created
Write-Host "`nSite contents:" -ForegroundColor Cyan
$lists = Get-PnPList
foreach ($l in $lists) {
    Write-Host "  - $($l.Title) (URL: $($l.DefaultViewUrl))"
}

Write-Host "`nDone!" -ForegroundColor Green
Write-Host "`nNEXT STEPS:" -ForegroundColor Yellow
Write-Host "1. Update Worker secrets:" -ForegroundColor White
Write-Host "   wrangler secret put SHAREPOINT_SITE_URL" -ForegroundColor DarkGray
Write-Host "   (enter: https://international260.sharepoint.com/sites/ASTReports)" -ForegroundColor DarkGray
Write-Host "2. Run Create-ICReportColumns.ps1 on the new site" -ForegroundColor White