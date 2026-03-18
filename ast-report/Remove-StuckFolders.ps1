# ============================================================
# Remove-StuckFolders.ps1
# Deletes everything inside a SharePoint Documents library.
#
# Run:
#   .\Remove-StuckFolders.ps1
# ============================================================

param(
    [string]$SiteUrl     = "https://international260.sharepoint.com/sites/reports",
    [string]$ClientId    = "66d5787b-f69b-468a-af0b-6791dee76928",
    [string]$TenantId    = "3e34eb62-b83a-4aa8-8ba8-d0959d15e612",
    [string]$LibraryPath = "Documents"
)

Write-Host "`nConnecting to SharePoint..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $TenantId -Interactive
Write-Host "Connected.`n" -ForegroundColor Green

# List everything at the root of the library
Write-Host "Contents of '$LibraryPath':" -ForegroundColor Cyan
$folders = Get-PnPFolderItem -FolderSiteRelativeUrl $LibraryPath -ItemType Folder -ErrorAction SilentlyContinue
$files   = Get-PnPFolderItem -FolderSiteRelativeUrl $LibraryPath -ItemType File   -ErrorAction SilentlyContinue

$folders | ForEach-Object { Write-Host "  [DIR]  $($_.Name)" }
$files   | ForEach-Object { Write-Host "  [FILE] $($_.Name)" }

if (!$folders -and !$files) {
    Write-Host "  (library is already empty)" -ForegroundColor DarkGray
    exit
}

Write-Host ""
$confirm = Read-Host "Delete ALL of the above from '$LibraryPath'? This cannot be undone. (y/n)"
if ($confirm -ne "y") {
    Write-Host "Aborted." -ForegroundColor Yellow
    exit
}

Write-Host ""

# Delete all files at root level
foreach ($file in $files) {
    try {
        Remove-PnPFile -ServerRelativeUrl $file.ServerRelativeUrl -Force -Recycle -ErrorAction Stop
        Write-Host "  Deleted file: $($file.Name)" -ForegroundColor Green
    } catch {
        Write-Host "  ERROR deleting file $($file.Name): $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Recursive folder delete (handles ghost files inside)
function Remove-FolderRecursive {
    param([string]$FolderSiteRelUrl, [string]$FolderName)

    # Delete child files first
    $childFiles = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderSiteRelUrl -ItemType File -ErrorAction SilentlyContinue
    foreach ($f in $childFiles) {
        try {
            Remove-PnPFile -ServerRelativeUrl $f.ServerRelativeUrl -Force -Recycle -ErrorAction Stop
            Write-Host "    Deleted file: $($f.Name)" -ForegroundColor DarkGray
        } catch {
            Write-Host "    ERROR deleting file $($f.Name): $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Recurse into subfolders
    $childFolders = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderSiteRelUrl -ItemType Folder -ErrorAction SilentlyContinue
    foreach ($sf in $childFolders) {
        Remove-FolderRecursive -FolderSiteRelUrl "$FolderSiteRelUrl/$($sf.Name)" -FolderName $sf.Name
    }

    # Delete the folder itself
    $parentPath = $FolderSiteRelUrl -replace "/[^/]+$", "" -replace "^/sites/reports/", ""
    try {
        Remove-PnPFolder -Name $FolderName -Folder $parentPath -Force -Recycle -ErrorAction Stop
        Write-Host "  Deleted folder: $FolderName" -ForegroundColor Green
    } catch {
        Write-Host "  ERROR deleting folder $FolderName : $($_.Exception.Message)" -ForegroundColor Red
    }
}

foreach ($folder in $folders) {
    Write-Host "  Processing folder: $($folder.Name)" -ForegroundColor Cyan
    Remove-FolderRecursive -FolderSiteRelUrl "$LibraryPath/$($folder.Name)" -FolderName $folder.Name
}

Write-Host "`nDone." -ForegroundColor Green
