$siteUrl = "https://international260.sharepoint.com/sites/reports"
$libraryName = "Documents"

Write-Host "Connected..." -ForegroundColor Green

# Get all folders
$folders = Get-PnPListItem -List $libraryName -PageSize 1000 |
    Where-Object { $_.FileSystemObjectType -eq "Folder" }

foreach ($folder in $folders) {
    $folderUrl = $folder["FileRef"]
    Write-Host "`nChecking folder: $folderUrl" -ForegroundColor Cyan

    # Try to get all items inside folder (force scope)
    $items = Get-PnPListItem -List "Documents" -PageSize 1000 -Fields FileRef, FileDirRef

foreach ($item in $items) {
    if ($item["FileRef"] -like "*ProblemFolderName*") {
        try {
            Remove-PnPListItem -List "Documents" -Identity $item.Id -Force
            Write-Host "Force deleted list item: $($item['FileRef'])"
        }
        catch {
            Write-Host "Failed: $($item['FileRef'])"
        }
    }
}
}

Write-Host "`nCleanup pass complete."