# ============================================================
# Create-ICReportColumns.ps1
# Creates all required columns on the IC Project Reports list
# Requires: PnP.PowerShell module
#
# Install the module if needed (run as Administrator):
#   Install-Module -Name PnP.PowerShell -AllowClobber
#
# If you get execution policy errors, run:
#   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
#
# Run:
#   .\Create-ICReportColumns.ps1
# ============================================================

param(
    [string]$SiteUrl  = "https://international260.sharepoint.com/sites/ASTReports",
    [string]$ListName = "Reports",
    [string]$ClientId = "66d5787b-f69b-468a-af0b-6791dee76928",
    [string]$TenantId = "3e34eb62-b83a-4aa8-d0959d15e612"
)

# Import module and connect
Write-Host "`nConnecting to SharePoint..." -ForegroundColor Cyan
Import-Module -Name PnP.PowerShell -ErrorAction SilentlyContinue
if (-not (Get-Command Connect-PnPOnline -ErrorAction SilentlyContinue)) {
    Write-Host "Error: PnP.PowerShell module not found. Run: Install-Module -Name PnP.PowerShell -AllowClobber" -ForegroundColor Red
    exit 1
}
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $TenantId -Interactive
Write-Host "Connected." -ForegroundColor Green

# Helper: add a column if it doesn't already exist
function Add-ColumnIfMissing {
    param(
        [string]$InternalName,
        [string]$DisplayName,
        [string]$Type,
        [bool]$Required = $false,
        [string[]]$Choices = @()
    )

    $existing = Get-PnPField -List $ListName -Identity $InternalName -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  SKIP  $InternalName (already exists)" -ForegroundColor DarkGray
        return
    }

    try {
        switch ($Type) {
            "Text" {
                Add-PnPField -List $ListName -InternalName $InternalName -DisplayName $DisplayName -Type Text -Required:$Required | Out-Null
            }
            "Note" {
                Add-PnPField -List $ListName -InternalName $InternalName -DisplayName $DisplayName -Type Note -Required:$Required | Out-Null
            }
            "Number" {
                Add-PnPField -List $ListName -InternalName $InternalName -DisplayName $DisplayName -Type Number -Required:$Required | Out-Null
            }
            "Currency" {
                Add-PnPField -List $ListName -InternalName $InternalName -DisplayName $DisplayName -Type Currency -Required:$Required | Out-Null
            }
            "DateTime" {
                Add-PnPField -List $ListName -InternalName $InternalName -DisplayName $DisplayName -Type DateTime -Required:$Required | Out-Null
            }
            "Boolean" {
                Add-PnPField -List $ListName -InternalName $InternalName -DisplayName $DisplayName -Type Boolean -Required:$Required | Out-Null
            }
            "Choice" {
                $field = Add-PnPField -List $ListName -InternalName $InternalName -DisplayName $DisplayName -Type Choice -Required:$Required
                $field.Choices = $Choices
                $field.Update()
                Invoke-PnPQuery
            }
        }
        Write-Host "  OK    $InternalName" -ForegroundColor Green
    }
    catch {
        Write-Host "  ERROR $InternalName - $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Create columns
Write-Host "`nCreating columns on list: '$ListName'" -ForegroundColor Cyan

# Section 1 - Project Info
Add-ColumnIfMissing -InternalName "City" -DisplayName "City / Cities" -Type "Text"
Add-ColumnIfMissing -InternalName "Country" -DisplayName "Country" -Type "Text"
Add-ColumnIfMissing -InternalName "Area" -DisplayName "Area" -Type "Choice" -Choices @("Latin America","Western Europe","Eastern Europe","Africa","North Africa / Middle East (N.A.M.E.)","South Asia","Asia","Oceania")
Add-ColumnIfMissing -InternalName "ProjectDateFrom" -DisplayName "Project Date From" -Type "DateTime"
Add-ColumnIfMissing -InternalName "ProjectDateTo" -DisplayName "Project Date To" -Type "DateTime"
Add-ColumnIfMissing -InternalName "Introduction" -DisplayName "Introduction" -Type "Note"

# Section 2 - Statistics
Add-ColumnIfMissing -InternalName "ChurchesParticipated" -DisplayName "Churches Participated" -Type "Number"
Add-ColumnIfMissing -InternalName "NationalParticipants" -DisplayName "National Participants" -Type "Number"
Add-ColumnIfMissing -InternalName "USAParticipants" -DisplayName "USA Participants" -Type "Number"
Add-ColumnIfMissing -InternalName "OtherCountriesParticipants" -DisplayName "Other Countries Participants" -Type "Number"
Add-ColumnIfMissing -InternalName "PeopleHeardGospel" -DisplayName "People Who Heard the Gospel" -Type "Number"
Add-ColumnIfMissing -InternalName "ProfessionsOfFaith" -DisplayName "Professions of Faith" -Type "Number"
Add-ColumnIfMissing -InternalName "Rededications" -DisplayName "Rededications to Christ" -Type "Number"
Add-ColumnIfMissing -InternalName "Baptisms" -DisplayName "Baptisms" -Type "Number"
Add-ColumnIfMissing -InternalName "NewChurchesPlanted" -DisplayName "New Churches Planted" -Type "Number"

# Section 3 - Testimonies
Add-ColumnIfMissing -InternalName "Testimonies" -DisplayName "Testimonies" -Type "Note"

# Section 5 - Financial Report
Add-ColumnIfMissing -InternalName "TotalFundsSent" -DisplayName "Total Funds Sent" -Type "Currency"
Add-ColumnIfMissing -InternalName "SpentOnMaterials" -DisplayName "Spent on Materials" -Type "Currency"
Add-ColumnIfMissing -InternalName "TicketsCost" -DisplayName "Tickets Cost" -Type "Currency"
Add-ColumnIfMissing -InternalName "FuelCost" -DisplayName "Fuel / Taxi Cost" -Type "Currency"
Add-ColumnIfMissing -InternalName "AccommodationCost" -DisplayName "Accommodation Cost" -Type "Currency"
Add-ColumnIfMissing -InternalName "FoodCost" -DisplayName "Food Cost" -Type "Currency"
Add-ColumnIfMissing -InternalName "FinancialHelpParticipants" -DisplayName "Financial Help to Participants" -Type "Currency"
Add-ColumnIfMissing -InternalName "NumParticipantsHelp" -DisplayName "Participants Receiving Help" -Type "Number"
Add-ColumnIfMissing -InternalName "RalliesExpenses" -DisplayName "Rallies Expenses" -Type "Currency"
Add-ColumnIfMissing -InternalName "RalliesDescription" -DisplayName "Rallies Description" -Type "Text"
Add-ColumnIfMissing -InternalName "AdditionalExpenses" -DisplayName "Additional Expenses" -Type "Currency"
Add-ColumnIfMissing -InternalName "AdditionalNeedDescription" -DisplayName "Additional Need Description" -Type "Text"

# Coordinator Info
Add-ColumnIfMissing -InternalName "CoordinatorName" -DisplayName "Coordinator Name" -Type "Text"
Add-ColumnIfMissing -InternalName "CoordinatorEmail" -DisplayName "Coordinator Email" -Type "Text"
Add-ColumnIfMissing -InternalName "SubmittedAt" -DisplayName "Submitted At" -Type "DateTime"

# Photo folder reference
Add-ColumnIfMissing -InternalName "PhotoFolderUrl" -DisplayName "Photo Folder URL" -Type "Text"
Add-ColumnIfMissing -InternalName "PhotoFolderServerRelativePath" -DisplayName "Photo Folder Server Relative Path" -Type "Text"
Add-ColumnIfMissing -InternalName "PhotoDriveId" -DisplayName "Photo Drive ID" -Type "Text"
Add-ColumnIfMissing -InternalName "PhotoFolderItemId" -DisplayName "Photo Folder Item ID" -Type "Text"

# Section 6 - Recycle Bin
Add-ColumnIfMissing -InternalName "IsRecycled" -DisplayName "Is Recycled" -Type "Boolean"

# Done
Write-Host "`nDone. All columns processed." -ForegroundColor Green
Write-Host "You can verify at: $SiteUrl/Lists/$([uri]::EscapeDataString($ListName))"