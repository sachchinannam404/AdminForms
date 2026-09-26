<#
.SYNOPSIS
  Provisions Admin Forms lists, columns, versioning, and security groups on a SharePoint site.

.DESCRIPTION
  Requires PnP.PowerShell. Run as site owner.

.EXAMPLE
  Connect-PnPOnline -Url "https://contoso.sharepoint.com/sites/Admin" -Interactive
  .\scripts\Provision-AdminForms.ps1
#>

[CmdletBinding()]
param(
  [string]$SiteUrl
)

if ($SiteUrl) {
  Connect-PnPOnline -Url $SiteUrl -Interactive
}

function Ensure-List {
  param([string]$Title, [string]$Description, [switch]$Versioning)
  $list = Get-PnPList -Identity $Title -ErrorAction SilentlyContinue
  if (-not $list) {
    $list = New-PnPList -Title $Title -Template GenericList -OnQuickLaunch
    Write-Host "Created list: $Title" -ForegroundColor Green
  } else {
    Write-Host "List exists: $Title" -ForegroundColor Cyan
  }
  if ($Versioning) {
    Set-PnPList -Identity $Title -EnableVersioning $true
    Write-Host "  Versioning enabled" -ForegroundColor Gray
  }
  return $list
}

function Ensure-Field {
  param(
    [string]$ListTitle,
    [string]$InternalName,
    [string]$Type,
    [string[]]$Choices = @()
  )
  $existing = Get-PnPField -List $ListTitle -Identity $InternalName -ErrorAction SilentlyContinue
  if ($existing) {
    Write-Host "  Field exists: $InternalName" -ForegroundColor Gray
    return
  }
  switch ($Type) {
    'Text' { Add-PnPField -List $ListTitle -DisplayName $InternalName -InternalName $InternalName -Type Text -AddToDefaultView | Out-Null }
    'Note' { Add-PnPField -List $ListTitle -DisplayName $InternalName -InternalName $InternalName -Type Note -AddToDefaultView | Out-Null }
    'Number' { Add-PnPField -List $ListTitle -DisplayName $InternalName -InternalName $InternalName -Type Number -AddToDefaultView | Out-Null }
    'Currency' { Add-PnPField -List $ListTitle -DisplayName $InternalName -InternalName $InternalName -Type Currency -AddToDefaultView | Out-Null }
    'DateTime' { Add-PnPField -List $ListTitle -DisplayName $InternalName -InternalName $InternalName -Type DateTime -AddToDefaultView | Out-Null }
    'Choice' {
      Add-PnPField -List $ListTitle -DisplayName $InternalName -InternalName $InternalName -Type Choice -Choices $Choices -AddToDefaultView | Out-Null
    }
  }
  Write-Host "  Added field: $InternalName ($Type)" -ForegroundColor Green
}

# --- Parent list ---
Ensure-List -Title "Admin Requests" -Description "Parent admin requests" -Versioning

$reqTypes = @('Stationery','ITEquipment','Travel','Leave','Facilities','Procurement','General')
$statuses = @('Draft','Pending','Approved','Rejected','InProgress','Completed','Cancelled')
$priorities = @('Low','Medium','High','Urgent')

Ensure-Field -ListTitle "Admin Requests" -InternalName "Description" -Type Note
Ensure-Field -ListTitle "Admin Requests" -InternalName "RequestType" -Type Choice -Choices $reqTypes
Ensure-Field -ListTitle "Admin Requests" -InternalName "RequesterName" -Type Text
Ensure-Field -ListTitle "Admin Requests" -InternalName "RequesterEmail" -Type Text
Ensure-Field -ListTitle "Admin Requests" -InternalName "Department" -Type Text
Ensure-Field -ListTitle "Admin Requests" -InternalName "Status" -Type Choice -Choices $statuses
Ensure-Field -ListTitle "Admin Requests" -InternalName "Priority" -Type Choice -Choices $priorities
Ensure-Field -ListTitle "Admin Requests" -InternalName "TargetDeliveryDate" -Type DateTime
Ensure-Field -ListTitle "Admin Requests" -InternalName "ApprovedBy" -Type Text
Ensure-Field -ListTitle "Admin Requests" -InternalName "ApprovedDate" -Type DateTime
Ensure-Field -ListTitle "Admin Requests" -InternalName "RejectionReason" -Type Note
Ensure-Field -ListTitle "Admin Requests" -InternalName "TotalBudget" -Type Currency
Ensure-Field -ListTitle "Admin Requests" -InternalName "Comments" -Type Note
Ensure-Field -ListTitle "Admin Requests" -InternalName "DetailsJson" -Type Note

# --- Child lists ---
$childLists = @(
  'Stationery Items',
  'IT Equipment Items',
  'Travel Items',
  'Facilities Items',
  'Procurement Items'
)
$itemStatuses = @('Pending','Ordered','InStock','Delivered','Cancelled')

foreach ($child in $childLists) {
  Ensure-List -Title $child -Description "Line items: $child"
  Ensure-Field -ListTitle $child -InternalName "RequestId" -Type Text
  Ensure-Field -ListTitle $child -InternalName "ItemName" -Type Text
  Ensure-Field -ListTitle $child -InternalName "Category" -Type Text
  Ensure-Field -ListTitle $child -InternalName "Quantity" -Type Number
  Ensure-Field -ListTitle $child -InternalName "Unit" -Type Text
  Ensure-Field -ListTitle $child -InternalName "UnitPrice" -Type Currency
  Ensure-Field -ListTitle $child -InternalName "TotalPrice" -Type Currency
  Ensure-Field -ListTitle $child -InternalName "Description" -Type Note
  Ensure-Field -ListTitle $child -InternalName "Status" -Type Choice -Choices $itemStatuses
  Ensure-Field -ListTitle $child -InternalName "VendorName" -Type Text
  Ensure-Field -ListTitle $child -InternalName "VendorEmail" -Type Text
  Ensure-Field -ListTitle $child -InternalName "ExpectedDeliveryDate" -Type DateTime
  Ensure-Field -ListTitle $child -InternalName "ActualDeliveryDate" -Type DateTime
  Ensure-Field -ListTitle $child -InternalName "Remarks" -Type Note
  Ensure-Field -ListTitle $child -InternalName "ExtraJson" -Type Note
}

# --- Groups ---
function Ensure-Group {
  param([string]$Title, [string]$Description)
  $g = Get-PnPGroup -Identity $Title -ErrorAction SilentlyContinue
  if (-not $g) {
    New-PnPGroup -Title $Title -Description $Description | Out-Null
    Write-Host "Created group: $Title" -ForegroundColor Green
  } else {
    Write-Host "Group exists: $Title" -ForegroundColor Cyan
  }
}

Ensure-Group -Title "Admin Forms Approvers" -Description "Can approve Admin Forms requests"
Ensure-Group -Title "Admin Forms Admins" -Description "Full admin for Admin Forms"

Write-Host "`nDone. Add managers to 'Admin Forms Approvers' and owners to 'Admin Forms Admins'." -ForegroundColor Yellow
Write-Host "Add the Admin Forms web part to a page and optionally enable 'Ensure lists & groups on load'." -ForegroundColor Yellow
