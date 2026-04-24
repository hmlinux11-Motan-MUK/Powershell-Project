# =========================
# OneDrive Single-Site Test
# =========================

# ---- VARIABLES ----
$SiteUrl           = "https://mpdevelopmentllc-my.sharepoint.com/personal/l_mashava_mpd_ge"
$TenantId          = "a1b70a0b-a9f5-43a5-8d86-f0ecb1208eb0"
$ClientId          = "a88375e8-7d6c-4478-88db-327c31c476df"
$CertificatePath   = "M:\PSproject\tenantassesmentcert.pfx"
$CertPlainPassword = "@Cert!"
$OutputFolder      = "M:\PSproject\Result\OneDriveTest"

# ---- PREPARE ----
if (!(Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}

$CertPassword = ConvertTo-SecureString $CertPlainPassword -AsPlainText -Force
$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"

# ---- FUNCTION: EXPORT CSV IN EXCEL-READABLE UTF8 ----
function Export-ExcelFriendlyCsv {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Data,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $Data | Export-Csv -NoTypeInformation -Encoding UTF8BOM -Path $Path
    Write-Host "Exported: $Path" -ForegroundColor Green
}

Write-Host "Connecting to OneDrive site..." -ForegroundColor Cyan

try {
    Connect-PnPOnline `
        -Url $SiteUrl `
        -ClientId $ClientId `
        -Tenant $TenantId `
        -CertificatePath $CertificatePath `
        -CertificatePassword $CertPassword

    Write-Host "Connected successfully." -ForegroundColor Green
}
catch {
    Write-Host "Connection failed: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# ---- STEP 1: TEST DOCUMENTS LIBRARY ----
Write-Host "Testing Documents library..." -ForegroundColor Cyan

try {
    $items = Get-PnPListItem -List "Documents" -PageSize 20 -Fields "FileLeafRef","FileRef","FSObjType"
}
catch {
    Write-Host "Failed to read Documents library: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-PnPOnline
    exit
}

if (!$items -or $items.Count -eq 0) {
    Write-Host "No items found in Documents library." -ForegroundColor Yellow
    Disconnect-PnPOnline
    exit
}

Write-Host "Found $($items.Count) items." -ForegroundColor Green

# Show sample items
$sampleItems = $items | Select-Object `
    @{Name="ID";Expression={$_.Id}},
    @{Name="Name";Expression={$_.FieldValues["FileLeafRef"]}},
    @{Name="Path";Expression={$_.FieldValues["FileRef"]}},
    @{Name="Type";Expression={ if ($_.FieldValues["FSObjType"] -eq 1) { "Folder" } else { "File" } }}

$sampleItems | Format-Table -AutoSize
Export-ExcelFriendlyCsv -Data $sampleItems -Path (Join-Path $OutputFolder "SampleItems_$TimeStamp.csv")

# ---- STEP 2: PICK ONE FILE/FOLDER ----
# Prefer first actual file; if not found, use first item
$testItem = $items | Where-Object { $_.FieldValues["FSObjType"] -eq 0 } | Select-Object -First 1
if (!$testItem) {
    $testItem = $items | Select-Object -First 1
}

Write-Host ""
Write-Host "Testing item:" -ForegroundColor Cyan
Write-Host "ID   : $($testItem.Id)"
Write-Host "Name : $($testItem.FieldValues['FileLeafRef'])"
Write-Host "Path : $($testItem.FieldValues['FileRef'])"

# ---- STEP 3: READ ITEM PERMISSIONS ----
Write-Host ""
Write-Host "Reading permissions..." -ForegroundColor Cyan

try {
    Get-PnPProperty -ClientObject $testItem -Property HasUniqueRoleAssignments, RoleAssignments
}
catch {
    Write-Host "Failed to read RoleAssignments: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-PnPOnline
    exit
}

$permResults = @()

foreach ($roleAssignment in $testItem.RoleAssignments) {
    try {
        Get-PnPProperty -ClientObject $roleAssignment -Property Member, RoleDefinitionBindings

        $roles = $roleAssignment.RoleDefinitionBindings | ForEach-Object { $_.Name }

        $permResults += [PSCustomObject]@{
            SiteUrl            = $SiteUrl
            ItemId             = $testItem.Id
            ItemName           = $testItem.FieldValues["FileLeafRef"]
            ItemPath           = $testItem.FieldValues["FileRef"]
            ItemType           = if ($testItem.FieldValues["FSObjType"] -eq 1) { "Folder" } else { "File" }
            HasUniquePerms     = $testItem.HasUniqueRoleAssignments
            PrincipalName      = $roleAssignment.Member.Title
            PrincipalLoginName = $roleAssignment.Member.LoginName
            PrincipalType      = $roleAssignment.Member.PrincipalType
            Roles              = ($roles -join ", ")
        }
    }
    catch {
        $permResults += [PSCustomObject]@{
            SiteUrl            = $SiteUrl
            ItemId             = $testItem.Id
            ItemName           = $testItem.FieldValues["FileLeafRef"]
            ItemPath           = $testItem.FieldValues["FileRef"]
            ItemType           = if ($testItem.FieldValues["FSObjType"] -eq 1) { "Folder" } else { "File" }
            HasUniquePerms     = $testItem.HasUniqueRoleAssignments
            PrincipalName      = "ERROR"
            PrincipalLoginName = ""
            PrincipalType      = ""
            Roles              = $_.Exception.Message
        }
    }
}

if ($permResults.Count -gt 0) {
    Write-Host "Permissions read successfully." -ForegroundColor Green
    $permResults | Format-Table -AutoSize
    Export-ExcelFriendlyCsv -Data $permResults -Path (Join-Path $OutputFolder "Permissions_$TimeStamp.csv")
}
else {
    Write-Host "No permissions found on selected item." -ForegroundColor Yellow
}

# ---- STEP 4: TRY TO READ SHARING LINKS ----
Write-Host ""
Write-Host "Trying to read sharing links..." -ForegroundColor Cyan

$linkResults = @()

try {
    $sharingLinks = Get-PnPFileSharingLink -Identity $testItem.FieldValues["FileRef"] -ErrorAction Stop

    foreach ($link in $sharingLinks) {
        $linkResults += [PSCustomObject]@{
            SiteUrl          = $SiteUrl
            ItemName         = $testItem.FieldValues["FileLeafRef"]
            ItemPath         = $testItem.FieldValues["FileRef"]
            ItemType         = if ($testItem.FieldValues["FSObjType"] -eq 1) { "Folder" } else { "File" }
            LinkKind         = $link.LinkKind
            WebUrl           = $link.WebUrl
            PreventsDownload = $link.PreventsDownload
        }
    }

    if ($linkResults.Count -gt 0) {
        Write-Host "Sharing links found." -ForegroundColor Green
        $linkResults | Format-Table -AutoSize
        Export-ExcelFriendlyCsv -Data $linkResults -Path (Join-Path $OutputFolder "SharingLinks_$TimeStamp.csv")
    }
    else {
        Write-Host "No sharing links found for the selected item." -ForegroundColor Yellow
    }
}
catch {
    Write-Host "Could not read sharing links. This may be normal if the file is not shared or the cmdlet/version does not support it." -ForegroundColor Yellow
    Write-Host $_.Exception.Message -ForegroundColor DarkYellow
}

# ---- STEP 5: SUMMARY ----
Write-Host ""
Write-Host "===== TEST SUMMARY =====" -ForegroundColor Cyan
Write-Host "Site tested      : $SiteUrl"
Write-Host "Selected item    : $($testItem.FieldValues['FileLeafRef'])"
Write-Host "Selected path    : $($testItem.FieldValues['FileRef'])"
Write-Host "Unique perms     : $($testItem.HasUniqueRoleAssignments)"
Write-Host "Permissions rows : $($permResults.Count)"
Write-Host "Output folder    : $OutputFolder"

Disconnect-PnPOnline
Write-Host "Done." -ForegroundColor Green