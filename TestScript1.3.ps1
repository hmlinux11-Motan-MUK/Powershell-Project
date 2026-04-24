# =========================
# OneDrive Single-Site Test
# Read ONLY unique permissions
# =========================

# ---- VARIABLES ----
$SiteUrl            = "https://mpdevelopmentllc-my.sharepoint.com/personal/l_mashava_mpd_ge"
$TenantId           = "a1b70a0b-a9f5-43a5-8d86-f0ecb1208eb0"
$ClientId           = "a88375e8-7d6c-4478-88db-327c31c476df"
$CertificatePath    = "M:\PSproject\tenantassesmentcert.pfx"
$CertPlainPassword  = "@Cert!"
$OutputFolder       = "M:\PSproject\Result\OneDriveTest"

# ---- PREPARE ----
if (!(Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}

$CertPassword = ConvertTo-SecureString $CertPlainPassword -AsPlainText -Force
$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"

Write-Host "Connecting to OneDrive site..." -ForegroundColor Cyan

try {
    Connect-PnPOnline `
        -Url $SiteUrl `
        -ClientId $ClientId `
        -Tenant $TenantId `
        -CertificatePath $CertificatePath `
        -CertificatePassword $CertPassword

    Write-Host "Connection successful." -ForegroundColor Green
}
catch {
    Write-Host "Connection failed: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# ---- READ ITEMS ----
Write-Host "Reading items from Documents library..." -ForegroundColor Cyan

try {
    $items = Get-PnPListItem -List "Documents" -PageSize 200 -Fields "FileLeafRef","FileRef","FSObjType"
}
catch {
    Write-Host "Failed to read the Documents library: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-PnPOnline
    exit
}

if (!$items -or $items.Count -eq 0) {
    Write-Host "No items found." -ForegroundColor Yellow
    Disconnect-PnPOnline
    exit
}

Write-Host "Total items found: $($items.Count)" -ForegroundColor Green

$results = @()
$counter = 0

foreach ($item in $items) {
    $counter++
    Write-Host "Processing item $counter of $($items.Count): $($item.FieldValues['FileLeafRef'])" -ForegroundColor DarkCyan

    try {
        # Load only the unique permission flag first
        Get-PnPProperty -ClientObject $item -Property HasUniqueRoleAssignments

        # Skip inherited permissions
        if (-not $item.HasUniqueRoleAssignments) {
            continue
        }

        # Only if unique permissions exist, load role assignments
        Get-PnPProperty -ClientObject $item -Property RoleAssignments

        foreach ($roleAssignment in $item.RoleAssignments) {
            try {
                Get-PnPProperty -ClientObject $roleAssignment -Property Member, RoleDefinitionBindings

                $roles = $roleAssignment.RoleDefinitionBindings | ForEach-Object { $_.Name }

                $results += [PSCustomObject]@{
                    SiteUrl              = $SiteUrl
                    ItemId               = $item.Id
                    ItemName             = $item.FieldValues["FileLeafRef"]
                    ItemPath             = $item.FieldValues["FileRef"]
                    ItemType             = if ($item.FieldValues["FSObjType"] -eq 1) { "Folder" } else { "File" }
                    HasUniquePermissions = $item.HasUniqueRoleAssignments
                    PrincipalName        = $roleAssignment.Member.Title
                    PrincipalLoginName   = $roleAssignment.Member.LoginName
                    PrincipalType        = $roleAssignment.Member.PrincipalType
                    PermissionRoles      = ($roles -join ", ")
                }
            }
            catch {
                $results += [PSCustomObject]@{
                    SiteUrl              = $SiteUrl
                    ItemId               = $item.Id
                    ItemName             = $item.FieldValues["FileLeafRef"]
                    ItemPath             = $item.FieldValues["FileRef"]
                    ItemType             = if ($item.FieldValues["FSObjType"] -eq 1) { "Folder" } else { "File" }
                    HasUniquePermissions = $item.HasUniqueRoleAssignments
                    PrincipalName        = "ERROR"
                    PrincipalLoginName   = ""
                    PrincipalType        = ""
                    PermissionRoles      = $_.Exception.Message
                }
            }
        }
    }
    catch {
        Write-Host "Failed to process item '$($item.FieldValues['FileLeafRef'])': $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# ---- EXPORT ----
if ($results.Count -gt 0) {
    $outFile = Join-Path $OutputFolder "UniquePermissionsOnly_$TimeStamp.csv"
    $results | Export-Csv -NoTypeInformation -Encoding utf8BOM -Path $outFile

    Write-Host ""
    Write-Host "Unique permission report created successfully." -ForegroundColor Green
    Write-Host "Rows exported: $($results.Count)" -ForegroundColor Green
    Write-Host "Output file  : $outFile" -ForegroundColor Green
}
else {
    Write-Host "No items with unique permissions were found." -ForegroundColor Yellow
}

Disconnect-PnPOnline
Write-Host "Done." -ForegroundColor Green