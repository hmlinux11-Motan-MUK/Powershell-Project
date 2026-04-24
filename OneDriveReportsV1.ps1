<#
.SYNOPSIS
    Export ONLY explicit (unique) permissions for SharePoint Online / OneDrive files and folders.

.DESCRIPTION
    This script reads site URLs from a text file, connects to each site using app-only certificate authentication,
    scans document libraries, and exports ONLY files/folders with unique permissions.

    Output worksheets:
    1. ExplicitPermissions
       - One row per file/folder per user/group permission
    2. SharingLinksOnly
       - One row per file/folder per sharing link (optional)
    3. Summary
       - Per site summary

    This script is READ-ONLY.
    It does NOT modify permissions, files, folders, or SharePoint content.

.NOTES
    Required modules:
    - PnP.PowerShell
    - ImportExcel
#>

# =================================================================================================
# USER CONFIGURATION
# =================================================================================================

# --- Tenant / App Details ---
$appID  = "a88375e8-7d6c-4478-88db-327c31c476df"
$tenant = "a1b70a0b-a9f5-43a5-8d86-f0ecb1208eb0"

# --- Certificate Authentication (PFX file method) ---
$certificatePath   = "M:\PSproject\tenantassesmentcert.pfx"
$certPlainPassword = "@Cert!"     # keep blank if PFX has no password
$certPassword      = ConvertTo-SecureString $certPlainPassword -AsPlainText -Force

# --- Input / Output ---
$inputFilePath = "M:\PSproject\Testing.txt"
$outputFolder  = "C:\TEMP\mpd"

# --- Performance / Behavior ---
$pageSize = 2000
$IncludeSharingLinks = $false     # TRUE = include sharing links, FALSE = faster
$ProgressUpdateEvery = 100        # show progress after every N items

# =================================================================================================
# END USER CONFIGURATION
# =================================================================================================

# ---------------------------------
# Ensure output folder exists
# ---------------------------------
if (!(Test-Path $outputFolder)) {
    New-Item -Path $outputFolder -ItemType Directory -Force | Out-Null
}

# ---------------------------------
# Load required modules
# ---------------------------------
$requiredModules = @("PnP.PowerShell", "ImportExcel")
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Installing module $module ..." -ForegroundColor Yellow
        Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $module -Force
}

# ---------------------------------
# Prepare files
# ---------------------------------
$timeStamp      = Get-Date -Format "yyyyMMdd_HHmmss"
$logFilePath    = Join-Path $outputFolder "ExplicitPermissions_$timeStamp.log"
$outputFilePath = Join-Path $outputFolder "ExplicitPermissions_$timeStamp.xlsx"

# ---------------------------------
# Result collections
# ---------------------------------
$global:permissionResults  = New-Object System.Collections.Generic.List[object]
$global:sharingLinkResults = New-Object System.Collections.Generic.List[object]
$global:summaryResults     = New-Object System.Collections.Generic.List[object]

# ---------------------------------
# Ignore system/internal folders/libraries
# ---------------------------------
$ignoreFolders = @(
    "_catalogs",
    "appdata",
    "forms",
    "Form Templates",
    "Site Assets",
    "Style Library",
    "Composed Looks",
    "Converted Forms",
    "_cts",
    "_private",
    "_vti_pvt",
    "Sharing Links",
    "Social",
    "User Information List",
    "appfiles",
    "Preservation Hold Library",
    "List Template Gallery",
    "Master Page Gallery",
    "Solution Gallery",
    "Theme Gallery",
    "Maintenance Log Library"
)

# ---------------------------------
# Logging function
# ---------------------------------
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )

    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Level] $Message"
    Add-Content -Path $logFilePath -Value $line

    switch ($Level) {
        "ERROR"   { Write-Host $Message -ForegroundColor Red }
        "WARNING" { Write-Host $Message -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $Message -ForegroundColor Green }
        default   { Write-Host $Message -ForegroundColor Cyan }
    }
}

# ---------------------------------
# Retry wrapper
# ---------------------------------
function Invoke-WithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [int]$MaxRetries = 5,
        [int]$DelaySeconds = 5
    )

    $attempt = 0
    while ($attempt -lt $MaxRetries) {
        try {
            return & $ScriptBlock
        }
        catch {
            $attempt++
            if ($attempt -ge $MaxRetries) {
                throw
            }

            Write-Log "Retry $attempt/$MaxRetries after error: $($_.Exception.Message)" "WARNING"
            Start-Sleep -Seconds $DelaySeconds
            $DelaySeconds = [Math]::Min(($DelaySeconds * 2), 60)
        }
    }
}

# ---------------------------------
# Read site URLs from text file
# ---------------------------------
function Read-SiteURLs {
    param([string]$Path)

    if (!(Test-Path $Path)) {
        throw "Input file not found: $Path"
    }

    return Get-Content -Path $Path | Where-Object { $_ -and $_.Trim() -ne "" }
}

# ---------------------------------
# Connect to SharePoint site
# ---------------------------------
function Connect-SharePointSite {
    param([string]$SiteUrl)

    try {
        Invoke-WithRetry -ScriptBlock {
            Connect-PnPOnline `
                -Url $SiteUrl `
                -ClientId $appID `
                -Tenant $tenant `
                -CertificatePath $certificatePath `
                -CertificatePassword $certPassword
        }

        $web = Get-PnPWeb -ErrorAction Stop
        Write-Log "Connected to site: $($web.Url)" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Connection failed for $SiteUrl : $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# ---------------------------------
# Check ignored paths
# ---------------------------------
function Test-IsIgnoredPath {
    param([string]$ItemPath)

    foreach ($ignore in $ignoreFolders) {
        if ($ItemPath -like "*/$ignore/*" -or $ItemPath -match "/$([regex]::Escape($ignore))$") {
            return $true
        }
    }
    return $false
}

# ---------------------------------
# Principal type
# ---------------------------------
function Get-PrincipalTypeName {
    param($Member)

    try {
        if ($null -ne $Member.PrincipalType) {
            return $Member.PrincipalType.ToString()
        }
        return "Unknown"
    }
    catch {
        return "Unknown"
    }
}

# ---------------------------------
# Detect external user
# ---------------------------------
function Get-IsExternalUser {
    param($Member)

    try {
        $loginName = [string]$Member.LoginName
        $email     = [string]$Member.Email

        if ($loginName -match "#ext#" -or $loginName -match "urn:spo:guest" -or $email -match "#EXT#") {
            return $true
        }
        return $false
    }
    catch {
        return $false
    }
}

# ---------------------------------
# Safe email read
# ---------------------------------
function Get-PrincipalEmail {
    param($Member)

    try {
        if ($null -ne $Member.Email -and $Member.Email -ne "") {
            return $Member.Email
        }
        return ""
    }
    catch {
        return ""
    }
}

# ---------------------------------
# Get sharing links (optional)
# ---------------------------------
function Get-ItemSharingLinks {
    param(
        [string]$SiteURL,
        [string]$LibraryName,
        [string]$ItemType,
        [string]$ItemName,
        [string]$ItemPath
    )

    if (-not $IncludeSharingLinks) {
        return
    }

    try {
        $sharingLinks = @()

        if ($ItemType -eq "File") {
            $sharingLinks = Get-PnPFileSharingLink -Identity $ItemPath -ErrorAction SilentlyContinue
        }
        elseif ($ItemType -eq "Folder") {
            $sharingLinks = Get-PnPFolderSharingLink -Folder $ItemPath -ErrorAction SilentlyContinue
        }

        if ($sharingLinks) {
            foreach ($link in $sharingLinks) {
                $linkUrl = ""
                if ($link.Link) { $linkUrl = $link.Link }
                elseif ($link.Url) { $linkUrl = $link.Url }
                elseif ($link.WebUrl) { $linkUrl = $link.WebUrl }

                $linkType = ""
                if ($link.LinkKind) { $linkType = $link.LinkKind }
                elseif ($link.Type) { $linkType = $link.Type }

                $linkScope = ""
                if ($link.Scope) { $linkScope = $link.Scope }

                $linkPermissions = ""
                if ($link.Roles) { $linkPermissions = ($link.Roles -join ", ") }

                $linkExpiration = $null
                if ($link.ExpirationDateTime) { $linkExpiration = $link.ExpirationDateTime }
                elseif ($link.Expiration) { $linkExpiration = $link.Expiration }

                $global:sharingLinkResults.Add([PSCustomObject]@{
                    SiteURL          = $SiteURL
                    LibraryName      = $LibraryName
                    ItemType         = $ItemType
                    FileOrFolderName = $ItemName
                    FullPath         = $ItemPath
                    LinkType         = $linkType
                    LinkScope        = $linkScope
                    LinkPermissions  = $linkPermissions
                    LinkExpiration   = $linkExpiration
                    ShareLink        = $linkUrl
                })
            }
        }
    }
    catch {
        Write-Log "Could not retrieve sharing links for $ItemPath : $($_.Exception.Message)" "WARNING"
    }
}

# ---------------------------------
# Process one item
# Only explicit / unique permissions
# ---------------------------------
function Process-ListItem {
    param(
        $Item,
        [string]$SiteURL,
        [string]$LibraryName
    )

    try {
        # Load permissions
        Get-PnPProperty -ClientObject $Item -Property HasUniqueRoleAssignments, RoleAssignments | Out-Null

        # Skip inherited permissions
        if (-not $Item.HasUniqueRoleAssignments) {
            return $false
        }

        # File or Folder
        $fsObjType = $Item["FSObjType"]
        $itemType  = ""

        if ($fsObjType -eq 0) {
            $itemType = "File"
        }
        elseif ($fsObjType -eq 1) {
            $itemType = "Folder"
        }
        else {
            return $false
        }

        # Item details
        $itemName = $Item["FileLeafRef"]
        $itemPath = $Item["FileRef"]

        if ([string]::IsNullOrWhiteSpace($itemPath)) {
            return $false
        }

        if (Test-IsIgnoredPath -ItemPath $itemPath) {
            return $false
        }

        # Optional sharing links
        Get-ItemSharingLinks -SiteURL $SiteURL -LibraryName $LibraryName -ItemType $itemType -ItemName $itemName -ItemPath $itemPath

        $permissionAdded = $false

        # Each user/group = one row
        foreach ($roleAssignment in $Item.RoleAssignments) {
            try {
                Get-PnPProperty -ClientObject $roleAssignment -Property Member, RoleDefinitionBindings | Out-Null

                $member = $roleAssignment.Member
                $principalName = $member.Title
                if ([string]::IsNullOrWhiteSpace($principalName)) {
                    $principalName = $member.LoginName
                }

                $principalEmail = Get-PrincipalEmail -Member $member
                $principalType  = Get-PrincipalTypeName -Member $member
                $isExternal     = Get-IsExternalUser -Member $member

                # Ignore Limited Access only
                $roles = $roleAssignment.RoleDefinitionBindings |
                    Where-Object { $_.Name -ne "Limited Access" } |
                    ForEach-Object { $_.Name }

                if (-not $roles -or $roles.Count -eq 0) {
                    continue
                }

                $global:permissionResults.Add([PSCustomObject]@{
                    SiteURL          = $SiteURL
                    LibraryName      = $LibraryName
                    ItemType         = $itemType
                    FileOrFolderName = $itemName
                    FullPath         = $itemPath
                    UserName         = $principalName
                    UserEmail        = $principalEmail
                    PrincipalType    = $principalType
                    IsExternal       = $isExternal
                    Roles            = ($roles -join ", ")
                })

                $permissionAdded = $true
            }
            catch {
                Write-Log "Failed permission extraction on $itemPath : $($_.Exception.Message)" "WARNING"
            }
        }

        return $permissionAdded
    }
    catch {
        Write-Log "Failed item processing: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# ---------------------------------
# MAIN
# ---------------------------------
Write-Log "Script started"
Write-Log "Reading sites from: $inputFilePath"
Write-Log "Output file: $outputFilePath"
Write-Log "Include sharing links: $IncludeSharingLinks"

$siteUrls = Read-SiteURLs -Path $inputFilePath
Write-Log "Found $($siteUrls.Count) site(s)"

foreach ($siteUrl in $siteUrls) {
    $siteStart = Get-Date
    $siteExplicitItemCount = 0
    $sitePermissionRowCount = 0
    $siteLinkRowCount = 0

    Write-Log "Starting site: $siteUrl"

    if (-not (Connect-SharePointSite -SiteUrl $siteUrl)) {
        continue
    }

    try {
        # Only document libraries
        $libraries = Get-PnPList -Includes Title, Hidden, BaseType, ItemCount |
            Where-Object {
                $_.Hidden -eq $false -and
                $_.BaseType -eq "DocumentLibrary" -and
                $_.Title -notin $ignoreFolders
            }

        Write-Log "Found $($libraries.Count) document libraries in site"

        foreach ($library in $libraries) {
            Write-Log "Scanning library: $($library.Title) | Items: $($library.ItemCount)"

            if ($library.ItemCount -eq 0) {
                Write-Log "Skipping empty library: $($library.Title)"
                continue
            }

            # Get items
            $items = @(Get-PnPListItem -List $library.Title -PageSize $pageSize -Fields "FileLeafRef","FileRef","FSObjType")
            $totalItems = $items.Count
            $currentItemNumber = 0

            Write-Log "Retrieved $totalItems items from library: $($library.Title)"

            foreach ($item in $items) {
                $currentItemNumber++

                if ($currentItemNumber % $ProgressUpdateEvery -eq 0 -or $currentItemNumber -eq $totalItems) {
                    $percent = [math]::Round(($currentItemNumber / $totalItems) * 100, 2)
                    Write-Host "Library [$($library.Title)] : Processed $currentItemNumber / $totalItems items ($percent%)" -ForegroundColor Yellow
                    Write-Log "Progress - Library [$($library.Title)] : $currentItemNumber / $totalItems items ($percent%)"
                }

                $beforePermCount = $global:permissionResults.Count
                $beforeLinkCount = $global:sharingLinkResults.Count

                $wasExplicit = Process-ListItem -Item $item -SiteURL $siteUrl -LibraryName $library.Title

                if ($wasExplicit) {
                    $siteExplicitItemCount++
                }

                $sitePermissionRowCount += ($global:permissionResults.Count - $beforePermCount)
                $siteLinkRowCount += ($global:sharingLinkResults.Count - $beforeLinkCount)
            }

            Write-Log "Completed library: $($library.Title)" "SUCCESS"
        }

        Disconnect-PnPOnline -ErrorAction SilentlyContinue

        $duration = (Get-Date) - $siteStart
        $global:summaryResults.Add([PSCustomObject]@{
            SiteURL            = $siteUrl
            ExplicitItemsFound = $siteExplicitItemCount
            PermissionRows     = $sitePermissionRowCount
            SharingLinkRows    = $siteLinkRowCount
            Duration           = $duration.ToString()
        })

        Write-Log "Completed site: $siteUrl | Explicit Items: $siteExplicitItemCount | Permission Rows: $sitePermissionRowCount | Sharing Link Rows: $siteLinkRowCount" "SUCCESS"
    }
    catch {
        Write-Log "Site processing failed for $siteUrl : $($_.Exception.Message)" "ERROR"
    }
}

# ---------------------------------
# Export results to Excel
# ---------------------------------
if (($global:permissionResults.Count -gt 0) -or ($global:sharingLinkResults.Count -gt 0) -or ($global:summaryResults.Count -gt 0)) {

    if ($global:permissionResults.Count -gt 0) {
        $global:permissionResults |
            Sort-Object SiteURL, LibraryName, FullPath, UserName |
            Export-Excel -Path $outputFilePath `
                         -WorksheetName "ExplicitPermissions" `
                         -TableName "ExplicitPermissionsTbl" `
                         -TableStyle Medium6 `
                         -AutoSize `
                         -FreezeTopRow `
                         -BoldTopRow
    }

    if ($IncludeSharingLinks -and $global:sharingLinkResults.Count -gt 0) {
        $global:sharingLinkResults |
            Sort-Object SiteURL, LibraryName, FullPath |
            Export-Excel -Path $outputFilePath `
                         -WorksheetName "SharingLinksOnly" `
                         -TableName "SharingLinksTbl" `
                         -TableStyle Medium4 `
                         -AutoSize `
                         -FreezeTopRow `
                         -BoldTopRow `
                         -Append
    }

    if ($global:summaryResults.Count -gt 0) {
        $global:summaryResults |
            Export-Excel -Path $outputFilePath `
                         -WorksheetName "Summary" `
                         -TableName "SummaryTbl" `
                         -TableStyle Medium2 `
                         -AutoSize `
                         -FreezeTopRow `
                         -BoldTopRow `
                         -Append
    }

    Write-Log "Excel report created successfully: $outputFilePath" "SUCCESS"
}
else {
    Write-Log "No explicit permissions found." "WARNING"
}

Write-Log "Script finished"

if (Test-Path $outputFilePath) {
    try {
        Start-Process $outputFilePath
    }
    catch {
        Write-Log "Could not auto-open Excel file. Open manually: $outputFilePath" "WARNING"
    }
}