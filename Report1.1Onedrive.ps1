# =================================================================================================
# OneDrive Unique Permissions Migration Report
# Reports ONLY unique permissions on files/folders and classifies migration handling.
# =================================================================================================

# ---------------------------------
# USER CONFIGURATION
# ---------------------------------
$appID             = "a88375e8-7d6c-4478-88db-327c31c476df"   # Entra App ID
$tenant            = "a1b70a0b-a9f5-43a5-8d86-f0ecb1208eb0"   # Tenant ID or tenant.onmicrosoft.com
$certificatePath   = "M:\PSproject\tenantassesmentcert.pfx"
$certPlainPassword = "@Cert!"   # leave empty if cert has no password
$inputFilePath     = "M:\PSproject\Testing.txt"   # one OneDrive site URL per line
$outputRoot        = "C:\TEMP\mpd"

$batchSize         = 200
$maxItemsPerSheet  = 5000

# ---------------------------------
# PREPARE
# ---------------------------------
if (!(Test-Path $outputRoot)) {
    New-Item -Path $outputRoot -ItemType Directory -Force | Out-Null
}

if (!(Test-Path $inputFilePath)) {
    throw "Input file not found: $inputFilePath"
}

$certPassword = ConvertTo-SecureString $certPlainPassword -AsPlainText -Force
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFilePath = Join-Path $outputRoot "OneDrive_UniquePermissions_$timestamp.log"
$outputFilePath = Join-Path $outputRoot "OneDrive_UniquePermissions_$timestamp.xlsx"

# ---------------------------------
# MODULES
# ---------------------------------
$requiredModules = @("PnP.PowerShell", "ImportExcel")
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Installing module: $module" -ForegroundColor Yellow
        Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $module -Force
}

Add-Type -AssemblyName System.Web

# ---------------------------------
# GLOBALS
# ---------------------------------
$script:StartTime = Get-Date
$global:currentBatch = @()
$global:totalRows = 0
$global:currentSheetNumber = 1
$global:excelInitialized = $false
$global:summaryData = @()

# ---------------------------------
# LOGGING
# ---------------------------------
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARNING","ERROR","SUCCESS")]
        [string]$Level = "INFO"
    )

    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Add-Content -Path $logFilePath -Value $line

    switch ($Level) {
        "ERROR"   { Write-Host $Message -ForegroundColor Red }
        "WARNING" { Write-Host $Message -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $Message -ForegroundColor Green }
        default   { Write-Host $Message -ForegroundColor Cyan }
    }
}

# ---------------------------------
# RETRY WRAPPER
# ---------------------------------
function Invoke-WithRetry {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,
        [int]$MaxRetries = 5,
        [int]$InitialDelaySeconds = 5
    )

    $attempt = 0
    $delay = $InitialDelaySeconds

    while ($attempt -lt $MaxRetries) {
        try {
            return & $ScriptBlock
        }
        catch {
            $attempt++
            $msg = $_.Exception.Message
            $isThrottle = $false

            if ($msg -match "throttl|too many requests|temporarily unavailable|429|503") {
                $isThrottle = $true
            }

            if ($isThrottle -and $attempt -lt $MaxRetries) {
                Write-Log "Throttling detected. Retry $attempt/$MaxRetries after $delay seconds." "WARNING"
                Start-Sleep -Seconds $delay
                $delay = [Math]::Min($delay * 2, 60)
            }
            else {
                Write-Log "Operation failed: $msg" "ERROR"
                throw
            }
        }
    }
}

# ---------------------------------
# READ SITE URLS
# ---------------------------------
function Read-SiteUrls {
    param([string]$Path)

    $urls = Get-Content -Path $Path |
        ForEach-Object { $_.Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    return $urls
}

# ---------------------------------
# CONNECT
# ---------------------------------
function Connect-OneDriveSite {
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
        Write-Log "Connected to $($web.Title) - $($web.Url)" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed to connect to $SiteUrl : $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# ---------------------------------
# IGNORE LISTS/FOLDERS
# ---------------------------------
$ignoreLibraries = @(
    "Form Templates",
    "Site Assets",
    "Style Library",
    "Composed Looks",
    "Converted Forms",
    "Master Page Gallery",
    "List Template Gallery",
    "Theme Gallery",
    "Web Part Gallery",
    "TaxonomyHiddenList",
    "User Information List",
    "appdata",
    "appfiles",
    "Sharing Links",
    "Social",
    "Events"
)

# ---------------------------------
# PERMISSION CLASSIFICATION
# ---------------------------------
function Resolve-PrincipalType {
    param(
        $RoleAssignment
    )

    $member = $RoleAssignment.Member
    $principalName = $member.Title
    $loginName = $member.LoginName
    $email = $member.Email

    $principalType = "Unknown"

    if ($loginName -match "#EXT#") {
        $principalType = "GuestUser"
    }
    elseif ($loginName -match "spo-grid-all-users|everyone") {
        $principalType = "EveryoneOrOrgWide"
    }
    elseif ($member.PrincipalType.ToString() -match "User") {
        $principalType = "InternalUser"
    }
    elseif ($member.PrincipalType.ToString() -match "SharePointGroup") {
        $principalType = "SharePointGroup"
    }
    elseif ($member.PrincipalType.ToString() -match "SecurityGroup|DistributionList") {
        $principalType = "AADOrMailGroup"
    }
    elseif ($member.PrincipalType.ToString() -match "Member") {
        $principalType = "Principal"
    }

    [PSCustomObject]@{
        PrincipalName = if ($principalName) { $principalName } else { $loginName }
        LoginName     = $loginName
        Email         = if ($email) { $email } else { "" }
        PrincipalType = $principalType
    }
}

function Get-MigrationDecision {
    param(
        [string]$PrincipalType,
        [string]$PrincipalName,
        [string]$Email,
        [string]$Roles
    )

    switch ($PrincipalType) {
        "InternalUser" {
            return [PSCustomObject]@{
                MigrationDecision = "Auto"
                CanAutoMigrate    = "Yes"
                Reason            = "Can be preserved if the destination user exists and is included in the identity mapping file."
                ManualAction      = "No manual action if identity mapping is correct."
            }
        }
        "AADOrMailGroup" {
            return [PSCustomObject]@{
                MigrationDecision = "Auto"
                CanAutoMigrate    = "Yes"
                Reason            = "Can be preserved if the destination group exists and is included in the identity mapping file."
                ManualAction      = "No manual action if group is precreated and mapped."
            }
        }
        "SharePointGroup" {
            return [PSCustomObject]@{
                MigrationDecision = "Review"
                CanAutoMigrate    = "Partial"
                Reason            = "SharePoint group membership and target-side group structure should be validated carefully."
                ManualAction      = "Verify the corresponding target group and membership after migration."
            }
        }
        "GuestUser" {
            return [PSCustomObject]@{
                MigrationDecision = "Review"
                CanAutoMigrate    = "Conditional"
                Reason            = "Guest/external access works only if the guest exists in the target tenant and is identity-mapped."
                ManualAction      = "Precreate guest in target tenant and validate access after migration."
            }
        }
        "EveryoneOrOrgWide" {
            return [PSCustomObject]@{
                MigrationDecision = "Manual"
                CanAutoMigrate    = "No"
                Reason            = "Broad tenant-wide access or everyone-style permissions must be revalidated in the target tenant."
                ManualAction      = "Reconfigure manually after migration if still required."
            }
        }
        default {
            return [PSCustomObject]@{
                MigrationDecision = "Manual"
                CanAutoMigrate    = "No"
                Reason            = "Unknown or unresolved principal cannot be guaranteed for automatic migration."
                ManualAction      = "Investigate and recreate manually if needed."
            }
        }
    }
}

# ---------------------------------
# EXCEL WRITING
# ---------------------------------
function Write-BatchToExcel {
    param(
        [array]$Data,
        [string]$FilePath,
        [int]$SheetNumber
    )

    if (-not $Data -or $Data.Count -eq 0) { return }

    $wsName = "Permissions_$SheetNumber"

    try {
        if (-not $global:excelInitialized) {
            $Data | Export-Excel `
                -Path $FilePath `
                -WorksheetName $wsName `
                -TableName "PermissionsTable$SheetNumber" `
                -TableStyle Medium6 `
                -AutoSize `
                -FreezeTopRow `
                -BoldTopRow

            $global:excelInitialized = $true
        }
        else {
            $Data | Export-Excel `
                -Path $FilePath `
                -WorksheetName $wsName `
                -TableName "PermissionsTable$SheetNumber" `
                -TableStyle Medium6 `
                -AutoSize `
                -FreezeTopRow `
                -BoldTopRow
        }

        Write-Log "Wrote $($Data.Count) rows to worksheet $wsName" "SUCCESS"
    }
    catch {
        Write-Log "Excel write failed: $($_.Exception.Message)" "ERROR"
        throw
    }
}

function Add-RowToBatch {
    param([pscustomobject]$Row)

    $global:currentBatch += $Row
    $global:totalRows++

    if ($global:currentBatch.Count -ge $batchSize) {
        Write-BatchToExcel -Data $global:currentBatch -FilePath $outputFilePath -SheetNumber $global:currentSheetNumber
        $global:currentBatch = @()

        if (($global:totalRows % $maxItemsPerSheet) -eq 0) {
            $global:currentSheetNumber++
        }
    }
}

function Create-SummaryWorksheet {
    param([string]$FilePath)

    $summary = [PSCustomObject]@{
        TotalSitesProcessed = $global:summaryData.Count
        TotalPermissionRows = $global:totalRows
        StartTime           = $script:StartTime
        EndTime             = Get-Date
        Duration            = ((Get-Date) - $script:StartTime).ToString()
        OutputFile          = $FilePath
    }

    $summary | Export-Excel -Path $FilePath -WorksheetName "Summary" -TableName "SummaryTable" -TableStyle Medium2 -MoveToStart -AutoSize

    if ($global:summaryData.Count -gt 0) {
        $global:summaryData | Export-Excel -Path $FilePath -WorksheetName "SiteSummary" -TableName "SiteSummaryTable" -TableStyle Medium4 -AutoSize -FreezeTopRow -BoldTopRow
    }
}

# ---------------------------------
# PROCESS ITEM
# ---------------------------------
function Process-UniquePermissionItem {
    param(
        $Item,
        [string]$SiteUrl,
        [string]$LibraryName,
        [string]$ItemType
    )

    try {
        Get-PnPProperty -ClientObject $Item -Property HasUniqueRoleAssignments, RoleAssignments | Out-Null

        # Requirement: ONLY unique permissions
        if (-not $Item.HasUniqueRoleAssignments) {
            return
        }

        $itemName = $Item["FileLeafRef"]
        $itemPath = $Item["FileRef"]

        $createdDate = $null
        $createdBy = ""
        try {
            $createdDate = $Item["Created"]
            $authorField = $Item["Author"]

            if ($null -ne $authorField -and $null -ne $authorField.LookupId) {
                $creator = Get-PnPUser -Identity $authorField.LookupId -ErrorAction SilentlyContinue
                if ($creator) {
                    $createdBy = if ($creator.Email) { "$($creator.Title) ($($creator.Email))" } else { $creator.Title }
                }
            }
        }
        catch {
            Write-Log "Could not resolve CreatedBy for $itemPath" "WARNING"
        }

        foreach ($RoleAssignment in $Item.RoleAssignments) {
            try {
                Get-PnPProperty -ClientObject $RoleAssignment -Property Member, RoleDefinitionBindings | Out-Null

                $roles = $RoleAssignment.RoleDefinitionBindings |
                    Where-Object { $_.Name -ne "Limited Access" } |
                    ForEach-Object { $_.Name }

                if (-not $roles -or $roles.Count -eq 0) {
                    continue
                }

                $principal = Resolve-PrincipalType -RoleAssignment $RoleAssignment
                $decision = Get-MigrationDecision `
                    -PrincipalType $principal.PrincipalType `
                    -PrincipalName $principal.PrincipalName `
                    -Email $principal.Email `
                    -Roles ($roles -join ", ")

                $row = [PSCustomObject]@{
                    SiteURL            = $SiteUrl
                    LibraryName        = $LibraryName
                    ItemType           = $ItemType
                    ItemName           = $itemName
                    ItemPath           = $itemPath
                    PermissionSource   = "DirectACL"
                    PermissionType     = "Unique"
                    GrantedTo          = $principal.PrincipalName
                    GrantedToEmail     = $principal.Email
                    PrincipalType      = $principal.PrincipalType
                    Roles              = ($roles -join ", ")
                    CanAutoMigrate     = $decision.CanAutoMigrate
                    MigrationDecision  = $decision.MigrationDecision
                    Reason             = $decision.Reason
                    ManualAction       = $decision.ManualAction
                    CreatedBy          = $createdBy
                    CreatedDate        = $createdDate
                }

                Add-RowToBatch -Row $row
            }
            catch {
                Write-Log "Failed to process role assignment on $itemPath : $($_.Exception.Message)" "WARNING"
            }
        }
    }
    catch {
        Write-Log "Failed to process item ID $($Item.Id): $($_.Exception.Message)" "ERROR"
    }
}

# ---------------------------------
# MAIN
# ---------------------------------
Write-Log "Script started"
Write-Log "Output file: $outputFilePath"

$siteUrls = Read-SiteUrls -Path $inputFilePath
Write-Log "Found $($siteUrls.Count) site URLs in input file"

foreach ($siteUrl in $siteUrls) {
    $siteStart = Get-Date
    $siteRowsBefore = $global:totalRows
    $siteUniqueItems = 0

    Write-Log "Processing site: $siteUrl"

    if (-not (Connect-OneDriveSite -SiteUrl $siteUrl)) {
        $global:summaryData += [PSCustomObject]@{
            SiteURL              = $siteUrl
            UniqueItemsProcessed = 0
            PermissionRows       = 0
            Duration             = ((Get-Date) - $siteStart).ToString()
            Status               = "Connection Failed"
        }
        continue
    }

    try {
        $lists = Get-PnPList -Includes Title, Hidden, BaseType, ItemCount |
            Where-Object {
                $_.Hidden -eq $false -and
                $_.BaseType -eq "DocumentLibrary" -and
                $_.Title -notin $ignoreLibraries
            }

        foreach ($list in $lists) {
            Write-Log "Processing library: $($list.Title)"

            if ($list.ItemCount -eq 0) {
                Write-Log "Skipping empty library: $($list.Title)"
                continue
            }

            $items = @(Get-PnPListItem -List $list -PageSize 2000 -Fields "FileLeafRef","FileRef","FSObjType","Created","Author")

            foreach ($item in $items) {
                try {
                    $itemType = switch ($item["FSObjType"]) {
                        0 { "File" }
                        1 { "Folder" }
                        default { $null }
                    }

                    if (-not $itemType) { continue }

                    $itemPath = $item["FileRef"]
                    if ([string]::IsNullOrWhiteSpace($itemPath)) { continue }

                    Process-UniquePermissionItem -Item $item -SiteUrl $siteUrl -LibraryName $list.Title -ItemType $itemType
                    $siteUniqueItems++
                }
                catch {
                    Write-Log "Error processing item in library $($list.Title): $($_.Exception.Message)" "WARNING"
                }
            }
        }

        $siteRowsAfter = $global:totalRows
        $rowsAdded = $siteRowsAfter - $siteRowsBefore

        $global:summaryData += [PSCustomObject]@{
            SiteURL              = $siteUrl
            UniqueItemsProcessed = $siteUniqueItems
            PermissionRows       = $rowsAdded
            Duration             = ((Get-Date) - $siteStart).ToString()
            Status               = "Success"
        }

        Write-Log "Completed site: $siteUrl. Rows added: $rowsAdded" "SUCCESS"
    }
    catch {
        Write-Log "Site processing failed for $siteUrl : $($_.Exception.Message)" "ERROR"

        $siteRowsAfter = $global:totalRows
        $rowsAdded = $siteRowsAfter - $siteRowsBefore

        $global:summaryData += [PSCustomObject]@{
            SiteURL              = $siteUrl
            UniqueItemsProcessed = $siteUniqueItems
            PermissionRows       = $rowsAdded
            Duration             = ((Get-Date) - $siteStart).ToString()
            Status               = "Failed"
        }
    }
}

# write final batch
if ($global:currentBatch.Count -gt 0) {
    Write-BatchToExcel -Data $global:currentBatch -FilePath $outputFilePath -SheetNumber $global:currentSheetNumber
    $global:currentBatch = @()
}

Create-SummaryWorksheet -FilePath $outputFilePath

Write-Log "Completed. Total permission rows: $($global:totalRows)" "SUCCESS"
Write-Log "Excel report saved to: $outputFilePath" "SUCCESS"

if (Test-Path $outputFilePath) {
    try {
        Start-Process $outputFilePath
    }
    catch {
        Write-Log "Could not auto-open Excel file. Open manually: $outputFilePath" "WARNING"
    }
}