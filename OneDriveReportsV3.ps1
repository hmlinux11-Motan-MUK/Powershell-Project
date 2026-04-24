#Requires -Modules @{ ModuleName="PnP.PowerShell"; RequiredVersion="2.3.0" }
#Requires -Version 7.2

[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [String[]]$SiteURL,

    [Parameter(Mandatory)]
    [String]$ClientId,

    [Parameter(Mandatory)]
    [String]$Tenant,

    [Parameter(Mandatory)]
    [String]$CertificatePath,

    [Parameter(Mandatory)]
    [String]$CertPlainPassword = "",

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [String[]]$SearchString,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [String[]]$Exclude,

    [Parameter()]
    [Switch]$ReturnResult,

    [Parameter()]
    [String]$OutputFile,

    [Parameter()]
    [Switch]$Quiet
)

begin {

    function Say {
        param(
            [Parameter(Mandatory)]
            [string]$Text,

            [Parameter()]
            [ValidateSet("Black","DarkBlue","DarkGreen","DarkCyan","DarkRed","DarkMagenta","DarkYellow","Gray","DarkGray","Blue","Green","Cyan","Red","Magenta","Yellow","White")]
            [string]$Color = "White"
        )

        if (-not $Quiet) {
            Write-Host $Text -ForegroundColor $Color
        }

        if ($script:LogFile) {
            Add-Content -Path $script:LogFile -Value $Text
        }
    }

    function ReplaceInternalName {
        param([Parameter(Mandatory)][string]$String)

        $internalNames = @{
            '_x0020_' = ' '
            '_x007e_' = '~'
            '_x0021_' = '!'
            '_x0040_' = '@'
            '_x0023_' = '#'
            '_x0024_' = '$'
            '_x0025_' = '%'
            '_x005E_' = '^'
            '_x0026_' = '&'
            '_x002a_' = '*'
            '_x0028_' = '('
            '_x0029_' = ')'
            '_x002B_' = '+'
            '_x002D_' = '-'
            '_x003D_' = '='
            '_x007B_' = '{'
            '_x007D_' = '}'
            '_x003A_' = ':'
            '_x0022_' = "'"
            '_x007C_' = '|'
            '_x003B_' = ';'
            '_x0027_' = "'"
            '_x005C_' = '\'
            '_x003C_' = '<'
            '_x003E_' = '>'
            '_x003F_' = '?'
            '_x002C_' = ','
            '_x002E_' = '.'
            '_x002F_' = '/'
            '_x0060_' = '`'
        }

        foreach ($key in $internalNames.Keys) {
            $String = $String -replace $key, $internalNames[$key]
        }
        return $String
    }

    $now = Get-Date
    $nowString = $now.ToString("yyyy-MM-dd_hh-mm-ss_tt")

    if (-not $OutputFile) {
        $OutputFile = "C:\Temp\SPO_File_Search_$($nowString)_$($env:USERNAME).csv"
    }

    $outputDir = Split-Path -Path $OutputFile -Parent
    if (-not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }

    $script:LogFile = [System.IO.Path]::ChangeExtension($OutputFile, ".log")
    New-Item -ItemType File -Path $OutputFile -Force | Out-Null
    New-Item -ItemType File -Path $script:LogFile -Force | Out-Null

    Say "Results file: $OutputFile" Yellow
    Say "Log file: $script:LogFile" Yellow
    Say "Start @ $((Get-Date).ToString('yyyy-MM-dd hh:mm:ss tt'))" Yellow

    if (-not (Test-Path $CertificatePath)) {
        throw "Certificate file not found: $CertificatePath"
    }

    if ([string]::IsNullOrWhiteSpace($CertPlainPassword)) {
        $script:CertPassword = $null
    }
    else {
        $script:CertPassword = ConvertTo-SecureString $CertPlainPassword -AsPlainText -Force
    }

    $urlPatternToExclude = ".*-my\.sharepoint\.com/$|.*\.sharepoint\.com/$|.*\.sharepoint\.com/search$|.*\.sharepoint\.com/portals/hub$|.*\.sharepoint\.com/sites/appcatalog$"
    $SiteURL = $SiteURL | Where-Object { $_ -notmatch $urlPatternToExclude }

    $SystemLibraries = @(
        'Form Templates',
        'Pages',
        'Preservation Hold Library',
        'Site Assets',
        'Site Pages',
        'Images',
        'Site Collection Documents',
        'Site Collection Images',
        'Style Library'
    )

    $filterPattern = (
        $SearchString | ForEach-Object {
            if ($_ -match '\*\.\w+') {
                "$([regex]::Escape($_) -replace '\\\*', '.*')$"
            }
            elseif ($_ -match '\*') {
                "$([regex]::Escape($_) -replace '\\\*', '.*')$"
            }
            else {
                "$([regex]::Escape($_))$"
            }
        }
    ) -join '|'

    if ($Exclude) {
        $excludePattern = (
            $Exclude | ForEach-Object {
                if ($_ -match '\*\.\w+') {
                    "$([regex]::Escape($_) -replace '\\\*', '.*')$"
                }
                elseif ($_ -match '\*') {
                    "$([regex]::Escape($_) -replace '\\\*', '.*')$"
                }
                else {
                    "$([regex]::Escape($_))$"
                }
            }
        ) -join '|'
    }

    Say "Sites to search: $($SiteURL.Count)" Yellow
    Say "Search Filter: $filterPattern" Yellow
    if ($excludePattern) {
        Say "Exclude Filter: $excludePattern" Yellow
    }
}

process {
    for ($urlIndex = 0; $urlIndex -lt $SiteURL.Count; $urlIndex++) {
        $url = $SiteURL[$urlIndex]
        $siteTitleForError = $url

        try {
            Say "Site $($urlIndex + 1) of $($SiteURL.Count): [$url]" Cyan

            if ($script:CertPassword) {
                Connect-PnPOnline `
                    -Url $url `
                    -ClientId $ClientId `
                    -Tenant $Tenant `
                    -CertificatePath $CertificatePath `
                    -CertificatePassword $script:CertPassword `
                    -ErrorAction Stop
            }
            else {
                Connect-PnPOnline `
                    -Url $url `
                    -ClientId $ClientId `
                    -Tenant $Tenant `
                    -CertificatePath $CertificatePath `
                    -ErrorAction Stop
            }

            $web = Get-PnPWeb -ErrorAction Stop
            $siteTitleForError = $web.Title

            if ($url -like "*-my.sharepoint.com*") {
                Say "  -> OneDrive Name: $($web.Title)" Yellow
                $siteType = "OneDrive"
            }
            else {
                Say "  -> Site Name: $($web.Title)" Yellow
                $siteType = "SharePoint"
            }

            $DocumentLibraries = @(
                Get-PnPList -ErrorAction Stop | Where-Object {
                    $_.BaseType -eq 'DocumentLibrary' -and
                    $_.Hidden -eq $false -and
                    $_.Title -notin $SystemLibraries
                }
            )

            foreach ($library in $DocumentLibraries) {
                $library | Add-Member -MemberType NoteProperty -Name Leaf -Value (ReplaceInternalName -String $library.EntityTypeName) -Force
            }

            foreach ($library in $DocumentLibraries) {
                Say "    -> Library name: [$($library.Title)]" Magenta

                [System.Collections.Generic.List[System.Object]]$searchResult = @()

                try {
                    $files = Get-PnPFolderItem -FolderSiteRelativeUrl $library.RootFolder.ServerRelativeUrl -ItemType File -Recursive -ErrorAction Stop
                }
                catch {
                    Say "    -> Failed to read library [$($library.Title)]: $($_.Exception.Message)" Red
                    continue
                }

                if ($Exclude) {
                    $matchedFiles = $files | Where-Object {
                        $_.Name -match $filterPattern -and
                        (($_.ServerRelativeUrl -split '/')[-2]) -ne "Forms" -and
                        $_.Name -notmatch $excludePattern
                    }
                }
                else {
                    $matchedFiles = $files | Where-Object {
                        $_.Name -match $filterPattern -and
                        (($_.ServerRelativeUrl -split '/')[-2]) -ne "Forms"
                    }
                }

                if ($matchedFiles) {
                    $searchResult.AddRange(@($matchedFiles))
                }

                if ($searchResult.Count -gt 0) {
                    Say "      -> Items: $($searchResult.Count)" Green

                    $output = $searchResult | Select-Object `
                        @{ n = "SiteUrl"; e = { $url } },
                        @{ n = "SiteName"; e = { $web.Title } },
                        @{ n = "SiteType"; e = { $siteType } },
                        @{ n = "LibraryName"; e = { $library.Title } },
                        @{ n = "ParentPath"; e = {
                            "/" + (((($_.ServerRelativeUrl -split '/') | Select-Object -Skip 3 | Select-Object -SkipLast 1) -join "/")) + "/"
                        }},
                        @{ n = "FileName"; e = { $_.Name } },
                        @{ n = "FileType"; e = {
                            if ($_.Name -match '\.') { ($_.Name.ToString().Split('.'))[-1] } else { "" }
                        }},
                        @{ n = "SizeKB"; e = {
                            if ($_.Length -ne $null) { [math]::Round(($_.Length / 1KB), 2) } else { "" }
                        }},
                        @{ n = "TimeCreated"; e = { $_.TimeCreated } },
                        @{ n = "TimeLastModified"; e = { $_.TimeLastModified } },
                        @{ n = "ServerRelativeUrl"; e = { $_.ServerRelativeUrl } }

                    if ($ReturnResult) {
                        $output
                    }

                    $output | Export-Csv -NoTypeInformation -Append -Path $OutputFile
                }
            }

            Disconnect-PnPOnline -ErrorAction SilentlyContinue
        }
        catch {
            Say "[ERROR] - [$siteTitleForError]: $($_.Exception.Message)" Red
        }
    }
}

end {
    Say "End @ $((Get-Date).ToString('yyyy-MM-dd hh:mm:ss tt'))" Yellow
}