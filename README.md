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

