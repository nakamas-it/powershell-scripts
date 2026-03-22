#Requires -Version 5.1
<#
.SYNOPSIS
    Scans files and folders on a Windows computer and generates detailed HTML reports about storage usage.

.DESCRIPTION
    Get-DiskStorageReport performs a comprehensive scan of one or more drives or paths on a Windows
    system and produces four standalone HTML report files:

      1. TopFolders.html   - Largest directories with drill-down by drive
      2. TopFiles.html     - Largest individual files across all scanned paths
      3. Dashboard.html    - Summary dashboard with key metrics and charts
      4. FileTypes.html    - Storage breakdown by file extension and category

    All HTML files are fully self-contained (inline CSS, inline JS charts) and work
    with zero internet access. Charts are rendered using pure HTML/CSS bar charts.

    The script is designed to run without administrator privileges. When access to a
    directory is denied, the path is logged and scanning continues gracefully.

.PARAMETER RootPath
    One or more root paths to scan. Accepts an array of strings.
    Default: all fixed local drives detected via Get-Volume / Get-WmiObject.

.PARAMETER OutputFolder
    Directory where the HTML report files will be written.
    Default: C:\Temp\DiskReport

.PARAMETER TopN
    Number of top items to display in each ranked report (top folders, top files).
    Default: 50

.PARAMETER ExcludePaths
    Array of full path prefixes to skip during scanning (case-insensitive).
    Default: @('C:\Windows', 'C:\$Recycle.Bin', 'C:\System Volume Information')

.EXAMPLE
    .\Get-DiskStorageReport.ps1
    Scans all fixed drives with default settings and writes reports to ~\DiskReport.

.EXAMPLE
    .\Get-DiskStorageReport.ps1 -RootPath 'D:\' -TopN 100 -OutputFolder 'C:\Reports'
    Scans only D:\ and shows the top 100 items per report.

.EXAMPLE
    .\Get-DiskStorageReport.ps1 -RootPath 'C:\','D:\' -ExcludePaths 'C:\Windows','C:\Program Files'
    Scans C:\ and D:\ while skipping the Windows and Program Files directories.

.NOTES
    Author  : AI-Assisted Script (PowerShell Expert)
    Version : 1.0.0
    Date    : 2026-03-21
    License : MIT
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]]$RootPath,

    [Parameter()]
    [string]$OutputFolder = 'C:\Temp\DiskReport',

    [Parameter()]
    [ValidateRange(1, 10000)]
    [int]$TopN = 50,

    [Parameter()]
    [string[]]$ExcludePaths = @(
        'C:\Windows',
        'C:\Program Files',
        'C:\Program Files (x86)',
        'C:\$Recycle.Bin',
        'C:\System Volume Information'
    )
)

# ---------------------------------------------------------------------------
# Region: Utility Functions
# ---------------------------------------------------------------------------

function Format-FileSize {
    <#
    .SYNOPSIS
        Converts a byte count to a human-readable string (B, KB, MB, GB, TB).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [double]$Bytes
    )

    if ($Bytes -ge 1TB) { return '{0:N2} TB' -f ($Bytes / 1TB) }
    if ($Bytes -ge 1GB) { return '{0:N2} GB' -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return '{0:N2} MB' -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return '{0:N2} KB' -f ($Bytes / 1KB) }
    return '{0:N0} B' -f $Bytes
}

function Get-SizeColorClass {
    <#
    .SYNOPSIS
        Returns a CSS class name based on the size of an item relative to thresholds.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [double]$Bytes
    )

    if ($Bytes -ge 10GB) { return 'size-critical' }
    if ($Bytes -ge 1GB)  { return 'size-warning'  }
    if ($Bytes -ge 100MB){ return 'size-moderate'  }
    return 'size-normal'
}

function Test-PathExcluded {
    <#
    .SYNOPSIS
        Returns $true if the given path starts with any of the excluded prefixes.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [string[]]$ExcludeList
    )

    foreach ($excluded in $ExcludeList) {
        if ($Path -like "$excluded*") {
            return $true
        }
    }
    return $false
}

function Get-FileCategory {
    <#
    .SYNOPSIS
        Maps a file extension to a human-readable category.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Extension
    )

    $ext = $Extension.ToLowerInvariant().TrimStart('.')

    $categoryMap = @{
        'Video'        = @('mp4','mkv','avi','mov','wmv','flv','webm','m4v','mpg','mpeg','3gp','ts','vob')
        'Audio'        = @('mp3','wav','flac','aac','ogg','wma','m4a','opus','aiff','alac')
        'Image'        = @('jpg','jpeg','png','gif','bmp','tiff','tif','svg','ico','webp','raw','cr2','nef','heic','heif','psd','ai')
        'Document'     = @('pdf','doc','docx','xls','xlsx','ppt','pptx','odt','ods','odp','rtf','txt','csv','epub','mobi')
        'Archive'      = @('zip','rar','7z','tar','gz','bz2','xz','cab','iso','img','dmg','wim')
        'Executable'   = @('exe','msi','dll','sys','drv','ocx','cpl','scr','com')
        'Script'       = @('ps1','psm1','psd1','bat','cmd','vbs','js','py','sh','rb','pl','php')
        'Database'     = @('mdb','accdb','sqlite','db','sql','bak','mdf','ldf','ndf')
        'VM / Disk'    = @('vhd','vhdx','vmdk','vdi','qcow2','ova','ovf')
        'Log'          = @('log','etl','evtx','dmp','mdmp')
        'Web'          = @('html','htm','css','json','xml','yaml','yml')
        'Source Code'  = @('cs','cpp','c','h','java','go','rs','ts','tsx','jsx','swift','kt')
        'Font'         = @('ttf','otf','woff','woff2','eot')
        'Backup'       = @('bkf','bkp','old','orig','backup')
    }

    foreach ($category in $categoryMap.Keys) {
        if ($categoryMap[$category] -contains $ext) {
            return $category
        }
    }

    if ([string]::IsNullOrWhiteSpace($ext)) {
        return 'No Extension'
    }

    return 'Other'
}

function Get-HtmlHeader {
    <#
    .SYNOPSIS
        Returns the common HTML header including inline CSS used by all reports.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Title,

        [Parameter()]
        [string]$ExtraStyle = ''
    )

    return @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>$Title</title>
<style>
/* ---- Base Reset & Typography ---- */
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
body {
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    background: #f0f2f5; color: #1a1a2e; line-height: 1.6;
    padding: 20px 30px 40px;
}
h1 { font-size: 1.8rem; margin-bottom: 6px; color: #16213e; }
h2 { font-size: 1.35rem; margin: 30px 0 12px; color: #1a1a2e; border-bottom: 2px solid #0f3460; padding-bottom: 4px; }
h3 { font-size: 1.1rem; margin: 20px 0 8px; color: #333; }
.subtitle { color: #555; font-size: 0.92rem; margin-bottom: 18px; }

/* ---- Navigation ---- */
.nav { display: flex; gap: 10px; margin-bottom: 24px; flex-wrap: wrap; }
.nav a {
    display: inline-block; padding: 8px 18px; background: #0f3460; color: #fff;
    text-decoration: none; border-radius: 5px; font-size: 0.88rem; transition: background 0.2s;
}
.nav a:hover { background: #533483; }
.nav a.active { background: #e94560; }

/* ---- Cards ---- */
.card-row { display: flex; gap: 16px; flex-wrap: wrap; margin-bottom: 20px; }
.card {
    flex: 1 1 200px; background: #fff; border-radius: 8px; padding: 18px 22px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.08); min-width: 180px;
}
.card .label { font-size: 0.82rem; color: #777; text-transform: uppercase; letter-spacing: 0.5px; }
.card .value { font-size: 1.55rem; font-weight: 700; color: #16213e; margin-top: 4px; }
.card .detail { font-size: 0.82rem; color: #999; margin-top: 2px; }

/* ---- Tables ---- */
.table-wrapper { overflow-x: auto; margin-bottom: 24px; }
table {
    width: 100%; border-collapse: collapse; background: #fff;
    border-radius: 8px; overflow: hidden; box-shadow: 0 2px 6px rgba(0,0,0,0.06);
}
thead th {
    background: #16213e; color: #fff; padding: 10px 14px; text-align: left;
    font-size: 0.85rem; text-transform: uppercase; letter-spacing: 0.3px;
    cursor: pointer; user-select: none; position: relative; white-space: nowrap;
}
thead th:hover { background: #1a1a4e; }
thead th .sort-arrow { margin-left: 4px; font-size: 0.7rem; opacity: 0.5; }
thead th.sorted-asc .sort-arrow::after { content: ' \u25B2'; opacity: 1; }
thead th.sorted-desc .sort-arrow::after { content: ' \u25BC'; opacity: 1; }
tbody td { padding: 9px 14px; border-bottom: 1px solid #eee; font-size: 0.88rem; }
tbody tr:hover { background: #f7f9fc; }
tbody tr:last-child td { border-bottom: none; }
.rank { color: #999; font-weight: 600; }

/* ---- Size color classes ---- */
.size-critical { color: #c0392b; font-weight: 700; }
.size-warning  { color: #e67e22; font-weight: 600; }
.size-moderate { color: #2980b9; }
.size-normal   { color: #27ae60; }

/* ---- Bar Charts (pure CSS) ---- */
.bar-chart { margin: 12px 0 20px; }
.bar-row { display: flex; align-items: center; margin-bottom: 5px; }
.bar-label {
    width: 200px; min-width: 120px; font-size: 0.82rem; text-align: right;
    padding-right: 12px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
    color: #333;
}
.bar-track { flex: 1; background: #e8ecf1; border-radius: 4px; height: 22px; position: relative; overflow: hidden; }
.bar-fill {
    height: 100%; border-radius: 4px; transition: width 0.4s ease;
    display: flex; align-items: center; padding-left: 8px;
    font-size: 0.75rem; color: #fff; font-weight: 600; white-space: nowrap;
    min-width: 2px;
}
.bar-value { margin-left: 10px; font-size: 0.82rem; color: #555; white-space: nowrap; min-width: 80px; }

/* ---- Palette for bars ---- */
.clr-0  { background: #0f3460; } .clr-1  { background: #e94560; }
.clr-2  { background: #533483; } .clr-3  { background: #16a085; }
.clr-4  { background: #2980b9; } .clr-5  { background: #8e44ad; }
.clr-6  { background: #d35400; } .clr-7  { background: #27ae60; }
.clr-8  { background: #c0392b; } .clr-9  { background: #f39c12; }
.clr-10 { background: #1abc9c; } .clr-11 { background: #e74c3c; }
.clr-12 { background: #3498db; } .clr-13 { background: #9b59b6; }
.clr-14 { background: #e67e22; }

/* ---- Drive usage bar ---- */
.drive-bar-outer {
    width: 100%; height: 20px; background: #dfe6e9; border-radius: 4px;
    overflow: hidden; position: relative;
}
.drive-bar-inner {
    height: 100%; border-radius: 4px; transition: width 0.3s;
}
.drive-bar-green  { background: linear-gradient(90deg, #27ae60, #2ecc71); }
.drive-bar-yellow { background: linear-gradient(90deg, #f39c12, #f1c40f); }
.drive-bar-red    { background: linear-gradient(90deg, #e74c3c, #c0392b); }

/* ---- Footer ---- */
.footer { margin-top: 40px; padding-top: 14px; border-top: 1px solid #ccc; font-size: 0.78rem; color: #888; }

/* ---- Access-denied log ---- */
.access-log { max-height: 300px; overflow-y: auto; background: #fff8e1; padding: 12px; border-radius: 6px; font-size: 0.8rem; color: #795548; margin-top: 10px; }
.access-log p { margin-bottom: 2px; }
$ExtraStyle
</style>
</head>
<body>
"@
}

function Get-NavHtml {
    <#
    .SYNOPSIS
        Returns the navigation bar HTML with the active page highlighted.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ActivePage
    )

    $pages = [ordered]@{
        'Dashboard.html'  = 'Dashboard'
        'TopFolders.html' = 'Top Folders'
        'TopFiles.html'   = 'Top Files'
        'FileTypes.html'  = 'File Types'
    }

    $html = '<div class="nav">'
    foreach ($file in $pages.Keys) {
        $label = $pages[$file]
        $cls = if ($file -eq $ActivePage) { ' class="active"' } else { '' }
        $html += "`n  <a href=`"$file`"$cls>$label</a>"
    }
    $html += "`n</div>"
    return $html
}

function Get-SortableTableScript {
    <#
    .SYNOPSIS
        Returns inline JavaScript that makes tables sortable by clicking column headers.
    #>
    [CmdletBinding()]
    param()

    return @'
<script>
(function(){
  document.querySelectorAll('table.sortable').forEach(function(table){
    var headers = table.querySelectorAll('thead th');
    headers.forEach(function(th, colIdx){
      th.innerHTML += '<span class="sort-arrow"></span>';
      th.addEventListener('click', function(){
        var tbody = table.querySelector('tbody');
        var rows = Array.from(tbody.querySelectorAll('tr'));
        var dir = th.classList.contains('sorted-asc') ? -1 : 1;
        headers.forEach(function(h){ h.classList.remove('sorted-asc','sorted-desc'); });
        th.classList.add(dir === 1 ? 'sorted-asc' : 'sorted-desc');
        rows.sort(function(a, b){
          var av = a.children[colIdx].getAttribute('data-sort') || a.children[colIdx].textContent.trim();
          var bv = b.children[colIdx].getAttribute('data-sort') || b.children[colIdx].textContent.trim();
          var an = parseFloat(av), bn = parseFloat(bv);
          if (!isNaN(an) && !isNaN(bn)) return (an - bn) * dir;
          return av.localeCompare(bv) * dir;
        });
        rows.forEach(function(r){ tbody.appendChild(r); });
      });
    });
  });
})();
</script>
'@
}

function Get-HtmlFooter {
    <#
    .SYNOPSIS
        Returns the standard footer and closing HTML tags.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [datetime]$ScanStart,

        [Parameter(Mandatory)]
        [datetime]$ScanEnd,

        [Parameter()]
        [string[]]$ScannedPaths = @()
    )

    $duration = $ScanEnd - $ScanStart
    $durationStr = '{0}h {1}m {2}s' -f [int]$duration.TotalHours, $duration.Minutes, $duration.Seconds
    $pathList = ($ScannedPaths | ForEach-Object { "<code>$_</code>" }) -join ', '

    return @"
<div class="footer">
  <p>Report generated: <strong>$($ScanEnd.ToString('yyyy-MM-dd HH:mm:ss'))</strong> &mdash;
     Scan duration: <strong>$durationStr</strong> &mdash;
     Scanned: $pathList</p>
  <p>Generated by <strong>Get-DiskStorageReport.ps1</strong> v1.0.0</p>
</div>
$(Get-SortableTableScript)
</body>
</html>
"@
}

function New-BarChartHtml {
    <#
    .SYNOPSIS
        Builds a pure-CSS horizontal bar chart from label/value pairs.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$Data,          # Array of objects with .Label and .Value (numeric bytes)

        [Parameter()]
        [int]$MaxBars = 15,

        [Parameter()]
        [string]$ChartTitle = ''
    )

    if (-not $Data -or $Data.Count -eq 0) { return '<p>No data available.</p>' }

    $subset = $Data | Select-Object -First $MaxBars
    $maxVal = ($subset | Measure-Object -Property Value -Maximum).Maximum
    if ($maxVal -le 0) { $maxVal = 1 }

    $html = ''
    if ($ChartTitle) { $html += "<h3>$ChartTitle</h3>`n" }
    $html += "<div class=`"bar-chart`">`n"

    $i = 0
    foreach ($item in $subset) {
        $pct = [math]::Round(($item.Value / $maxVal) * 100, 1)
        if ($pct -lt 0.5) { $pct = 0.5 }
        $clr = 'clr-{0}' -f ($i % 15)
        $sizeStr = Format-FileSize -Bytes $item.Value
        $safeLabel = [System.Net.WebUtility]::HtmlEncode($item.Label)
        $html += @"
  <div class="bar-row">
    <div class="bar-label" title="$safeLabel">$safeLabel</div>
    <div class="bar-track"><div class="bar-fill $clr" style="width:${pct}%;">$( if ($pct -ge 12) { $sizeStr } else { '' } )</div></div>
    <div class="bar-value">$sizeStr</div>
  </div>

"@
        $i++
    }

    $html += "</div>`n"
    return $html
}

# ---------------------------------------------------------------------------
# Region: Drive Discovery
# ---------------------------------------------------------------------------

function Get-FixedDrives {
    <#
    .SYNOPSIS
        Returns an array of drive root paths (e.g., C:\) for all fixed local drives.
    #>
    [CmdletBinding()]
    param()

    try {
        # Prefer Get-CimInstance (available in PS 5.1+)
        $drives = Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction Stop |
                  Where-Object { $_.Size -gt 0 } |
                  ForEach-Object { "$($_.DeviceID)\" }
        return $drives
    }
    catch {
        Write-Warning "Get-CimInstance failed, falling back to Get-PSDrive: $_"
        return Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue |
               Where-Object { $_.Used -ne $null -and $_.Free -ne $null } |
               ForEach-Object { "$($_.Root)" }
    }
}

function Get-DriveSpaceInfo {
    <#
    .SYNOPSIS
        Returns drive space details (total, used, free, percent used) for a drive root.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DriveRoot
    )

    try {
        $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$($DriveRoot.TrimEnd('\'))'" -ErrorAction Stop
        if ($disk) {
            return [PSCustomObject]@{
                Drive       = $DriveRoot
                Label       = if ($disk.VolumeName) { $disk.VolumeName } else { '(No Label)' }
                TotalBytes  = [double]$disk.Size
                FreeBytes   = [double]$disk.FreeSpace
                UsedBytes   = [double]($disk.Size - $disk.FreeSpace)
                PercentUsed = if ($disk.Size -gt 0) { [math]::Round(($disk.Size - $disk.FreeSpace) / $disk.Size * 100, 1) } else { 0 }
            }
        }
    }
    catch {
        Write-Warning "Could not retrieve disk info for ${DriveRoot}: $_"
    }

    return $null
}

# ---------------------------------------------------------------------------
# Region: Scanning Engine
# ---------------------------------------------------------------------------

function Start-DiskScan {
    <#
    .SYNOPSIS
        Recursively scans all files under the given root paths, collecting file info
        and computing folder sizes. Returns a result object with all collected data.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$Paths,

        [Parameter(Mandatory)]
        [string[]]$ExcludeList,

        [Parameter()]
        [int]$TopN = 50
    )

    # Data collections
    $allFiles       = [System.Collections.Generic.List[PSCustomObject]]::new()
    $folderSizes    = @{}    # path -> total bytes
    $extensionStats = @{}    # extension -> @{ Count; TotalBytes }
    $accessDenied   = [System.Collections.Generic.List[string]]::new()
    $totalFiles     = 0
    $totalFolders   = 0
    $totalBytes     = [double]0
    $pathIndex      = 0

    foreach ($rootPath in $Paths) {
        $pathIndex++
        Write-Verbose "Scanning root path $pathIndex of $($Paths.Count): $rootPath"

        # Use .NET for performance: enumerate all file-system entries
        $queue = [System.Collections.Generic.Queue[string]]::new()
        $queue.Enqueue($rootPath)

        while ($queue.Count -gt 0) {
            $currentDir = $queue.Dequeue()

            # Check exclusion
            if (Test-PathExcluded -Path $currentDir -ExcludeList $ExcludeList) {
                continue
            }

            # Progress
            $totalFolders++
            if ($totalFolders % 500 -eq 0) {
                Write-Progress -Activity 'Scanning disk' -Status "Folders: $totalFolders | Files: $totalFiles | $(Format-FileSize -Bytes $totalBytes)" -CurrentOperation $currentDir -PercentComplete -1
            }

            # Try to get child directories
            try {
                $subDirs = [System.IO.Directory]::GetDirectories($currentDir)
                foreach ($sd in $subDirs) {
                    $queue.Enqueue($sd)
                }
            }
            catch [System.UnauthorizedAccessException] {
                $accessDenied.Add($currentDir)
                continue
            }
            catch {
                $accessDenied.Add("$currentDir (Error: $($_.Exception.Message))")
                continue
            }

            # Try to get files in this directory
            try {
                $files = [System.IO.Directory]::GetFiles($currentDir)
            }
            catch [System.UnauthorizedAccessException] {
                $accessDenied.Add($currentDir)
                continue
            }
            catch {
                continue
            }

            $dirSize = [double]0

            foreach ($filePath in $files) {
                try {
                    $fi = [System.IO.FileInfo]::new($filePath)
                    $len = $fi.Length
                    $dirSize += $len
                    $totalBytes += $len
                    $totalFiles++

                    # Collect file record
                    $allFiles.Add([PSCustomObject]@{
                        FullName      = $fi.FullName
                        Name          = $fi.Name
                        Extension     = $fi.Extension
                        Length        = $len
                        LastWriteTime = $fi.LastWriteTime
                        DirectoryName = $fi.DirectoryName
                    })

                    # Extension stats
                    $ext = $fi.Extension.ToLowerInvariant()
                    if (-not $extensionStats.ContainsKey($ext)) {
                        $extensionStats[$ext] = @{ Count = 0; TotalBytes = [double]0 }
                    }
                    $extensionStats[$ext].Count++
                    $extensionStats[$ext].TotalBytes += $len
                }
                catch {
                    # Silently skip files we cannot read metadata for
                }
            }

            # Accumulate directory size up the tree
            $dir = $currentDir
            while ($true) {
                if (-not $folderSizes.ContainsKey($dir)) {
                    $folderSizes[$dir] = [double]0
                }
                $folderSizes[$dir] += $dirSize

                $parent = [System.IO.Path]::GetDirectoryName($dir)
                if ([string]::IsNullOrEmpty($parent) -or $parent -eq $dir) { break }
                $dir = $parent
            }
        }
    }

    Write-Progress -Activity 'Scanning disk' -Completed

    # Build top-files list
    Write-Verbose 'Sorting top files...'
    $topFiles = $allFiles |
                Sort-Object -Property Length -Descending |
                Select-Object -First $TopN

    # Build top-folders list (drive roots like C:\ are excluded — they are not meaningful as "folders")
    Write-Verbose 'Sorting top folders...'
    $topFolders = $folderSizes.GetEnumerator() |
                  Where-Object { $_.Key -notmatch '^[A-Za-z]:\\$' } |
                  Sort-Object -Property Value -Descending |
                  Select-Object -First $TopN |
                  ForEach-Object {
                      [PSCustomObject]@{
                          Path      = $_.Key
                          SizeBytes = $_.Value
                      }
                  }

    # Build extension summary
    Write-Verbose 'Building extension summary...'
    $extSummary = $extensionStats.GetEnumerator() |
                  ForEach-Object {
                      [PSCustomObject]@{
                          Extension  = if ($_.Key) { $_.Key } else { '(none)' }
                          Category   = Get-FileCategory -Extension $_.Key
                          Count      = $_.Value.Count
                          TotalBytes = $_.Value.TotalBytes
                      }
                  } |
                  Sort-Object -Property TotalBytes -Descending

    # Build category summary
    $categorySummary = $extSummary |
                       Group-Object -Property Category |
                       ForEach-Object {
                           [PSCustomObject]@{
                               Category   = $_.Name
                               Count      = ($_.Group | Measure-Object -Property Count -Sum).Sum
                               TotalBytes = ($_.Group | Measure-Object -Property TotalBytes -Sum).Sum
                           }
                       } |
                       Sort-Object -Property TotalBytes -Descending

    # Build per-drive statistics for the dashboard
    Write-Verbose 'Building per-drive statistics...'
    $perDriveStats = [ordered]@{}
    foreach ($rootPath in $Paths) {
        $dl = (($rootPath -split '\\')[0] + '\').ToUpperInvariant()

        $driveFiles   = $allFiles | Where-Object { $_.FullName.ToUpperInvariant().StartsWith($dl) }
        $driveBytes   = ($driveFiles | Measure-Object -Property Length -Sum).Sum
        if ($null -eq $driveBytes) { $driveBytes = [double]0 }

        $driveTopFile = $driveFiles | Sort-Object -Property Length -Descending | Select-Object -First 1

        $driveFolderEntry = $folderSizes.GetEnumerator() |
            Where-Object { $_.Key.ToUpperInvariant().StartsWith($dl) -and $_.Key -notmatch '^[A-Za-z]:\\$' } |
            Sort-Object -Property Value -Descending |
            Select-Object -First 1

        $driveAccessDeniedCount = ($accessDenied | Where-Object { $_.ToUpperInvariant().StartsWith($dl) } | Measure-Object).Count
        $driveFolderCount       = ($folderSizes.Keys | Where-Object { $_.ToUpperInvariant().StartsWith($dl) } | Measure-Object).Count

        $perDriveStats[$dl] = [PSCustomObject]@{
            Drive             = $dl
            TotalFiles        = ($driveFiles | Measure-Object).Count
            TotalFolders      = $driveFolderCount
            TotalBytes        = $driveBytes
            LargestFile       = $driveTopFile
            LargestFolder     = if ($driveFolderEntry) { [PSCustomObject]@{ Path = $driveFolderEntry.Key; SizeBytes = $driveFolderEntry.Value } } else { $null }
            AccessDeniedCount = $driveAccessDeniedCount
        }
    }

    return [PSCustomObject]@{
        TotalFiles       = $totalFiles
        TotalFolders     = $totalFolders
        TotalBytes       = $totalBytes
        TopFiles         = $topFiles
        TopFolders       = $topFolders
        ExtensionSummary = $extSummary
        CategorySummary  = $categorySummary
        AccessDenied     = $accessDenied
        AllFiles         = $allFiles
        FolderSizes      = $folderSizes
        PerDriveStats    = $perDriveStats
    }
}

# ---------------------------------------------------------------------------
# Region: Report Generators
# ---------------------------------------------------------------------------

function New-DashboardReport {
    <#
    .SYNOPSIS
        Generates the Dashboard.html summary report.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $ScanResult,
        [Parameter(Mandatory)] [string]$OutputFile,
        [Parameter(Mandatory)] [array]$DriveInfo,
        [Parameter(Mandatory)] [datetime]$ScanStart,
        [Parameter(Mandatory)] [datetime]$ScanEnd,
        [Parameter(Mandatory)] [string[]]$ScannedPaths
    )

    $html = [System.Text.StringBuilder]::new()
    [void]$html.Append((Get-HtmlHeader -Title 'Disk Report - Dashboard'))
    [void]$html.AppendLine((Get-NavHtml -ActivePage 'Dashboard.html'))

    [void]$html.AppendLine('<h1>Storage Dashboard</h1>')
    [void]$html.AppendLine("<p class=`"subtitle`">Scanned: $($ScannedPaths -join ', ') &mdash; $(($ScanEnd).ToString('yyyy-MM-dd HH:mm:ss'))</p>")

    # ---- Per-Drive Summary Cards ----
    foreach ($driveLetter in $ScanResult.PerDriveStats.Keys) {
        $ds = $ScanResult.PerDriveStats[$driveLetter]
        $di = $DriveInfo | Where-Object { ($_.Drive.ToUpperInvariant().TrimEnd('\') + '\') -eq $driveLetter }

        $largestFileSize   = if ($ds.LargestFile)   { Format-FileSize -Bytes $ds.LargestFile.Length }          else { 'N/A' }
        $largestFileName   = if ($ds.LargestFile)   { [System.Net.WebUtility]::HtmlEncode($ds.LargestFile.Name) } else { '&mdash;' }
        $largestFolderSize = if ($ds.LargestFolder) { Format-FileSize -Bytes $ds.LargestFolder.SizeBytes }     else { 'N/A' }
        $largestFolderName = if ($ds.LargestFolder) { [System.Net.WebUtility]::HtmlEncode(($ds.LargestFolder.Path -split '\\' | Where-Object { $_ } | Select-Object -Last 1)) } else { '&mdash;' }
        $capacity          = if ($di) { Format-FileSize -Bytes $di.TotalBytes } else { 'N/A' }
        $free              = if ($di) { Format-FileSize -Bytes $di.FreeBytes  } else { 'N/A' }
        $driveLabel        = if ($di -and $di.Label -and $di.Label -ne '(No Label)') { " &mdash; $([System.Net.WebUtility]::HtmlEncode($di.Label))" } else { '' }

        [void]$html.AppendLine("<h2>Drive $driveLetter$driveLabel</h2>")
        [void]$html.AppendLine('<div class="card-row">')
        [void]$html.AppendLine("  <div class=`"card`"><div class=`"label`">Total Scanned</div><div class=`"value`">$(Format-FileSize -Bytes $ds.TotalBytes)</div><div class=`"detail`">$($ds.TotalFiles.ToString('N0')) files, $($ds.TotalFolders.ToString('N0')) folders</div></div>")
        [void]$html.AppendLine("  <div class=`"card`"><div class=`"label`">Total Capacity</div><div class=`"value`">$capacity</div><div class=`"detail`">Free: $free</div></div>")
        [void]$html.AppendLine("  <div class=`"card`"><div class=`"label`">Largest File</div><div class=`"value`">$largestFileSize</div><div class=`"detail`">$largestFileName</div></div>")
        [void]$html.AppendLine("  <div class=`"card`"><div class=`"label`">Largest Folder</div><div class=`"value`">$largestFolderSize</div><div class=`"detail`">$largestFolderName</div></div>")
        [void]$html.AppendLine("  <div class=`"card`"><div class=`"label`">Access Denied</div><div class=`"value`">$($ds.AccessDeniedCount)</div><div class=`"detail`">paths skipped</div></div>")
        [void]$html.AppendLine('</div>')
    }

    # ---- Drive Usage ----
    [void]$html.AppendLine('<h2>Drive Usage</h2>')
    [void]$html.AppendLine('<div class="table-wrapper"><table class="sortable"><thead><tr><th>Drive</th><th>Label</th><th>Total</th><th>Used</th><th>Free</th><th>% Used</th><th style="min-width:200px;">Usage</th></tr></thead><tbody>')

    foreach ($d in $DriveInfo) {
        $barClass = if ($d.PercentUsed -ge 90) { 'drive-bar-red' } elseif ($d.PercentUsed -ge 70) { 'drive-bar-yellow' } else { 'drive-bar-green' }
        [void]$html.AppendLine("<tr>")
        [void]$html.AppendLine("  <td>$($d.Drive)</td>")
        [void]$html.AppendLine("  <td>$([System.Net.WebUtility]::HtmlEncode($d.Label))</td>")
        [void]$html.AppendLine("  <td data-sort=`"$($d.TotalBytes)`">$(Format-FileSize -Bytes $d.TotalBytes)</td>")
        [void]$html.AppendLine("  <td data-sort=`"$($d.UsedBytes)`">$(Format-FileSize -Bytes $d.UsedBytes)</td>")
        [void]$html.AppendLine("  <td data-sort=`"$($d.FreeBytes)`">$(Format-FileSize -Bytes $d.FreeBytes)</td>")
        [void]$html.AppendLine("  <td data-sort=`"$($d.PercentUsed)`">$($d.PercentUsed)%</td>")
        [void]$html.AppendLine("  <td><div class=`"drive-bar-outer`"><div class=`"drive-bar-inner $barClass`" style=`"width:$($d.PercentUsed)%`"></div></div></td>")
        [void]$html.AppendLine("</tr>")
    }

    [void]$html.AppendLine('</tbody></table></div>')

    # ---- Category chart ----
    $catChartData = $ScanResult.CategorySummary | ForEach-Object {
        [PSCustomObject]@{ Label = $_.Category; Value = $_.TotalBytes }
    }
    [void]$html.AppendLine('<h2>Storage by File Category</h2>')
    [void]$html.AppendLine((New-BarChartHtml -Data $catChartData -MaxBars 15))

    # ---- Top 10 largest folders chart ----
    $folderChartData = $ScanResult.TopFolders | Select-Object -First 10 | ForEach-Object {
        # Shorten path for chart label
        $label = $_.Path
        if ($label.Length -gt 50) { $label = '...' + $label.Substring($label.Length - 47) }
        [PSCustomObject]@{ Label = $label; Value = $_.SizeBytes }
    }
    [void]$html.AppendLine('<h2>Top 10 Largest Folders</h2>')
    [void]$html.AppendLine((New-BarChartHtml -Data $folderChartData -MaxBars 10))

    # ---- Top 10 largest files chart ----
    $fileChartData = $ScanResult.TopFiles | Select-Object -First 10 | ForEach-Object {
        [PSCustomObject]@{ Label = $_.Name; Value = $_.Length }
    }
    [void]$html.AppendLine('<h2>Top 10 Largest Files</h2>')
    [void]$html.AppendLine((New-BarChartHtml -Data $fileChartData -MaxBars 10))

    [void]$html.AppendLine((Get-HtmlFooter -ScanStart $ScanStart -ScanEnd $ScanEnd -ScannedPaths $ScannedPaths))

    $html.ToString() | Out-File -FilePath $OutputFile -Encoding utf8 -Force
    Write-Verbose "Dashboard report written to $OutputFile"
}

function New-TopFoldersReport {
    <#
    .SYNOPSIS
        Generates the TopFolders.html report with the largest directories.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $ScanResult,
        [Parameter(Mandatory)] [string]$OutputFile,
        [Parameter(Mandatory)] [int]$TopN,
        [Parameter(Mandatory)] [datetime]$ScanStart,
        [Parameter(Mandatory)] [datetime]$ScanEnd,
        [Parameter(Mandatory)] [string[]]$ScannedPaths
    )

    $html = [System.Text.StringBuilder]::new()
    [void]$html.Append((Get-HtmlHeader -Title 'Disk Report - Top Folders'))
    [void]$html.AppendLine((Get-NavHtml -ActivePage 'TopFolders.html'))

    [void]$html.AppendLine('<h1>Top Folders by Size</h1>')
    [void]$html.AppendLine("<p class=`"subtitle`">Showing the $TopN largest directories across all scanned paths.</p>")

    # Chart
    $chartData = $ScanResult.TopFolders | Select-Object -First 15 | ForEach-Object {
        $label = $_.Path
        if ($label.Length -gt 60) { $label = '...' + $label.Substring($label.Length - 57) }
        [PSCustomObject]@{ Label = $label; Value = $_.SizeBytes }
    }
    [void]$html.AppendLine((New-BarChartHtml -Data $chartData -MaxBars 15 -ChartTitle 'Top 15 Folders'))

    # Group by drive letter for drill-down
    $byDrive = $ScanResult.TopFolders | Group-Object { ($_.Path -split '\\')[0] + '\' }

    foreach ($driveGroup in $byDrive) {
        [void]$html.AppendLine("<h2>Drive: $($driveGroup.Name)</h2>")
        [void]$html.AppendLine('<div class="table-wrapper"><table class="sortable"><thead><tr><th>#</th><th>Folder Path</th><th>Size</th></tr></thead><tbody>')

        $rank = 0
        foreach ($folder in ($driveGroup.Group | Sort-Object -Property SizeBytes -Descending)) {
            $rank++
            $sizeClass = Get-SizeColorClass -Bytes $folder.SizeBytes
            $safePath = [System.Net.WebUtility]::HtmlEncode($folder.Path)
            [void]$html.AppendLine("<tr><td class=`"rank`">$rank</td><td title=`"$safePath`">$safePath</td><td class=`"$sizeClass`" data-sort=`"$($folder.SizeBytes)`">$(Format-FileSize -Bytes $folder.SizeBytes)</td></tr>")
        }

        [void]$html.AppendLine('</tbody></table></div>')
    }

    # Access denied section
    if ($ScanResult.AccessDenied.Count -gt 0) {
        [void]$html.AppendLine('<h2>Inaccessible Paths</h2>')
        [void]$html.AppendLine("<p>$($ScanResult.AccessDenied.Count) paths could not be accessed (access denied or error).</p>")
        [void]$html.AppendLine('<div class="access-log">')
        $maxShow = [math]::Min($ScanResult.AccessDenied.Count, 200)
        for ($i = 0; $i -lt $maxShow; $i++) {
            [void]$html.AppendLine("<p>$([System.Net.WebUtility]::HtmlEncode($ScanResult.AccessDenied[$i]))</p>")
        }
        if ($ScanResult.AccessDenied.Count -gt 200) {
            [void]$html.AppendLine("<p><em>... and $($ScanResult.AccessDenied.Count - 200) more.</em></p>")
        }
        [void]$html.AppendLine('</div>')
    }

    [void]$html.AppendLine((Get-HtmlFooter -ScanStart $ScanStart -ScanEnd $ScanEnd -ScannedPaths $ScannedPaths))

    $html.ToString() | Out-File -FilePath $OutputFile -Encoding utf8 -Force
    Write-Verbose "Top Folders report written to $OutputFile"
}

function New-TopFilesReport {
    <#
    .SYNOPSIS
        Generates the TopFiles.html report with the largest individual files.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $ScanResult,
        [Parameter(Mandatory)] [string]$OutputFile,
        [Parameter(Mandatory)] [int]$TopN,
        [Parameter(Mandatory)] [datetime]$ScanStart,
        [Parameter(Mandatory)] [datetime]$ScanEnd,
        [Parameter(Mandatory)] [string[]]$ScannedPaths
    )

    $html = [System.Text.StringBuilder]::new()
    [void]$html.Append((Get-HtmlHeader -Title 'Disk Report - Top Files'))
    [void]$html.AppendLine((Get-NavHtml -ActivePage 'TopFiles.html'))

    [void]$html.AppendLine('<h1>Top Files by Size</h1>')
    [void]$html.AppendLine("<p class=`"subtitle`">Showing the $TopN largest individual files.</p>")

    # Chart
    $chartData = $ScanResult.TopFiles | Select-Object -First 15 | ForEach-Object {
        [PSCustomObject]@{ Label = $_.Name; Value = $_.Length }
    }
    [void]$html.AppendLine((New-BarChartHtml -Data $chartData -MaxBars 15 -ChartTitle 'Top 15 Files'))

    # Table
    [void]$html.AppendLine('<div class="table-wrapper"><table class="sortable"><thead><tr><th>#</th><th>File Name</th><th>Size</th><th>Type</th><th>Last Modified</th><th>Location</th></tr></thead><tbody>')

    $rank = 0
    foreach ($file in $ScanResult.TopFiles) {
        $rank++
        $sizeClass = Get-SizeColorClass -Bytes $file.Length
        $safeName = [System.Net.WebUtility]::HtmlEncode($file.Name)
        $safePath = [System.Net.WebUtility]::HtmlEncode($file.DirectoryName)
        $ext = if ($file.Extension) { $file.Extension.ToUpperInvariant() } else { '-' }
        $modified = $file.LastWriteTime.ToString('yyyy-MM-dd HH:mm')

        [void]$html.AppendLine("<tr><td class=`"rank`">$rank</td><td title=`"$([System.Net.WebUtility]::HtmlEncode($file.FullName))`">$safeName</td><td class=`"$sizeClass`" data-sort=`"$($file.Length)`">$(Format-FileSize -Bytes $file.Length)</td><td>$ext</td><td data-sort=`"$($file.LastWriteTime.ToString('yyyyMMddHHmm'))`">$modified</td><td title=`"$safePath`" style=`"max-width:350px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;`">$safePath</td></tr>")
    }

    [void]$html.AppendLine('</tbody></table></div>')

    [void]$html.AppendLine((Get-HtmlFooter -ScanStart $ScanStart -ScanEnd $ScanEnd -ScannedPaths $ScannedPaths))

    $html.ToString() | Out-File -FilePath $OutputFile -Encoding utf8 -Force
    Write-Verbose "Top Files report written to $OutputFile"
}

function New-FileTypesReport {
    <#
    .SYNOPSIS
        Generates the FileTypes.html report with storage breakdown by extension and category.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $ScanResult,
        [Parameter(Mandatory)] [string]$OutputFile,
        [Parameter(Mandatory)] [datetime]$ScanStart,
        [Parameter(Mandatory)] [datetime]$ScanEnd,
        [Parameter(Mandatory)] [string[]]$ScannedPaths
    )

    $html = [System.Text.StringBuilder]::new()
    [void]$html.Append((Get-HtmlHeader -Title 'Disk Report - File Type Analysis'))
    [void]$html.AppendLine((Get-NavHtml -ActivePage 'FileTypes.html'))

    [void]$html.AppendLine('<h1>File Type Analysis</h1>')
    [void]$html.AppendLine('<p class="subtitle">Storage consumption broken down by file category and extension.</p>')

    # ---- Category chart ----
    $catChartData = $ScanResult.CategorySummary | ForEach-Object {
        [PSCustomObject]@{ Label = $_.Category; Value = $_.TotalBytes }
    }
    [void]$html.AppendLine((New-BarChartHtml -Data $catChartData -MaxBars 15 -ChartTitle 'Storage by Category'))

    # ---- Category table ----
    [void]$html.AppendLine('<h2>Category Summary</h2>')
    [void]$html.AppendLine('<div class="table-wrapper"><table class="sortable"><thead><tr><th>Category</th><th>Total Size</th><th>File Count</th><th>% of Scanned</th></tr></thead><tbody>')

    foreach ($cat in $ScanResult.CategorySummary) {
        $pct = if ($ScanResult.TotalBytes -gt 0) { [math]::Round($cat.TotalBytes / $ScanResult.TotalBytes * 100, 2) } else { 0 }
        $sizeClass = Get-SizeColorClass -Bytes $cat.TotalBytes
        [void]$html.AppendLine("<tr><td><strong>$([System.Net.WebUtility]::HtmlEncode($cat.Category))</strong></td><td class=`"$sizeClass`" data-sort=`"$($cat.TotalBytes)`">$(Format-FileSize -Bytes $cat.TotalBytes)</td><td data-sort=`"$($cat.Count)`">$($cat.Count.ToString('N0'))</td><td data-sort=`"$pct`">${pct}%</td></tr>")
    }

    [void]$html.AppendLine('</tbody></table></div>')

    # ---- Top extensions chart ----
    $topExtData = $ScanResult.ExtensionSummary | Select-Object -First 20 | ForEach-Object {
        $label = if ($_.Extension -and $_.Extension -ne '(none)') { $_.Extension } else { '(no ext)' }
        [PSCustomObject]@{ Label = $label; Value = $_.TotalBytes }
    }
    [void]$html.AppendLine((New-BarChartHtml -Data $topExtData -MaxBars 20 -ChartTitle 'Top 20 Extensions by Size'))

    # ---- Full extension table ----
    [void]$html.AppendLine('<h2>All Extensions (Top 200)</h2>')
    [void]$html.AppendLine('<div class="table-wrapper"><table class="sortable"><thead><tr><th>Extension</th><th>Category</th><th>Total Size</th><th>File Count</th><th>Avg Size</th></tr></thead><tbody>')

    $extRank = 0
    foreach ($ext in ($ScanResult.ExtensionSummary | Select-Object -First 200)) {
        $extRank++
        $sizeClass = Get-SizeColorClass -Bytes $ext.TotalBytes
        $avgSize = if ($ext.Count -gt 0) { $ext.TotalBytes / $ext.Count } else { 0 }
        $extLabel = if ($ext.Extension -and $ext.Extension -ne '(none)') { $ext.Extension } else { '(no ext)' }

        [void]$html.AppendLine("<tr><td><code>$([System.Net.WebUtility]::HtmlEncode($extLabel))</code></td><td>$([System.Net.WebUtility]::HtmlEncode($ext.Category))</td><td class=`"$sizeClass`" data-sort=`"$($ext.TotalBytes)`">$(Format-FileSize -Bytes $ext.TotalBytes)</td><td data-sort=`"$($ext.Count)`">$($ext.Count.ToString('N0'))</td><td data-sort=`"$avgSize`">$(Format-FileSize -Bytes $avgSize)</td></tr>")
    }

    [void]$html.AppendLine('</tbody></table></div>')

    # ---- Per-category drill-down ----
    [void]$html.AppendLine('<h2>Extensions by Category</h2>')

    foreach ($cat in $ScanResult.CategorySummary) {
        $catExts = $ScanResult.ExtensionSummary | Where-Object { $_.Category -eq $cat.Category } | Select-Object -First 30
        if ($catExts.Count -eq 0) { continue }

        [void]$html.AppendLine("<h3>$([System.Net.WebUtility]::HtmlEncode($cat.Category)) &mdash; $(Format-FileSize -Bytes $cat.TotalBytes) ($($cat.Count.ToString('N0')) files)</h3>")

        $catExtChart = $catExts | ForEach-Object {
            [PSCustomObject]@{ Label = $_.Extension; Value = $_.TotalBytes }
        }
        [void]$html.AppendLine((New-BarChartHtml -Data $catExtChart -MaxBars 15))
    }

    [void]$html.AppendLine((Get-HtmlFooter -ScanStart $ScanStart -ScanEnd $ScanEnd -ScannedPaths $ScannedPaths))

    $html.ToString() | Out-File -FilePath $OutputFile -Encoding utf8 -Force
    Write-Verbose "File Types report written to $OutputFile"
}

# ---------------------------------------------------------------------------
# Region: Main Execution
# ---------------------------------------------------------------------------

$scanStart = Get-Date

# Determine root paths
if (-not $RootPath -or $RootPath.Count -eq 0) {
    Write-Verbose 'No RootPath specified - detecting fixed drives...'
    $detectedDrives = @(Get-FixedDrives)
    if ($detectedDrives.Count -eq 0) {
        Write-Error 'No fixed drives detected. Please specify -RootPath explicitly.'
        return
    }

    if ($detectedDrives.Count -eq 1) {
        # Only one drive — use it automatically
        $RootPath = $detectedDrives
        Write-Verbose "Single drive detected: $($RootPath -join ', ')"
    }
    else {
        # Multiple drives — let the user choose
        Write-Host ''
        Write-Host '  Detected fixed drives:' -ForegroundColor Cyan
        Write-Host ''
        for ($i = 0; $i -lt $detectedDrives.Count; $i++) {
            $di = Get-DriveSpaceInfo -DriveRoot $detectedDrives[$i]
            $detail = if ($di) {
                '  Total: {0,-10} Free: {1}' -f (Format-FileSize -Bytes $di.TotalBytes), (Format-FileSize -Bytes $di.FreeBytes)
            } else { '' }
            Write-Host ('  [{0}] {1,-6}{2}' -f ($i + 1), $detectedDrives[$i], $detail) -ForegroundColor White
        }
        Write-Host ('  [{0}] All drives' -f ($detectedDrives.Count + 1)) -ForegroundColor White
        Write-Host ''

        $selectedDrives = $null
        while ($null -eq $selectedDrives) {
            $userInput = Read-Host "  Select drives to scan (e.g. 1  or  1,3  or  $($detectedDrives.Count + 1) for all)"
            $userInput = $userInput.Trim()

            if ($userInput -eq ($detectedDrives.Count + 1).ToString()) {
                $selectedDrives = $detectedDrives
            }
            else {
                $picks = $userInput -split '[,\s]+' |
                         Where-Object { $_ -match '^\d+$' } |
                         ForEach-Object { [int]$_ } |
                         Where-Object { $_ -ge 1 -and $_ -le $detectedDrives.Count } |
                         Select-Object -Unique
                if ($picks.Count -gt 0) {
                    $selectedDrives = $picks | ForEach-Object { $detectedDrives[$_ - 1] }
                }
                else {
                    Write-Host ("  Invalid selection. Enter numbers between 1 and {0}." -f ($detectedDrives.Count + 1)) -ForegroundColor Yellow
                }
            }
        }

        $RootPath = @($selectedDrives)
        Write-Host ''
        Write-Host "  Scanning: $($RootPath -join ', ')" -ForegroundColor Green
        Write-Host ''
    }
}

# Validate root paths
$validPaths = [System.Collections.Generic.List[string]]::new()
foreach ($p in $RootPath) {
    if (Test-Path -LiteralPath $p -PathType Container) {
        $validPaths.Add($p)
    }
    else {
        Write-Warning "Path does not exist or is not a directory: $p"
    }
}
if ($validPaths.Count -eq 0) {
    Write-Error 'No valid paths to scan. Exiting.'
    return
}

# Create output folder
if (-not (Test-Path -LiteralPath $OutputFolder)) {
    try {
        New-Item -Path $OutputFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
        Write-Verbose "Created output folder: $OutputFolder"
    }
    catch {
        Write-Error "Failed to create output folder '$OutputFolder': $_"
        return
    }
}

# Collect drive space info
Write-Verbose 'Collecting drive space information...'
$driveInfoList = @()
$driveLettersSeen = @{}
foreach ($p in $validPaths) {
    $driveLetter = (Split-Path -Qualifier $p -ErrorAction SilentlyContinue)
    if ($driveLetter -and -not $driveLettersSeen.ContainsKey($driveLetter)) {
        $driveLettersSeen[$driveLetter] = $true
        $info = Get-DriveSpaceInfo -DriveRoot "$driveLetter\"
        if ($info) { $driveInfoList += $info }
    }
}

# Run the scan
Write-Verbose "Starting disk scan of $($validPaths.Count) path(s) with TopN=$TopN..."
Write-Verbose "Excluded paths: $($ExcludePaths -join ', ')"

$scanResult = Start-DiskScan -Paths $validPaths -ExcludeList $ExcludePaths -TopN $TopN

$scanEnd = Get-Date

Write-Verbose "Scan complete. Files: $($scanResult.TotalFiles) | Folders: $($scanResult.TotalFolders) | Size: $(Format-FileSize -Bytes $scanResult.TotalBytes)"
Write-Verbose "Access-denied entries: $($scanResult.AccessDenied.Count)"

# Generate reports
$reportParams = @{
    ScanResult   = $scanResult
    ScanStart    = $scanStart
    ScanEnd      = $scanEnd
    ScannedPaths = [string[]]$validPaths
}

Write-Verbose 'Generating Dashboard report...'
New-DashboardReport @reportParams -OutputFile (Join-Path $OutputFolder 'Dashboard.html') -DriveInfo $driveInfoList

Write-Verbose 'Generating Top Folders report...'
New-TopFoldersReport @reportParams -OutputFile (Join-Path $OutputFolder 'TopFolders.html') -TopN $TopN

Write-Verbose 'Generating Top Files report...'
New-TopFilesReport @reportParams -OutputFile (Join-Path $OutputFolder 'TopFiles.html') -TopN $TopN

Write-Verbose 'Generating File Types report...'
New-FileTypesReport @reportParams -OutputFile (Join-Path $OutputFolder 'FileTypes.html')

# Summary output
$duration = $scanEnd - $scanStart
$durationStr = '{0}h {1}m {2}s' -f [int]$duration.TotalHours, $duration.Minutes, $duration.Seconds

Write-Output ''
Write-Output '============================================='
Write-Output '  Disk Storage Report - Complete'
Write-Output '============================================='
Write-Output "  Scanned Paths  : $($validPaths -join ', ')"
Write-Output "  Total Files    : $($scanResult.TotalFiles.ToString('N0'))"
Write-Output "  Total Folders  : $($scanResult.TotalFolders.ToString('N0'))"
Write-Output "  Total Size     : $(Format-FileSize -Bytes $scanResult.TotalBytes)"
Write-Output "  Access Denied  : $($scanResult.AccessDenied.Count) paths"
Write-Output "  Scan Duration  : $durationStr"
Write-Output "  Report Folder  : $OutputFolder"
Write-Output '============================================='
Write-Output ''
Write-Output "  Reports generated:"
Write-Output "    - Dashboard.html    (summary with charts)"
Write-Output "    - TopFolders.html   (largest directories)"
Write-Output "    - TopFiles.html     (largest files)"
Write-Output "    - FileTypes.html    (extension & category analysis)"
Write-Output ''

# Open the dashboard in the default browser (non-blocking)
$dashboardPath = Join-Path $OutputFolder 'Dashboard.html'
if (Test-Path -LiteralPath $dashboardPath) {
    try {
        Start-Process -FilePath $dashboardPath -ErrorAction SilentlyContinue
    }
    catch {
        Write-Verbose "Could not auto-open browser: $_"
    }
}

<#
# ===========================================================================
# STORAGE, DISTRIBUTION & SECURITY RECOMMENDATIONS
# ===========================================================================
#
# 1. RECOMMENDED HOSTING: GitHub Repository (Public or Private)
#    ---------------------------------------------------------------
#    Use a dedicated GitHub repository rather than a Gist. Reasons:
#
#    - A repository supports versioning with tags (v1.0.0, v1.1.0), branching,
#      pull requests, and issue tracking -- all critical for a production script.
#    - Gists are fine for throwaway snippets but lack branch protection, CI/CD
#      pipelines, contributor access controls, and discoverability.
#    - A repo lets you add a README, LICENSE (MIT recommended), CHANGELOG,
#      and Pester tests alongside the script.
#    - GitHub Actions can run PSScriptAnalyzer and Pester on every push.
#    - Example repo structure:
#        Get-DiskStorageReport/
#          Get-DiskStorageReport.ps1
#          README.md
#          LICENSE
#          CHANGELOG.md
#          tests/
#            Get-DiskStorageReport.Tests.ps1
#          .github/
#            workflows/
#              ci.yml
#
# 2. ONE-LINE DOWNLOAD & RUN FROM A WINDOWS MACHINE
#    ---------------------------------------------------------------
#    Option A - Download then run (RECOMMENDED - allows inspection first):
#
#      Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/YOUR_USER/Get-DiskStorageReport/main/Get-DiskStorageReport.ps1' -OutFile "$env:TEMP\Get-DiskStorageReport.ps1"; & "$env:TEMP\Get-DiskStorageReport.ps1"
#
#    Option B - Invoke-Expression (iex) one-liner (LESS SECURE):
#
#      & ([scriptblock]::Create((Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/YOUR_USER/Get-DiskStorageReport/main/Get-DiskStorageReport.ps1' -UseBasicParsing).Content))
#
#    WARNING: iex + web download is convenient but skips inspection.
#    Always use HTTPS URLs and pin to a tagged release commit, never to
#    a mutable branch head in production.
#
#    Option C - PowerShell Gallery (BEST for wide distribution):
#
#      Install-Script -Name Get-DiskStorageReport
#      Get-DiskStorageReport.ps1
#
#      Publishing to the PowerShell Gallery provides discoverability,
#      dependency management, and versioning. Requires a NuGet API key.
#
# 3. SECURITY CONSIDERATIONS
#    ---------------------------------------------------------------
#    a) EXECUTION POLICY
#       - Set to RemoteSigned on target machines (the safe default).
#       - NEVER advise users to set Unrestricted or Bypass globally.
#       - For a one-off run: powershell -ExecutionPolicy Bypass -File .\Get-DiskStorageReport.ps1
#         This scopes the bypass to the single process.
#
#    b) CODE SIGNING
#       - For enterprise/RMM deployment, sign the script with a trusted
#         code-signing certificate (internal CA or a commercial cert).
#       - Set-AuthenticodeSignature -FilePath .\Get-DiskStorageReport.ps1 -Certificate $cert
#       - This ensures RemoteSigned machines accept it without prompts
#         and provides tamper evidence.
#
#    c) HTTPS ONLY
#       - All download URLs must be HTTPS. GitHub raw URLs are HTTPS by default.
#       - If self-hosting, enforce TLS 1.2+:
#         [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
#
#    d) INTEGRITY VERIFICATION
#       - Publish a SHA256 hash alongside each release.
#       - Users can verify:  (Get-FileHash .\Get-DiskStorageReport.ps1).Hash
#       - Consider signing commits with GPG keys on GitHub.
#
#    e) NO CREDENTIALS / NO DATA EXFILTRATION
#       - This script collects ONLY file metadata (names, sizes, dates).
#       - It does NOT read file contents, transmit data over the network,
#         or require any credentials.
#       - If extending the script, never add network upload without explicit
#         user consent and secure transport.
#
# 4. GITHUB REPO vs. GIST vs. OTHER OPTIONS
#    ---------------------------------------------------------------
#    | Option              | Pros                           | Cons                            |
#    |---------------------|--------------------------------|---------------------------------|
#    | GitHub Repo         | Full CI/CD, issues, PRs,       | Slightly more setup             |
#    |                     | branch protection, releases    |                                 |
#    | GitHub Gist         | Quick to create, embeddable    | No CI/CD, no issues, limited    |
#    |                     |                                | visibility for multi-file       |
#    | PowerShell Gallery  | Built-in Install-Script,       | Requires NuGet key, publishing  |
#    |                     | versioning, dependency mgmt    | process, review time            |
#    | Azure DevOps Repo   | Enterprise AAD integration,    | Less discoverable publicly      |
#    |                     | pipelines, board integration   |                                 |
#    | Self-hosted (IIS)   | Full control                   | Maintenance burden, TLS setup   |
#    ---------------------------------------------------------------
#
#    RECOMMENDATION: Use a PUBLIC GitHub repository for the script itself
#    (it contains no secrets) and ALSO publish to the PowerShell Gallery
#    for maximum reach. Use GitHub Releases with semantic version tags
#    for each stable version.
#
# ===========================================================================
#>
