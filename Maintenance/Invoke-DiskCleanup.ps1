#Requires -Version 5.1

<#
.SYNOPSIS
    Cleans common Windows cache and temporary folders to reclaim disk space.

.DESCRIPTION
    Invoke-DiskCleanup targets a pre-approved list of system and per-user cache
    directories (browser caches, crash dumps, temp files, thumbnail caches, etc.)
    and safely removes their contents. It never deletes the folder itself, only
    the files and subfolders within it.

    System-wide targets require administrator privileges; if the script is not
    elevated those targets are skipped with a warning. Locked files are skipped
    silently and counted separately.

    An optional set of targets (Windows Update download cache, Office file cache,
    npm/pip/NuGet caches) can be included with -IncludeOptional. The Windows
    Update target automatically stops and restarts the wuauserv service.

    A self-contained HTML report summarising every target is generated unless
    -NoReport is specified.

.PARAMETER IncludeOptional
    Also clean optional targets: SoftwareDistribution\Download (stops wuauserv),
    Office 16.0 file cache, npm-cache, pip cache, and NuGet cache.

.PARAMETER AllUsers
    Clean per-user cache folders for ALL user profiles on the machine. Requires
    administrator privileges. Without this switch only the current user's folders
    are cleaned.

.PARAMETER OutputFolder
    Directory where the HTML summary report is saved. Defaults to
    $env:USERPROFILE\DiskCleanup. Created automatically if it does not exist.

.PARAMETER NoReport
    Skip generating the HTML summary report.

.PARAMETER Force
    Skip the interactive confirmation prompt before deleting. Useful for
    RMM / scheduled-task / unattended execution.

.EXAMPLE
    .\Invoke-DiskCleanup.ps1

    Cleans default targets for the current user only, prompts for confirmation,
    and generates an HTML report.

.EXAMPLE
    .\Invoke-DiskCleanup.ps1 -AllUsers -IncludeOptional -WhatIf

    Simulates a full cleanup of all users and optional targets without deleting
    anything. Shows what would be removed.

.EXAMPLE
    .\Invoke-DiskCleanup.ps1 -Force -NoReport -AllUsers -IncludeOptional

    Runs a full unattended cleanup with no confirmation and no HTML report.
    Suitable for RMM / automated deployment.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [switch]$IncludeOptional,

    [switch]$AllUsers,

    [string]$OutputFolder = 'C:\Temp',

    [switch]$NoReport,

    [switch]$Force
)

# ── Script metadata ──────────────────────────────────────────────────────────
$ScriptVersion = '1.0.0'
$StartTime = Get-Date

# ── Helper: Admin check ─────────────────────────────────────────────────────
function Test-IsAdmin {
    <#
    .SYNOPSIS
        Returns $true when the current process is elevated.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object -TypeName Security.Principal.WindowsPrincipal -ArgumentList $identity
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

$IsAdmin = Test-IsAdmin
Write-Verbose "Running as administrator: $IsAdmin"

# ── Helper: Measure folder ──────────────────────────────────────────────────
function Measure-FolderSize {
    <#
    .SYNOPSIS
        Returns total byte count and file count for a directory tree.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $bytes = [long]0
    $count = [int]0

    try {
        $files = [System.IO.Directory]::GetFiles($Path, '*', [System.IO.SearchOption]::AllDirectories)
        foreach ($file in $files) {
            try {
                $info = New-Object -TypeName System.IO.FileInfo -ArgumentList $file
                $bytes += $info.Length
                $count++
            }
            catch {
                # Inaccessible file — skip
            }
        }
    }
    catch {
        Write-Verbose "Could not enumerate '$Path': $_"
    }

    return [PSCustomObject]@{
        Bytes     = $bytes
        FileCount = $count
    }
}

# ── Helper: Format bytes ────────────────────────────────────────────────────
function Format-ByteSize {
    <#
    .SYNOPSIS
        Converts a byte count to a human-readable string (KB / MB / GB / TB).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [long]$Bytes
    )

    switch ($Bytes) {
        { $_ -ge 1TB } { return '{0:N2} TB' -f ($Bytes / 1TB) }
        { $_ -ge 1GB } { return '{0:N2} GB' -f ($Bytes / 1GB) }
        { $_ -ge 1MB } { return '{0:N2} MB' -f ($Bytes / 1MB) }
        { $_ -ge 1KB } { return '{0:N2} KB' -f ($Bytes / 1KB) }
        default        { return '{0} B'      -f $Bytes         }
    }
}

# ── Resolve and validate output folder ───────────────────────────────────────
# Normalise: treat empty/whitespace as if not supplied and fall back to default
if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
    $OutputFolder = 'C:\Temp'
    Write-Verbose "OutputFolder was empty — defaulting to '$OutputFolder'"
}

if (-not $NoReport -and -not $WhatIfPreference) {
    if (-not (Test-Path -LiteralPath $OutputFolder -PathType Container)) {
        try {
            New-Item -Path $OutputFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
            Write-Verbose "Created output folder: $OutputFolder"
        }
        catch {
            Write-Warning "Cannot create output folder '$OutputFolder': $_. Reports will be skipped."
            $NoReport = $true
        }
    }
}

# ── Build target list ────────────────────────────────────────────────────────
# Each target: Path, Description, IsSystem, IsOptional, Filter (optional glob)

$Targets = [System.Collections.Generic.List[PSCustomObject]]::new()

# --- System-wide (default) ---------------------------------------------------
$SystemDefaults = @(
    @{ Path = 'C:\Windows\Temp';              Description = 'Windows Temp' }
    @{ Path = 'C:\Windows\CbsTemp';           Description = 'CBS Temp' }
    @{ Path = 'C:\Windows\Prefetch';          Description = 'Prefetch' }
    @{ Path = 'C:\Windows\LiveKernelReports'; Description = 'Live Kernel Reports' }
    @{ Path = 'C:\Windows\Logs\CBS';          Description = 'CBS Logs' }
)

foreach ($entry in $SystemDefaults) {
    $Targets.Add([PSCustomObject]@{
        Path        = $entry.Path
        Description = $entry.Description
        IsSystem    = $true
        IsOptional  = $false
        Filter      = $null
    })
}

# --- System-wide (optional) ---------------------------------------------------
if ($IncludeOptional) {
    $Targets.Add([PSCustomObject]@{
        Path        = 'C:\Windows\SoftwareDistribution\Download'
        Description = 'Windows Update Download Cache'
        IsSystem    = $true
        IsOptional  = $true
        Filter      = $null
    })
}

# --- Resolve user profile base paths -----------------------------------------
$UserLocalAppDataPaths = [System.Collections.Generic.List[string]]::new()

if ($AllUsers) {
    if (-not $IsAdmin) {
        Write-Warning '-AllUsers requires administrator privileges. Falling back to current user only.'
    }
    else {
        $profileRoot = Split-Path -Path $env:USERPROFILE -Parent
        $profiles = Get-ChildItem -Path $profileRoot -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notin @('Public', 'Default', 'Default User', 'All Users') }
        foreach ($profile in $profiles) {
            $localAppData = Join-Path -Path $profile.FullName -ChildPath 'AppData\Local'
            if (Test-Path -Path $localAppData -PathType Container) {
                $UserLocalAppDataPaths.Add($localAppData)
            }
        }
    }
}

# Always include the current user
if ($UserLocalAppDataPaths.Count -eq 0 -or (-not $AllUsers)) {
    if ($env:LOCALAPPDATA -and -not $UserLocalAppDataPaths.Contains($env:LOCALAPPDATA)) {
        $UserLocalAppDataPaths.Add($env:LOCALAPPDATA)
    }
}

# --- Per-user default targets -------------------------------------------------
$UserDefaultRelPaths = @(
    @{ Rel = 'Temp';                                                  Desc = 'User Temp' }
    @{ Rel = 'Microsoft\Windows\INetCache';                           Desc = 'Internet Cache' }
    @{ Rel = 'Microsoft\Windows\WER';                                 Desc = 'Windows Error Reporting' }
    @{ Rel = 'CrashDumps';                                            Desc = 'Crash Dumps' }
    @{ Rel = 'D3DSCache';                                             Desc = 'D3D Shader Cache' }
    @{ Rel = 'NVIDIA\DXCache';                                        Desc = 'NVIDIA DX Cache' }
    @{ Rel = 'AMD\DXCache';                                           Desc = 'AMD DX Cache' }
    @{ Rel = 'Google\Chrome\User Data\Default\Cache';                 Desc = 'Chrome Cache' }
    @{ Rel = 'Google\Chrome\User Data\Default\Code Cache';            Desc = 'Chrome Code Cache' }
    @{ Rel = 'Microsoft\Edge\User Data\Default\Cache';                Desc = 'Edge Cache' }
    @{ Rel = 'Microsoft\Edge\User Data\Default\Code Cache';           Desc = 'Edge Code Cache' }
    @{ Rel = 'Microsoft\Teams\Cache';                                 Desc = 'Teams Cache' }
    @{ Rel = 'Microsoft\Teams\Application Cache';                     Desc = 'Teams App Cache' }
    @{ Rel = 'Microsoft\Teams\blob_storage';                          Desc = 'Teams Blob Storage' }
    @{ Rel = 'Microsoft\Teams\GPUCache';                              Desc = 'Teams GPU Cache' }
    @{ Rel = 'Microsoft\Teams\IndexedDB';                             Desc = 'Teams IndexedDB' }
    @{ Rel = 'Microsoft\Teams\Local Storage';                         Desc = 'Teams Local Storage' }
)

# Filtered per-user targets (Explorer thumbnails / icon caches)
$UserFilteredRelPaths = @(
    @{ Rel = 'Microsoft\Windows\Explorer'; Desc = 'Thumbnail & Icon Caches'; Filter = 'thumbcache_*.db','iconcache_*.db' }
)

# Firefox uses a glob across profiles
$FirefoxProfileGlob = @(
    @{ Rel = 'Mozilla\Firefox\Profiles'; Desc = 'Firefox Cache'; SubGlob = 'cache2' }
)

# --- Per-user optional targets ------------------------------------------------
$UserOptionalRelPaths = @()
if ($IncludeOptional) {
    $UserOptionalRelPaths = @(
        @{ Rel = 'Microsoft\Office\16.0\OfficeFileCache'; Desc = 'Office File Cache' }
        @{ Rel = 'npm-cache';                             Desc = 'npm Cache' }
        @{ Rel = 'pip\cache';                             Desc = 'pip Cache' }
        @{ Rel = 'nuget\cache';                           Desc = 'NuGet Cache' }
    )
}

# Build actual path entries for every user
foreach ($localAppData in $UserLocalAppDataPaths) {
    $userName = Split-Path -Path (Split-Path -Path (Split-Path -Path $localAppData -Parent) -Parent) -Leaf

    foreach ($entry in $UserDefaultRelPaths) {
        $fullPath = Join-Path -Path $localAppData -ChildPath $entry.Rel
        $Targets.Add([PSCustomObject]@{
            Path        = $fullPath
            Description = "$($entry.Desc) ($userName)"
            IsSystem    = $false
            IsOptional  = $false
            Filter      = $null
        })
    }

    foreach ($entry in $UserFilteredRelPaths) {
        $fullPath = Join-Path -Path $localAppData -ChildPath $entry.Rel
        $Targets.Add([PSCustomObject]@{
            Path        = $fullPath
            Description = "$($entry.Desc) ($userName)"
            IsSystem    = $false
            IsOptional  = $false
            Filter      = $entry.Filter
        })
    }

    # Firefox profiles — resolve the glob
    foreach ($entry in $FirefoxProfileGlob) {
        $profilesRoot = Join-Path -Path $localAppData -ChildPath $entry.Rel
        if (Test-Path -Path $profilesRoot -PathType Container) {
            $profileDirs = Get-ChildItem -Path $profilesRoot -Directory -ErrorAction SilentlyContinue
            foreach ($profileDir in $profileDirs) {
                $cachePath = Join-Path -Path $profileDir.FullName -ChildPath $entry.SubGlob
                $Targets.Add([PSCustomObject]@{
                    Path        = $cachePath
                    Description = "$($entry.Desc) [$($profileDir.Name)] ($userName)"
                    IsSystem    = $false
                    IsOptional  = $false
                    Filter      = $null
                })
            }
        }
    }

    foreach ($entry in $UserOptionalRelPaths) {
        $fullPath = Join-Path -Path $localAppData -ChildPath $entry.Rel
        $Targets.Add([PSCustomObject]@{
            Path        = $fullPath
            Description = "$($entry.Desc) ($userName)"
            IsSystem    = $false
            IsOptional  = $true
            Filter      = $null
        })
    }
}

# ── Pre-scan and confirmation ────────────────────────────────────────────────
Write-Verbose "Total cleanup targets: $($Targets.Count)"

# Quick pre-scan for the confirmation prompt
$existingTargets = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($target in $Targets) {
    if ($target.IsSystem -and -not $IsAdmin) {
        continue
    }
    if (Test-Path -Path $target.Path -PathType Container) {
        $existingTargets.Add($target)
    }
}

if ($existingTargets.Count -eq 0) {
    Write-Host 'No cleanup targets found on this system.' -ForegroundColor Yellow
    return
}

# Show confirmation unless -Force or -WhatIf
if (-not $Force -and -not $WhatIfPreference) {
    Write-Host "`n=== Invoke-DiskCleanup v$ScriptVersion ===" -ForegroundColor Cyan
    Write-Host "The following $($existingTargets.Count) target(s) will be cleaned:`n" -ForegroundColor White

    foreach ($target in $existingTargets) {
        $tag = ''
        if ($target.IsOptional) { $tag = ' [optional]' }
        if ($target.IsSystem)   { $tag += ' [system]' }
        Write-Host "  - $($target.Description)$tag" -ForegroundColor Gray
        Write-Host "    $($target.Path)" -ForegroundColor DarkGray
    }

    Write-Host ''
    $answer = Read-Host -Prompt 'Are you sure? (Y/N)'
    if ($answer -notin @('Y', 'y', 'Yes', 'yes')) {
        Write-Host 'Operation cancelled by user.' -ForegroundColor Yellow
        return
    }
}

# ── Process targets ──────────────────────────────────────────────────────────
$Results = [System.Collections.Generic.List[PSCustomObject]]::new()
$targetIndex = 0
$totalTargets = $Targets.Count

foreach ($target in $Targets) {
    $targetIndex++
    $percentComplete = [math]::Round(($targetIndex / $totalTargets) * 100)

    Write-Progress -Activity 'Invoke-DiskCleanup' `
        -Status "$($target.Description) ($targetIndex of $totalTargets)" `
        -PercentComplete $percentComplete

    $result = [PSCustomObject]@{
        FolderPath   = $target.Path
        Description  = $target.Description
        BytesBefore  = [long]0
        FilesDeleted = [int]0
        BytesFreed   = [long]0
        FilesSkipped = [int]0
        Status       = 'NotFound'
    }

    # --- Skip system targets if not admin ---
    if ($target.IsSystem -and -not $IsAdmin) {
        $result.Status = 'Skipped'
        $Results.Add($result)
        Write-Warning "Skipped (not admin): $($target.Path)"
        continue
    }

    # --- Skip non-existent folders ---
    if (-not (Test-Path -Path $target.Path -PathType Container)) {
        $Results.Add($result)
        Write-Verbose "Not found: $($target.Path)"
        continue
    }

    # --- Handle SoftwareDistribution\Download (stop/start wuauserv) ---
    $stoppedWuauserv = $false
    if ($target.Path -eq 'C:\Windows\SoftwareDistribution\Download') {
        Write-Verbose 'Stopping Windows Update service (wuauserv)...'
        try {
            Stop-Service -Name wuauserv -Force -ErrorAction Stop
            $stoppedWuauserv = $true
            Write-Verbose 'wuauserv stopped successfully.'
        }
        catch {
            Write-Warning "Could not stop wuauserv — skipping SoftwareDistribution\Download. Error: $_"
            $result.Status = 'Error'
            $Results.Add($result)
            continue
        }
    }

    # --- Measure before ---
    $before = Measure-FolderSize -Path $target.Path
    $result.BytesBefore = $before.Bytes
    Write-Verbose "Before: $($target.Path) — $(Format-ByteSize -Bytes $before.Bytes) in $($before.FileCount) files"

    # --- Determine items to remove ---
    $itemsToRemove = $null

    try {
        if ($null -ne $target.Filter) {
            # Filtered removal (e.g. thumbcache_*.db, iconcache_*.db)
            $itemsToRemove = @()
            foreach ($pattern in $target.Filter) {
                $matched = Get-ChildItem -Path $target.Path -Filter $pattern -File -ErrorAction SilentlyContinue
                if ($matched) {
                    $itemsToRemove += $matched
                }
            }
        }
        else {
            # Full folder contents
            $itemsToRemove = Get-ChildItem -Path $target.Path -Force -ErrorAction SilentlyContinue
        }
    }
    catch {
        Write-Warning "Error enumerating '$($target.Path)': $_"
        $result.Status = 'Error'
        $Results.Add($result)
        if ($stoppedWuauserv) {
            try { Start-Service -Name wuauserv -ErrorAction SilentlyContinue } catch {}
        }
        continue
    }

    if ($null -eq $itemsToRemove -or @($itemsToRemove).Count -eq 0) {
        $result.Status = 'Cleaned'
        $Results.Add($result)
        Write-Host "[  OK  ] $($target.Description) — already empty" -ForegroundColor Green
        if ($stoppedWuauserv) {
            try { Start-Service -Name wuauserv -ErrorAction SilentlyContinue } catch {}
        }
        continue
    }

    # --- Delete items ---
    $deletedCount = 0
    $skippedCount = 0
    $deletedBytes = [long]0

    foreach ($item in $itemsToRemove) {
        $itemPath = $item.FullName

        if ($PSCmdlet.ShouldProcess($itemPath, 'Remove')) {
            try {
                if ($item.PSIsContainer) {
                    # Measure folder size before removing
                    $folderSize = Measure-FolderSize -Path $itemPath
                    Remove-Item -Path $itemPath -Recurse -Force -ErrorAction Stop
                    $deletedCount += $folderSize.FileCount
                    $deletedBytes += $folderSize.Bytes
                }
                else {
                    $fileLen = $item.Length
                    Remove-Item -Path $itemPath -Force -ErrorAction Stop
                    $deletedCount++
                    $deletedBytes += $fileLen
                }
            }
            catch {
                $skippedCount++
                Write-Verbose "Locked/skipped: $itemPath — $_"
            }
        }
        else {
            # WhatIf mode — count as would-be-deleted for reporting
            if ($item.PSIsContainer) {
                $folderSize = Measure-FolderSize -Path $itemPath
                $deletedCount += $folderSize.FileCount
                $deletedBytes += $folderSize.Bytes
            }
            else {
                $deletedCount++
                $deletedBytes += $item.Length
            }
        }
    }

    $result.FilesDeleted = $deletedCount
    $result.BytesFreed   = $deletedBytes
    $result.FilesSkipped  = $skippedCount
    $result.Status        = 'Cleaned'

    $Results.Add($result)

    # --- Per-target result line ---
    $freedStr  = Format-ByteSize -Bytes $deletedBytes
    $whatIfTag  = if ($WhatIfPreference) { ' (WhatIf)' } else { '' }
    $skipNote   = if ($skippedCount -gt 0) { ", $skippedCount skipped" } else { '' }

    Write-Host "[  OK  ] $($target.Description) — $deletedCount files, $freedStr freed${skipNote}${whatIfTag}" -ForegroundColor Green

    # --- Restart wuauserv if we stopped it ---
    if ($stoppedWuauserv) {
        Write-Verbose 'Restarting Windows Update service (wuauserv)...'
        try {
            Start-Service -Name wuauserv -ErrorAction Stop
            Write-Verbose 'wuauserv restarted successfully.'
        }
        catch {
            Write-Warning "Could not restart wuauserv: $_. You may need to start it manually."
        }
    }
}

Write-Progress -Activity 'Invoke-DiskCleanup' -Completed

# ── Summary ──────────────────────────────────────────────────────────────────
$EndTime       = Get-Date
$Duration      = $EndTime - $StartTime
$TotalFreed    = ($Results | Measure-Object -Property BytesFreed   -Sum).Sum
$TotalDeleted  = ($Results | Measure-Object -Property FilesDeleted -Sum).Sum
$TotalSkipped  = ($Results | Measure-Object -Property FilesSkipped -Sum).Sum

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  Disk Cleanup Complete" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Total freed  : $(Format-ByteSize -Bytes $TotalFreed)" -ForegroundColor White
Write-Host "  Files deleted : $TotalDeleted" -ForegroundColor White
Write-Host "  Files skipped : $TotalSkipped" -ForegroundColor White
Write-Host "  Duration      : $($Duration.ToString('mm\:ss'))" -ForegroundColor White
if ($WhatIfPreference) {
    Write-Host "  Mode          : WhatIf (no files were actually deleted)" -ForegroundColor Yellow
}
Write-Host "========================================`n" -ForegroundColor Cyan

# ── HTML Report ──────────────────────────────────────────────────────────────
if (-not $NoReport -and -not $WhatIfPreference) {
    Write-Verbose 'Generating HTML report...'

    $timestamp   = $EndTime.ToString('yyyy-MM-dd HH:mm:ss')
    $fileStamp   = $EndTime.ToString('yyyyMMdd_HHmmss')
    $reportFile  = Join-Path -Path $OutputFolder -ChildPath "CleanupReport_$fileStamp.html"
    $computerName = $env:COMPUTERNAME

    # Build table rows
    $tableRows = [System.Text.StringBuilder]::new()

    foreach ($r in $Results) {
        $statusColor = switch ($r.Status) {
            'Cleaned'  { '#27ae60' }
            'Skipped'  { '#f39c12' }
            'NotFound' { '#95a5a6' }
            'Error'    { '#e74c3c' }
            default    { '#95a5a6' }
        }

        [void]$tableRows.AppendLine("            <tr>")
        [void]$tableRows.AppendLine("                <td style=`"word-break:break-all;`">$([System.Web.HttpUtility]::HtmlEncode($r.FolderPath))</td>")
        [void]$tableRows.AppendLine("                <td>$([System.Web.HttpUtility]::HtmlEncode($r.Description))</td>")
        [void]$tableRows.AppendLine("                <td style=`"text-align:right;`">$(Format-ByteSize -Bytes $r.BytesBefore)</td>")
        [void]$tableRows.AppendLine("                <td style=`"text-align:right;`">$(Format-ByteSize -Bytes $r.BytesFreed)</td>")
        [void]$tableRows.AppendLine("                <td style=`"text-align:right;`">$($r.FilesDeleted)</td>")
        [void]$tableRows.AppendLine("                <td style=`"text-align:right;`">$($r.FilesSkipped)</td>")
        [void]$tableRows.AppendLine("                <td style=`"text-align:center;`"><span style=`"background:$statusColor;color:#fff;padding:2px 10px;border-radius:4px;font-size:0.85em;`">$($r.Status)</span></td>")
        [void]$tableRows.AppendLine("            </tr>")
    }

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Disk Cleanup Report — $computerName</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f0f2f5; color: #333; padding: 24px; }
        .container { max-width: 1200px; margin: 0 auto; }
        .header-card {
            background: linear-gradient(135deg, #2c3e50, #3498db);
            color: #fff; border-radius: 10px; padding: 28px 32px; margin-bottom: 24px;
            display: flex; flex-wrap: wrap; justify-content: space-between; align-items: center;
        }
        .header-card h1 { font-size: 1.6em; margin-bottom: 4px; }
        .header-card .subtitle { font-size: 0.9em; opacity: 0.85; }
        .stat-grid { display: flex; flex-wrap: wrap; gap: 20px; margin-top: 12px; }
        .stat-box { background: rgba(255,255,255,0.15); border-radius: 8px; padding: 14px 22px; min-width: 140px; text-align: center; }
        .stat-box .value { font-size: 1.6em; font-weight: 700; }
        .stat-box .label { font-size: 0.8em; opacity: 0.8; margin-top: 2px; }
        table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 4px rgba(0,0,0,0.08); }
        thead { background: #2c3e50; color: #fff; }
        th { padding: 12px 14px; text-align: left; font-weight: 600; font-size: 0.85em; text-transform: uppercase; letter-spacing: 0.5px; }
        td { padding: 10px 14px; border-bottom: 1px solid #ecf0f1; font-size: 0.88em; }
        tr:last-child td { border-bottom: none; }
        tr:hover { background: #f8f9fa; }
        .footer { text-align: center; margin-top: 20px; color: #999; font-size: 0.8em; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header-card">
            <div>
                <h1>Disk Cleanup Report</h1>
                <div class="subtitle">$computerName &mdash; $timestamp</div>
            </div>
            <div class="stat-grid">
                <div class="stat-box">
                    <div class="value">$(Format-ByteSize -Bytes $TotalFreed)</div>
                    <div class="label">Space Freed</div>
                </div>
                <div class="stat-box">
                    <div class="value">$TotalDeleted</div>
                    <div class="label">Files Deleted</div>
                </div>
                <div class="stat-box">
                    <div class="value">$TotalSkipped</div>
                    <div class="label">Files Skipped</div>
                </div>
            </div>
        </div>

        <table>
            <thead>
                <tr>
                    <th>Path</th>
                    <th>Description</th>
                    <th style="text-align:right;">Size Before</th>
                    <th style="text-align:right;">Freed</th>
                    <th style="text-align:right;">Deleted</th>
                    <th style="text-align:right;">Skipped</th>
                    <th style="text-align:center;">Status</th>
                </tr>
            </thead>
            <tbody>
$($tableRows.ToString())
            </tbody>
        </table>

        <div class="footer">
            Invoke-DiskCleanup v$ScriptVersion &mdash; completed in $($Duration.ToString('mm\:ss'))
        </div>
    </div>
</body>
</html>
"@

    try {
        Set-Content -Path $reportFile -Value $html -Encoding UTF8 -Force -ErrorAction Stop
        Write-Host "Report saved: $reportFile" -ForegroundColor Cyan

        # Open in default browser
        try {
            Start-Process -FilePath $reportFile -ErrorAction SilentlyContinue
        }
        catch {
            Write-Verbose "Could not open report in browser: $_"
        }
    }
    catch {
        Write-Warning "Could not write report to '$reportFile': $_"
    }
}
