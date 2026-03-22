#Requires -Version 5.1
<#
.SYNOPSIS
    Launcher script for the nakamas-it/powershell-scripts repository.

.DESCRIPTION
    Run.ps1 is the single entry point for the nakamas-it/powershell-scripts repo.
    It downloads scripts.json from the repository to discover available scripts,
    presents an interactive numbered menu (or accepts a script name directly),
    then downloads and executes the chosen script in a temporary location.

    Users can pull and run this launcher with a single one-liner:
        irm https://raw.githubusercontent.com/nakamas-it/powershell-scripts/main/Run.ps1 | iex

.PARAMETER ScriptName
    Run a specific script by name, bypassing the interactive menu.
    The name must match a script entry in scripts.json exactly (case-insensitive).

.PARAMETER Tag
    The Git tag or branch to pull scripts from. Defaults to 'main'.

.PARAMETER List
    Print all available scripts in a formatted table and exit without running anything.

.PARAMETER Search
    Filter the available scripts by name, description, or tag. The search is
    case-insensitive and matches partial strings. Matching results are shown
    as a numbered menu for selection.

.EXAMPLE
    .\Run.ps1
    Launches the interactive menu, listing all available scripts for selection.

.EXAMPLE
    .\Run.ps1 -ScriptName "Get-DiskStorageReport" -Tag "v1.2.0"
    Downloads and runs Get-DiskStorageReport.ps1 from the v1.2.0 tag without
    showing the interactive menu.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Position = 0)]
    [string]$ScriptName,

    [Parameter()]
    [string]$Tag = 'main',

    [Parameter()]
    [switch]$List,

    [Parameter()]
    [string]$Search
)

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$BaseRawUrl = "https://raw.githubusercontent.com/nakamas-it/powershell-scripts/$Tag"
$ScriptsJsonUrl = "$BaseRawUrl/scripts.json"

# ---------------------------------------------------------------------------
# Helper — Write-Banner
# ---------------------------------------------------------------------------
function Write-Banner {
    <#
    .SYNOPSIS
        Prints a styled header banner to the console.
    #>
    [CmdletBinding()]
    param()

    $banner = @"

  _   _   _   _  __   _   __  __   _   ___     ___ _____
 | \ | | /_\ | |/ /  /_\ |  \/  | /_\ / __|   |_ _|_   _|
 |  \| |/ _ \| ' <  / _ \| |\/| |/ _ \\__ \    | |  | |
 |_|\__/_/ \_\_|\_\/_/ \_\_|  |_/_/ \_\|___/   |___| |_|

  PowerShell Script Runner
"@
    Write-Host $banner -ForegroundColor Cyan
    Write-Host "  Tag: $Tag" -ForegroundColor DarkGray
    Write-Host ""
}

# ---------------------------------------------------------------------------
# Helper — Get-ScriptCatalog
# ---------------------------------------------------------------------------
function Get-ScriptCatalog {
    <#
    .SYNOPSIS
        Downloads and parses scripts.json from the repository.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Url
    )

    Write-Verbose "Fetching script catalog from: $Url"

    try {
        $response = Invoke-RestMethod -Uri $Url -UseBasicParsing -ErrorAction Stop
    }
    catch {
        $statusCode = $_.Exception.Response.StatusCode.value__
        if ($statusCode -eq 404) {
            Write-Error "Script catalog not found. The tag '$Tag' may not exist, or the repository is unreachable. URL: $Url"
        }
        else {
            Write-Error "Failed to download script catalog: $($_.Exception.Message)"
        }
        return $null
    }

    if ($null -eq $response.scripts -or $response.scripts.Count -eq 0) {
        Write-Error "Script catalog is empty or invalid. No scripts are available for tag '$Tag'."
        return $null
    }

    Write-Verbose "Found $($response.scripts.Count) script(s) in catalog."
    return $response.scripts
}

# ---------------------------------------------------------------------------
# Helper — Show-ScriptTable
# ---------------------------------------------------------------------------
function Show-ScriptTable {
    <#
    .SYNOPSIS
        Displays scripts in a formatted table.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Scripts
    )

    # Calculate column widths for clean alignment
    $nameWidth = ($Scripts | ForEach-Object { $_.name.Length } | Measure-Object -Maximum).Maximum
    $descWidth = ($Scripts | ForEach-Object { $_.description.Length } | Measure-Object -Maximum).Maximum
    if ($nameWidth -lt 4) { $nameWidth = 4 }
    if ($descWidth -lt 11) { $descWidth = 11 }
    # Cap description width to keep the table readable
    if ($descWidth -gt 70) { $descWidth = 70 }

    $headerName = 'Name'.PadRight($nameWidth)
    $headerDesc = 'Description'.PadRight($descWidth)
    $headerAdmin = 'Admin?'
    $headerTags = 'Tags'

    Write-Host ""
    Write-Host "  $headerName   $headerDesc   $headerAdmin   $headerTags" -ForegroundColor Yellow
    Write-Host "  $('-' * $nameWidth)   $('-' * $descWidth)   $('-' * 6)   $('-' * 20)" -ForegroundColor DarkGray

    foreach ($script in $Scripts) {
        $displayName  = $script.name.PadRight($nameWidth)
        $desc = $script.description
        if ($desc.Length -gt $descWidth) {
            $desc = $desc.Substring(0, $descWidth - 3) + '...'
        }
        $displayDesc  = $desc.PadRight($descWidth)
        $displayAdmin = if ($script.requiresAdmin) { 'Yes   ' } else { 'No    ' }
        $displayTags  = ($script.tags -join ', ')

        Write-Host "  $displayName   $displayDesc   $displayAdmin   $displayTags"
    }
    Write-Host ""
}

# ---------------------------------------------------------------------------
# Helper — Show-ScriptMenu
# ---------------------------------------------------------------------------
function Show-ScriptMenu {
    <#
    .SYNOPSIS
        Displays a numbered menu and returns the selected script object.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Scripts
    )

    Write-Host "  Available Scripts:" -ForegroundColor Yellow
    Write-Host ""

    for ($i = 0; $i -lt $Scripts.Count; $i++) {
        $num = ($i + 1).ToString().PadLeft(2)
        $name = $Scripts[$i].name
        $desc = $Scripts[$i].description
        $adminFlag = if ($Scripts[$i].requiresAdmin) { ' [Admin]' } else { '' }
        Write-Host "  [$num] " -ForegroundColor Cyan -NoNewline
        Write-Host "$name" -ForegroundColor White -NoNewline
        Write-Host "$adminFlag" -ForegroundColor Red -NoNewline
        Write-Host " - $desc" -ForegroundColor Gray
    }

    Write-Host ""

    while ($true) {
        Write-Host "  Select a script (1-$($Scripts.Count)) or Q to quit: " -ForegroundColor Yellow -NoNewline
        $choice = Read-Host

        if ($choice -eq 'Q' -or $choice -eq 'q') {
            Write-Host "  Exiting." -ForegroundColor DarkGray
            return $null
        }

        $selection = $null
        if ([int]::TryParse($choice, [ref]$selection)) {
            if ($selection -ge 1 -and $selection -le $Scripts.Count) {
                return $Scripts[$selection - 1]
            }
        }

        Write-Warning "Invalid selection. Please enter a number between 1 and $($Scripts.Count), or Q to quit."
    }
}

# ---------------------------------------------------------------------------
# Helper — Test-IsAdmin
# ---------------------------------------------------------------------------
function Test-IsAdmin {
    <#
    .SYNOPSIS
        Returns $true if the current session is running elevated (Administrator).
    #>
    [CmdletBinding()]
    param()

    $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]$identity
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# ---------------------------------------------------------------------------
# Helper — Confirm-Prerequisites
# ---------------------------------------------------------------------------
function Confirm-Prerequisites {
    <#
    .SYNOPSIS
        Checks admin and PS version requirements. Returns $true to proceed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$ScriptInfo
    )

    # Check minimum PowerShell version
    if ($ScriptInfo.minPSVersion) {
        $requiredVersion = [version]$ScriptInfo.minPSVersion
        $currentVersion  = [version]$PSVersionTable.PSVersion.ToString()
        if ($currentVersion -lt $requiredVersion) {
            Write-Warning "Script '$($ScriptInfo.name)' requires PowerShell $($ScriptInfo.minPSVersion) but you are running $($PSVersionTable.PSVersion)."
            Write-Warning "The script may not function correctly."
        }
    }

    # Check admin requirement
    if ($ScriptInfo.requiresAdmin) {
        $isAdmin = $false
        try {
            $isAdmin = Test-IsAdmin
        }
        catch {
            Write-Verbose "Could not determine admin status (non-Windows OS?). Skipping admin check."
            return $true
        }

        if (-not $isAdmin) {
            Write-Warning "Script '$($ScriptInfo.name)' requires Administrator privileges, but this session is NOT elevated."
            Write-Host ""
            Write-Host "  Continue anyway? (Y/N): " -ForegroundColor Yellow -NoNewline
            $confirm = Read-Host
            if ($confirm -ne 'Y' -and $confirm -ne 'y') {
                Write-Host "  Aborted." -ForegroundColor DarkGray
                return $false
            }
        }
    }

    return $true
}

# ---------------------------------------------------------------------------
# Helper — Invoke-RemoteScript
# ---------------------------------------------------------------------------
function Invoke-RemoteScript {
    <#
    .SYNOPSIS
        Downloads a script to $env:TEMP, executes it, and cleans up.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [object]$ScriptInfo,

        [Parameter(Mandatory)]
        [string]$BaseUrl,

        [Parameter(ValueFromRemainingArguments)]
        [object[]]$PassthroughArgs = @()
    )

    $scriptUrl = "$BaseUrl/$($ScriptInfo.path)"
    $safeName  = $ScriptInfo.name -replace '[^a-zA-Z0-9_-]', '_'
    $tempFile  = Join-Path -Path $env:TEMP -ChildPath "ps_runner_${safeName}.ps1"

    Write-Verbose "Downloading script from: $scriptUrl"
    Write-Verbose "Temp file: $tempFile"

    if (-not $PSCmdlet.ShouldProcess($ScriptInfo.name, "Download and execute script")) {
        return
    }

    try {
        try {
            $scriptContent = Invoke-RestMethod -Uri $scriptUrl -UseBasicParsing -ErrorAction Stop
        }
        catch {
            $statusCode = $_.Exception.Response.StatusCode.value__
            if ($statusCode -eq 404) {
                Write-Error "Script file not found at: $scriptUrl. The path in scripts.json may be incorrect for tag '$Tag'."
            }
            else {
                Write-Error "Failed to download script '$($ScriptInfo.name)': $($_.Exception.Message)"
            }
            return
        }

        Set-Content -Path $tempFile -Value $scriptContent -Encoding UTF8 -Force -ErrorAction Stop
        Write-Verbose "Script saved to: $tempFile"

        Write-Host ""
        Write-Host "  Running: $($ScriptInfo.name)" -ForegroundColor Green
        Write-Host "  $('-' * 50)" -ForegroundColor DarkGray
        Write-Host ""

        & $tempFile @PassthroughArgs
    }
    catch {
        Write-Error "Error executing script '$($ScriptInfo.name)': $($_.Exception.Message)"
    }
    finally {
        if (Test-Path -Path $tempFile -ErrorAction SilentlyContinue) {
            Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue
            Write-Verbose "Cleaned up temp file: $tempFile"
        }
    }
}

# ===========================================================================
# Main
# ===========================================================================

Write-Banner

# --- Fetch the script catalog ---
$catalog = Get-ScriptCatalog -Url $ScriptsJsonUrl
if ($null -eq $catalog) {
    exit 1
}

# --- List mode: print table and exit ---
if ($List) {
    Show-ScriptTable -Scripts $catalog
    exit 0
}

# --- Search mode: filter the catalog ---
if ($Search) {
    Write-Verbose "Searching for: '$Search'"
    $filtered = $catalog | Where-Object {
        $_.name -match [regex]::Escape($Search) -or
        $_.description -match [regex]::Escape($Search) -or
        ($_.tags -join ' ') -match [regex]::Escape($Search)
    }

    if ($null -eq $filtered -or @($filtered).Count -eq 0) {
        Write-Warning "No scripts matched the search term '$Search'."
        exit 0
    }

    $filtered = @($filtered)
    Write-Host "  Search results for: '$Search'" -ForegroundColor Magenta
    Write-Host ""

    $selectedScript = Show-ScriptMenu -Scripts $filtered
    if ($null -eq $selectedScript) {
        exit 0
    }
}
# --- Direct name mode: find the script by name ---
elseif ($ScriptName) {
    Write-Verbose "Looking up script by name: '$ScriptName'"
    $selectedScript = $catalog | Where-Object { $_.name -eq $ScriptName }

    if ($null -eq $selectedScript) {
        Write-Error "Script '$ScriptName' not found in the catalog. Use -List to see available scripts."
        exit 1
    }

    # If somehow multiple matches, take the first
    $selectedScript = @($selectedScript)[0]
    Write-Host "  Selected: $($selectedScript.name)" -ForegroundColor Green
}
# --- Interactive menu mode ---
else {
    $selectedScript = Show-ScriptMenu -Scripts @($catalog)
    if ($null -eq $selectedScript) {
        exit 0
    }
}

# --- Pre-flight checks ---
if (-not (Confirm-Prerequisites -ScriptInfo $selectedScript)) {
    exit 0
}

# --- Download and execute ---
Invoke-RemoteScript -ScriptInfo $selectedScript -BaseUrl $BaseRawUrl
