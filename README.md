# Nakamas IT - PowerShell Scripts

![GitHub release](https://img.shields.io/github/v/release/nakamas-it/powershell-scripts?label=latest%20release)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)
![License](https://img.shields.io/github/license/nakamas-it/powershell-scripts)

A collection of PowerShell scripts for Windows administration — disk reports, inventory, maintenance, and more.
All scripts are standalone `.ps1` files. No modules to install. No cloning required.

---

## Quick Start

Paste into any PowerShell window:

```powershell
irm https://raw.githubusercontent.com/nakamas-it/powershell-scripts/main/Run.ps1 | iex
```

This launches an **interactive menu** listing all available scripts. Pick one and it runs automatically.

### Run a specific script by name

```powershell
& ([scriptblock]::Create((irm 'https://raw.githubusercontent.com/nakamas-it/powershell-scripts/main/Run.ps1'))) -ScriptName Get-DiskStorageReport
```

### List all available scripts

```powershell
& ([scriptblock]::Create((irm 'https://raw.githubusercontent.com/nakamas-it/powershell-scripts/main/Run.ps1'))) -List
```

### Search scripts by keyword

```powershell
& ([scriptblock]::Create((irm 'https://raw.githubusercontent.com/nakamas-it/powershell-scripts/main/Run.ps1'))) -Search disk
```

> **Older Windows 10 builds (TLS fix):** Prepend `[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;` before any command above.

---

## Available Scripts

| Name | Category | Description | Admin? |
|------|----------|-------------|--------|
| [Get-DiskStorageReport](Inventory/Get-DiskStorageReport.ps1) | Inventory | Scans all drives and generates offline HTML reports: top folders, top files, file type breakdown, and a summary dashboard | No |
| [Invoke-DiskCleanup](Maintenance/Invoke-DiskCleanup.ps1) | Maintenance | Cleans common Windows cache and temp folders (browser caches, Teams, WER, prefetch, shader caches, and more). Supports dry-run, all-user mode, optional targets, and generates an HTML summary report | Optional |
| [Remove-StaleUserProfiles](Maintenance/Remove-StaleUserProfiles.ps1) | Maintenance | Identifies and removes local user profiles inactive beyond a configurable threshold (default: 90 days). Supports `-WhatIf` dry-run, per-profile confirmation, custom exclusions, and CSV reporting | Yes |

---

## Repository Structure

```
powershell-scripts/
├── Run.ps1               ← central launcher (start here)
├── scripts.json          ← machine-readable manifest
├── README.md
├── SECURITY.md
├── .gitignore
├── Inventory/
│   └── Get-DiskStorageReport.ps1
└── Maintenance/
    ├── Invoke-DiskCleanup.ps1
    └── Remove-StaleUserProfiles.ps1
```

New categories (Security, Networking, etc.) will be added as folders alongside `Inventory/` and `Maintenance/`.

---

## Optional: Add a shortcut to your PowerShell profile

Add this to your `$PROFILE` (`notepad $PROFILE`) for quick access from any session:

```powershell
function Invoke-NakamasScript {
    param(
        [string]$ScriptName,
        [string]$Tag = 'main',
        [switch]$List,
        [string]$Search
    )
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $launcher = irm "https://raw.githubusercontent.com/nakamas-it/powershell-scripts/$Tag/Run.ps1"
    $params = @{}
    if ($ScriptName) { $params['ScriptName'] = $ScriptName }
    if ($List)       { $params['List']       = $true }
    if ($Search)     { $params['Search']     = $Search }
    & ([scriptblock]::Create($launcher)) @params
}
Set-Alias -Name nps -Value Invoke-NakamasScript
```

Then use: `nps`, `nps Get-DiskStorageReport`, `nps -List`, `nps -Search disk`

---

## Security

- All downloads use **HTTPS only**
- Scripts are fetched on-demand to `$env:TEMP` and deleted after execution
- No credentials, no network uploads — scripts only read local file metadata
- For enterprise/RMM deployment, sign scripts with `Set-AuthenticodeSignature`
- Verify integrity: `(Get-FileHash .\script.ps1 -Algorithm SHA256).Hash`

---

## Contributing

1. Add your script to the appropriate category folder (create one if needed)
2. Add an entry to `scripts.json` with name, path, description, requiresAdmin, minPSVersion, and tags
3. Open a pull request

---

## License

MIT
