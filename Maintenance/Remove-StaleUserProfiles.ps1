#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Identifies and removes local user profiles that have been inactive beyond a
    configurable threshold.

.DESCRIPTION
    Remove-StaleUserProfiles queries all local user profiles via the
    Win32_UserProfile WMI class and compares each profile's LastUseTime against
    a staleness threshold (default: 90 days). System and built-in service
    accounts are automatically excluded.

    When run without -Confirm:$false, the script operates in dry-run mode by
    default thanks to SupportsShouldProcess. Use -WhatIf to explicitly preview
    deletions, or -Confirm:$false to proceed without prompting.

    Profile removal is performed through the CIM Delete() method, which removes
    both the profile folder and the corresponding registry entries -- the same
    operation as deluser or the System Properties "User Profiles" dialog.

    A CSV report is generated after every run containing the username, SID,
    last use time, profile path, and action taken for each evaluated profile.

.PARAMETER DaysInactive
    Number of days since last use after which a profile is considered stale.
    Defaults to 90.

.PARAMETER ReportFolder
    Directory where the CSV report is saved. Defaults to C:\Temp.
    Created automatically if it does not exist.

.PARAMETER NoReport
    Skip generating the CSV report file.

.PARAMETER ExcludeUsers
    An array of additional usernames (SAMAccountName, no domain prefix) to
    exclude from removal. Built-in and service accounts are always excluded
    regardless of this parameter.

.EXAMPLE
    .\Remove-StaleUserProfiles.ps1 -WhatIf

    Previews which profiles would be removed using the default 90-day threshold.
    No profiles are deleted and no CSV is written.

.EXAMPLE
    .\Remove-StaleUserProfiles.ps1 -DaysInactive 60 -Confirm:$false

    Removes all profiles inactive for 60+ days without prompting for each one.

.EXAMPLE
    .\Remove-StaleUserProfiles.ps1 -DaysInactive 120 -ExcludeUsers 'kiosk','labuser'

    Removes profiles inactive for 120+ days, skipping 'kiosk' and 'labuser'
    in addition to the built-in exclusions. Prompts for confirmation on each.

.EXAMPLE
    .\Remove-StaleUserProfiles.ps1 -DaysInactive 90 -Confirm:$false -NoReport

    Production / RMM usage: removes stale profiles, no confirmation prompt,
    no CSV report. Results are still written to the verbose stream.

.NOTES
    Author  : AI-Assisted (Claude)
    Version : 1.0.0
    Requires: PowerShell 5.1+, Administrator privileges
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [ValidateRange(1, 3650)]
    [int]$DaysInactive = 90,

    [string]$ReportFolder = 'C:\Temp',

    [switch]$NoReport,

    [string[]]$ExcludeUsers = @()
)

# ── Script metadata ──────────────────────────────────────────────────────────
$ScriptVersion = '1.0.0'
$StartTime     = Get-Date
$CutoffDate    = $StartTime.AddDays(-$DaysInactive)

Write-Verbose "Script version : $ScriptVersion"
Write-Verbose "Start time     : $StartTime"
Write-Verbose "Cutoff date    : $CutoffDate ($DaysInactive days ago)"

# ── Built-in exclusion list ──────────────────────────────────────────────────
# These profiles are never removed regardless of inactivity.
$BuiltInExclusions = @(
    'Administrator'
    'Default'
    'Default User'
    'Public'
    'NetworkService'
    'LocalService'
    'systemprofile'
    'SYSTEM'
    'DefaultAppPool'
)

# Well-known SID prefixes / exact SIDs for service and system accounts
$ExcludedSidPatterns = @(
    'S-1-5-18'   # SYSTEM
    'S-1-5-19'   # LOCAL SERVICE
    'S-1-5-20'   # NETWORK SERVICE
)

# Merge caller-supplied exclusions (case-insensitive comparison later)
$AllExcludedUsers = $BuiltInExclusions + $ExcludeUsers

# ── Helper: Resolve username from profile path ──────────────────────────────
function Get-UsernameFromProfile {
    <#
    .SYNOPSIS
        Extracts the leaf folder name from a profile's LocalPath as a best-effort
        username. Falls back to the SID when the path is empty or unusual.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$LocalPath,

        [Parameter(Mandatory)]
        [string]$Sid
    )

    if ([string]::IsNullOrWhiteSpace($LocalPath)) {
        return $Sid
    }

    return Split-Path -Path $LocalPath -Leaf
}

# ── Helper: Convert CIM datetime to .NET DateTime ───────────────────────────
function ConvertTo-DateTimeFromCim {
    <#
    .SYNOPSIS
        Converts a CIM/WMI datetime string (yyyyMMddHHmmss.ffffff+UUU) to a
        .NET DateTime, returning $null on failure.
    #>
    [CmdletBinding()]
    [OutputType([datetime])]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$CimDateTime
    )

    if ([string]::IsNullOrWhiteSpace($CimDateTime)) {
        return $null
    }

    try {
        return [Management.ManagementDateTimeConverter]::ToDateTime($CimDateTime)
    }
    catch {
        Write-Verbose "Could not parse CIM datetime '$CimDateTime': $_"
        return $null
    }
}

# ── Collect profiles ────────────────────────────────────────────────────────
Write-Verbose 'Querying Win32_UserProfile via CIM ...'

try {
    $AllProfiles = Get-CimInstance -ClassName Win32_UserProfile -ErrorAction Stop
}
catch {
    Write-Error "Failed to query Win32_UserProfile: $_"
    exit 1
}

Write-Verbose "Total profiles found: $($AllProfiles.Count)"

# ── Evaluate each profile ───────────────────────────────────────────────────
$Results  = [System.Collections.Generic.List[PSCustomObject]]::new()
$Removed  = 0
$Skipped  = 0
$Errors   = 0

foreach ($Profile in $AllProfiles) {

    $Sid         = $Profile.SID
    $LocalPath   = $Profile.LocalPath
    $Username    = Get-UsernameFromProfile -LocalPath $LocalPath -Sid $Sid
    $IsSpecial   = $Profile.Special -eq $true
    $IsLoaded    = $Profile.Loaded  -eq $true
    $LastUseTime = $null

    # Win32_UserProfile on PS 5.1 returns LastUseTime as a CIM datetime string
    # via WMI; on PS 7+ with CIM it may already be a DateTime object.
    if ($Profile.LastUseTime -is [datetime]) {
        $LastUseTime = $Profile.LastUseTime
    }
    elseif ($Profile.LastUseTime) {
        $LastUseTime = ConvertTo-DateTimeFromCim -CimDateTime ($Profile.LastUseTime.ToString())
    }

    $LastUseStr = if ($LastUseTime) { $LastUseTime.ToString('yyyy-MM-dd HH:mm:ss') } else { 'Unknown' }

    # ── Determine exclusion reason ───────────────────────────────────────
    $Action = $null
    $Reason = $null

    if ($IsSpecial) {
        $Action = 'Skipped'
        $Reason = 'Special/system profile'
    }
    elseif ($Sid -in $ExcludedSidPatterns) {
        $Action = 'Skipped'
        $Reason = 'Well-known service SID'
    }
    elseif ($Username -in $AllExcludedUsers) {
        $Action = 'Skipped'
        $Reason = 'Excluded username'
    }
    elseif ($IsLoaded) {
        $Action = 'Skipped'
        $Reason = 'Profile currently loaded (user logged in)'
    }
    elseif ($null -eq $LastUseTime) {
        $Action = 'Skipped'
        $Reason = 'LastUseTime not available'
    }
    elseif ($LastUseTime -ge $CutoffDate) {
        $Action = 'Skipped'
        $Reason = "Last used within $DaysInactive days ($LastUseStr)"
    }

    # ── If no exclusion reason, this profile is stale ────────────────────
    if ($null -eq $Action) {
        $DaysIdle = [math]::Floor(($StartTime - $LastUseTime).TotalDays)

        if ($PSCmdlet.ShouldProcess(
                "$Username (SID: $Sid, Last used: $LastUseStr, Idle: $DaysIdle days)",
                'Remove user profile')) {

            try {
                $Profile | Remove-CimInstance -ErrorAction Stop
                $Action = 'Removed'
                $Reason = "Inactive $DaysIdle days (last used $LastUseStr)"
                $Removed++
                Write-Verbose "REMOVED: $Username | SID: $Sid | Idle $DaysIdle days"
            }
            catch {
                $Action = 'Error'
                $Reason = $_.Exception.Message
                $Errors++
                Write-Warning "FAILED to remove profile '$Username' (SID: $Sid): $_"
            }
        }
        else {
            # -WhatIf was specified
            $Action = 'WouldRemove'
            $Reason = "Inactive $DaysIdle days (last used $LastUseStr)"
        }
    }
    else {
        $Skipped++
    }

    # ── Record result ────────────────────────────────────────────────────
    $Results.Add([PSCustomObject]@{
        Username    = $Username
        SID         = $Sid
        ProfilePath = $LocalPath
        LastUseTime = $LastUseStr
        Action      = $Action
        Reason      = $Reason
    })

    Write-Verbose "$($Action.ToUpper().PadRight(12)) $($Username.PadRight(25)) $LastUseStr  $Reason"
}

# ── CSV Report ───────────────────────────────────────────────────────────────
if (-not $NoReport -and -not $WhatIfPreference) {
    try {
        if (-not (Test-Path -Path $ReportFolder -PathType Container)) {
            New-Item -Path $ReportFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
            Write-Verbose "Created report folder: $ReportFolder"
        }

        $Timestamp  = $StartTime.ToString('yyyyMMdd_HHmmss')
        $Hostname   = $env:COMPUTERNAME
        $ReportPath = Join-Path -Path $ReportFolder -ChildPath "StaleProfiles_${Hostname}_${Timestamp}.csv"

        $Results | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        Write-Verbose "Report saved to: $ReportPath"
    }
    catch {
        Write-Warning "Failed to write CSV report: $_"
    }
}

# ── Summary ──────────────────────────────────────────────────────────────────
$Elapsed = (Get-Date) - $StartTime

$SummaryBlock = @"

====================================================================
  Remove-StaleUserProfiles  --  Summary
====================================================================
  Computer       : $env:COMPUTERNAME
  Run date       : $($StartTime.ToString('yyyy-MM-dd HH:mm:ss'))
  Threshold      : $DaysInactive days  (cutoff: $($CutoffDate.ToString('yyyy-MM-dd')))
  Profiles found : $($AllProfiles.Count)
  Skipped        : $Skipped
  Removed        : $Removed
  Would remove   : $(($Results | Where-Object { $_.Action -eq 'WouldRemove' }).Count)
  Errors         : $Errors
  Elapsed        : $($Elapsed.ToString('mm\:ss\.fff'))
"@

if (-not $NoReport -and -not $WhatIfPreference) {
    $SummaryBlock += "  Report         : $ReportPath`n"
}

$SummaryBlock += "===================================================================="

Write-Output $SummaryBlock
