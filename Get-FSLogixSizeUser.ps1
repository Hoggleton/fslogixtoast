<#
.SYNOPSIS
    FSLogix Profile & ODFC Size Toast Notification - User Logon Script
.DESCRIPTION
    Checks the user's FSLogix Profile Container AND Office Data File Container
    (ODFC) against their GPO-configured maximums and displays a Windows toast
    notification if either is near or at capacity.

    Deploy via GPO User Logon Script, Scheduled Task at logon, or Citrix WEM.
.PARAMETER WarningThresholdPct
    Percentage of max size to trigger a warning toast (default 80)
.PARAMETER CriticalThresholdPct
    Percentage of max size to trigger a critical toast (default 95)
.PARAMETER CooldownHours
    Suppress repeat notifications within this window (default 8)
.PARAMETER Force
    Ignore cooldown timer
.PARAMETER DryRun
    Run all detection and show summary, but don't fire toasts
.PARAMETER ShowDebug
    Stream all log messages to console
.EXAMPLE
    # Full debug dry run
    .\FSLogixToast.ps1 -DryRun -ShowDebug

    # Test with actual toast
    .\FSLogixToast.ps1 -Force -ShowDebug

    # Production (silent, logs to file)
    .\FSLogixToast.ps1
#>

#Requires -Version 5.1

param(
    [int]$WarningThresholdPct = 80,
    [int]$CriticalThresholdPct = 95,
    [string]$LogPath = "$env:LOCALAPPDATA\FSLogix\ProfileToast.log",
    [int]$CooldownHours = 8,
    [switch]$Force,
    [switch]$DryRun,
    [switch]$ShowDebug
)

if ($ShowDebug) { $VerbosePreference = "Continue" }

# ============================================================================
# LOGGING
# ============================================================================
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$ts] [$Level] $Message"
    try {
        $dir = Split-Path $LogPath -Parent
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        Add-Content -Path $LogPath -Value $entry -ErrorAction SilentlyContinue
    } catch { }

    switch ($Level) {
        "ERROR" { Write-Host "  [!] $Message" -ForegroundColor Red }
        "WARN"  { Write-Host "  [~] $Message" -ForegroundColor Yellow }
        default { Write-Verbose "  [+] $Message" }
    }
}

# ============================================================================
# GET FSLOGIX MAX SIZE FROM REGISTRY (SET BY GPO) - PER CONTAINER TYPE
# ============================================================================
function Get-FSLogixMaxSizeMB {
    param(
        [ValidateSet("Profile","ODFC")]
        [string]$ContainerType
    )

    if ($ContainerType -eq "Profile") {
        $regPaths = @(
            "HKLM:\SOFTWARE\Policies\FSLogix\Profiles",    # GPO-managed
            "HKLM:\SOFTWARE\FSLogix\Profiles"               # Direct config
        )
        $defaultMB = 30720  # 30 GB default
    }
    else {
        $regPaths = @(
            "HKLM:\SOFTWARE\Policies\FSLogix\ODFC",        # GPO-managed
            "HKLM:\SOFTWARE\FSLogix\ODFC"                   # Direct config
        )
        $defaultMB = 30720  # 30 GB default
    }

    foreach ($path in $regPaths) {
        try {
            $val = Get-ItemPropertyValue -Path $path -Name "SizeInMBs" -ErrorAction Stop
            if ($val -and $val -gt 0) {
                Write-Log "$ContainerType container: Found SizeInMBs=$val MB at $path"
                return [int]$val
            }
        } catch { }
    }

    Write-Log "$ContainerType container: No GPO SizeInMBs found, using default $defaultMB MB" -Level "WARN"
    return $defaultMB
}

# ============================================================================
# CHECK IF A CONTAINER TYPE IS ENABLED
# ============================================================================
function Test-ContainerEnabled {
    param(
        [ValidateSet("Profile","ODFC")]
        [string]$ContainerType
    )

    if ($ContainerType -eq "Profile") {
        $regPaths = @(
            "HKLM:\SOFTWARE\Policies\FSLogix\Profiles",
            "HKLM:\SOFTWARE\FSLogix\Profiles"
        )
    }
    else {
        $regPaths = @(
            "HKLM:\SOFTWARE\Policies\FSLogix\ODFC",
            "HKLM:\SOFTWARE\FSLogix\ODFC"
        )
    }

    foreach ($path in $regPaths) {
        try {
            $enabled = Get-ItemPropertyValue -Path $path -Name "Enabled" -ErrorAction Stop
            if ($enabled -eq 1) {
                Write-Log "$ContainerType container is ENABLED at $path"
                return $true
            }
        } catch { }
    }

    Write-Log "$ContainerType container: Enabled key not found or not set to 1" -Level "WARN"
    return $false
}

# ============================================================================
# GET VHD LOCATIONS FROM REGISTRY
# ============================================================================
function Get-ContainerVHDLocations {
    param(
        [ValidateSet("Profile","ODFC")]
        [string]$ContainerType
    )

    if ($ContainerType -eq "Profile") {
        $regPaths = @(
            "HKLM:\SOFTWARE\Policies\FSLogix\Profiles",
            "HKLM:\SOFTWARE\FSLogix\Profiles"
        )
    }
    else {
        $regPaths = @(
            "HKLM:\SOFTWARE\Policies\FSLogix\ODFC",
            "HKLM:\SOFTWARE\FSLogix\ODFC"
        )
    }

    foreach ($path in $regPaths) {
        try {
            $loc = Get-ItemPropertyValue -Path $path -Name "VHDLocations" -ErrorAction Stop
            if ($loc) {
                Write-Log "$ContainerType container VHDLocations: $loc"
                return $loc
            }
        } catch { }
    }
    return $null
}

# ============================================================================
# FIND MOUNTED FSLOGIX VOLUMES (LETTERLESS, VIRTUAL DISK)
# ============================================================================
function Get-FSLogixMountedVolumes {
    # Returns an array of objects with volume info for all letterless
    # virtual-disk-backed volumes (BusType 6) in the FSLogix size range
    $results = @()

    try {
        $volumes = Get-CimInstance -ClassName MSFT_Volume `
                   -Namespace root\Microsoft\Windows\Storage -ErrorAction Stop

        foreach ($vol in $volumes) {
            # FSLogix volumes have no drive letter
            if ($vol.DriveLetter) { continue }
            # Skip tiny (EFI/recovery) and huge (OS/data disk) volumes
            if ($vol.Size -lt 500MB -or $vol.Size -gt 500GB) { continue }
            if ($null -eq $vol.SizeRemaining) { continue }

            # Walk back to the disk and check if it's virtual
            $partition = $vol | Get-CimAssociatedInstance `
                         -ResultClassName MSFT_Partition -ErrorAction SilentlyContinue
            if (-not $partition) { continue }

            $disk = $partition | Get-CimAssociatedInstance `
                    -ResultClassName MSFT_Disk -ErrorAction SilentlyContinue
            if (-not $disk) { continue }

            # BusType 6 = Virtual (VHD/VHDX mounted)
            if ($disk.BusType -ne 6) { continue }

            $usedMB = [math]::Round(($vol.Size - $vol.SizeRemaining) / 1MB, 0)
            $totalMB = [math]::Round($vol.Size / 1MB, 0)

            Write-Log "Found virtual disk volume: Path=$($vol.Path) DiskNumber=$($disk.Number) Size=$([math]::Round($vol.Size/1GB,2))GB Used=$([math]::Round(($vol.Size-$vol.SizeRemaining)/1GB,2))GB Label=$($vol.FileSystemLabel)"

            $results += [PSCustomObject]@{
                VolumePath      = $vol.Path
                Label           = $vol.FileSystemLabel
                DiskNumber      = $disk.Number
                DiskLocation    = $disk.Location
                DiskFriendly    = $disk.FriendlyName
                UsedMB          = $usedMB
                TotalMB         = $totalMB
                SizeRemaining   = $vol.SizeRemaining
            }
        }
    } catch {
        Write-Log "MSFT_Volume enumeration failed: $_" -Level "ERROR"
    }

    # Fallback: Win32_Volume for older OS
    if ($results.Count -eq 0) {
        Write-Log "Trying Win32_Volume fallback..."
        try {
            $volumes = Get-CimInstance -ClassName Win32_Volume -ErrorAction Stop |
                       Where-Object {
                           -not $_.DriveLetter -and
                           $_.Capacity -gt 500MB -and
                           $_.Capacity -lt 500GB -and
                           $_.FileSystem -eq 'NTFS'
                       }

            foreach ($vol in $volumes) {
                $usedMB = [math]::Round(($vol.Capacity - $vol.FreeSpace) / 1MB, 0)
                $totalMB = [math]::Round($vol.Capacity / 1MB, 0)

                Write-Log "Win32_Volume candidate: DeviceID=$($vol.DeviceID) Label=$($vol.Label) Size=$([math]::Round($vol.Capacity/1GB,2))GB Used=$([math]::Round(($vol.Capacity-$vol.FreeSpace)/1GB,2))GB"

                $results += [PSCustomObject]@{
                    VolumePath      = $vol.DeviceID
                    Label           = $vol.Label
                    DiskNumber      = $null
                    DiskLocation    = $null
                    DiskFriendly    = $null
                    UsedMB          = $usedMB
                    TotalMB         = $totalMB
                    SizeRemaining   = $vol.FreeSpace
                }
            }
        } catch {
            Write-Log "Win32_Volume fallback failed: $_" -Level "ERROR"
        }
    }

    Write-Log "Total candidate FSLogix volumes found: $($results.Count)"
    return $results
}

# ============================================================================
# IDENTIFY WHICH VOLUME IS PROFILE vs ODFC
# ============================================================================
function Resolve-ContainerVolumes {
    param(
        [array]$Volumes,
        [int]$ProfileMaxMB,
        [int]$ODFCMaxMB,
        [bool]$ProfileEnabled,
        [bool]$ODFCEnabled
    )

    # Strategy: Match volumes to containers based on their total size
    # matching the GPO-configured SizeInMBs. FSLogix creates the VHDX
    # at the size specified in GPO, so total volume size ≈ SizeInMBs.

    $result = @{
        Profile = $null
        ODFC    = $null
    }

    # If only one container type is enabled and we have one volume, it's easy
    $enabledCount = ([int]$ProfileEnabled + [int]$ODFCEnabled)
    if ($Volumes.Count -eq 1 -and $enabledCount -eq 1) {
        $type = if ($ProfileEnabled) { "Profile" } else { "ODFC" }
        $result[$type] = $Volumes[0]
        Write-Log "Single volume, single container type: assigned to $type"
        return $result
    }

    # If sizes differ, match by closest to GPO max
    # Allow 5% tolerance since filesystem overhead reduces usable space
    foreach ($vol in $Volumes) {
        $volTotalMB = $vol.TotalMB
        $profileDiff = [math]::Abs($volTotalMB - $ProfileMaxMB) / $ProfileMaxMB
        $odfcDiff = [math]::Abs($volTotalMB - $ODFCMaxMB) / $ODFCMaxMB

        Write-Log "Volume $($vol.VolumePath): TotalMB=$volTotalMB ProfileDiff=$([math]::Round($profileDiff*100,1))% ODFCDiff=$([math]::Round($odfcDiff*100,1))%"

        if ($ProfileEnabled -and $ODFCEnabled) {
            if ($ProfileMaxMB -ne $ODFCMaxMB) {
                # Different max sizes - match by closest
                if ($profileDiff -lt $odfcDiff -and -not $result.Profile) {
                    $result.Profile = $vol
                    Write-Log "  -> Assigned to Profile (closer size match)"
                }
                elseif (-not $result.ODFC) {
                    $result.ODFC = $vol
                    Write-Log "  -> Assigned to ODFC (closer size match)"
                }
            }
            else {
                # Same max sizes - try volume label, disk location, or order
                if ($vol.Label -match "Profile" -and -not $result.Profile) {
                    $result.Profile = $vol
                    Write-Log "  -> Assigned to Profile (label match)"
                }
                elseif ($vol.Label -match "O365|Office|ODFC" -and -not $result.ODFC) {
                    $result.ODFC = $vol
                    Write-Log "  -> Assigned to ODFC (label match)"
                }
                elseif (-not $result.Profile) {
                    $result.Profile = $vol
                    Write-Log "  -> Assigned to Profile (first unmatched)"
                }
                elseif (-not $result.ODFC) {
                    $result.ODFC = $vol
                    Write-Log "  -> Assigned to ODFC (second unmatched)"
                }
            }
        }
        elseif ($ProfileEnabled -and -not $result.Profile) {
            $result.Profile = $vol
            Write-Log "  -> Assigned to Profile (only Profile enabled)"
        }
        elseif ($ODFCEnabled -and -not $result.ODFC) {
            $result.ODFC = $vol
            Write-Log "  -> Assigned to ODFC (only ODFC enabled)"
        }
    }

    return $result
}

# ============================================================================
# COOLDOWN CHECK
# ============================================================================
function Test-Cooldown {
    $cooldownFile = "$env:LOCALAPPDATA\FSLogix\ProfileToast.cooldown"
    if ($Force) { return $false }

    if (Test-Path $cooldownFile) {
        try {
            $lastRun = Get-Content $cooldownFile -Raw | ForEach-Object { [datetime]::Parse($_) }
            if ((Get-Date) -lt $lastRun.AddHours($CooldownHours)) {
                Write-Log "Within cooldown window (last: $lastRun, window: ${CooldownHours}h)"
                return $true
            }
        } catch { }
    }
    return $false
}

function Set-Cooldown {
    $cooldownFile = "$env:LOCALAPPDATA\FSLogix\ProfileToast.cooldown"
    try {
        $dir = Split-Path $cooldownFile -Parent
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        (Get-Date).ToString("o") | Set-Content -Path $cooldownFile -Force
    } catch {
        Write-Log "Failed to set cooldown: $_" -Level "WARN"
    }
}

# ============================================================================
# TOAST NOTIFICATION
# ============================================================================
function Show-ToastNotification {
    param(
        [string]$Title,
        [string]$Message,
        [string]$Level
    )

    # BurntToast
    if (Get-Module -ListAvailable -Name BurntToast -ErrorAction SilentlyContinue) {
        try {
            Import-Module BurntToast -ErrorAction Stop
            $btParams = @{ Text = $Title, $Message; AppLogo = $null }
            if ($Level -eq "Critical") { $btParams.Sound = "Alarm" }
            New-BurntToastNotification @btParams
            Write-Log "Toast sent via BurntToast ($Level)"
            return $true
        } catch {
            Write-Log "BurntToast failed: $_" -Level "WARN"
        }
    }

    # Native Windows AppNotification
    try {
        [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
        [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom, ContentType = WindowsRuntime] | Out-Null

        $escapedTitle = [System.Security.SecurityElement]::Escape($Title)
        $escapedMessage = [System.Security.SecurityElement]::Escape($Message)

        $template = @"
<toast duration="long">
    <visual>
        <binding template="ToastGeneric">
            <text>$escapedTitle</text>
            <text>$escapedMessage</text>
        </binding>
    </visual>
    <audio src="ms-winsoundevent:Notification.Default" />
</toast>
"@

        $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
        $xml.LoadXml($template)
        $appId = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'
        $toast = [Windows.UI.Notifications.ToastNotification]::new($xml)
        [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($appId).Show($toast)
        Write-Log "Toast sent via native AppNotification ($Level)"
        return $true
    } catch {
        Write-Log "Native toast failed: $_" -Level "WARN"
    }

    # Balloon tip fallback
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $icon = if ($Level -eq "Critical") { [System.Windows.Forms.ToolTipIcon]::Error } else { [System.Windows.Forms.ToolTipIcon]::Warning }
        $notify = New-Object System.Windows.Forms.NotifyIcon
        $notify.Icon = [System.Drawing.SystemIcons]::Warning
        $notify.Visible = $true
        $notify.BalloonTipTitle = $Title
        $notify.BalloonTipText = $Message
        $notify.BalloonTipIcon = $icon
        $notify.ShowBalloonTip(15000)
        Start-Sleep -Seconds 16
        $notify.Dispose()
        Write-Log "Toast sent via balloon tip fallback ($Level)"
        return $true
    } catch {
        Write-Log "All notification methods failed: $_" -Level "ERROR"
        return $false
    }
}

# ============================================================================
# EVALUATE A SINGLE CONTAINER
# ============================================================================
function Get-ContainerStatus {
    param(
        [string]$ContainerType,
        [PSCustomObject]$Volume,
        [int]$MaxSizeMB
    )

    if (-not $Volume) { return $null }

    $usedMB = $Volume.UsedMB
    $maxGB = [math]::Round($MaxSizeMB / 1024, 1)
    $usedGB = [math]::Round($usedMB / 1024, 1)
    $pct = [math]::Round(($usedMB / $MaxSizeMB) * 100, 1)
    $remainGB = [math]::Round(($MaxSizeMB - $usedMB) / 1024, 1)

    $status = "Healthy"
    if ($pct -ge $CriticalThresholdPct) { $status = "Critical" }
    elseif ($pct -ge $WarningThresholdPct) { $status = "Warning" }

    return [PSCustomObject]@{
        Type        = $ContainerType
        UsedMB      = $usedMB
        MaxMB       = $MaxSizeMB
        UsedGB      = $usedGB
        MaxGB       = $maxGB
        Pct         = $pct
        RemainingGB = $remainGB
        Status      = $status
    }
}

# ============================================================================
# DISPLAY SUMMARY BOX
# ============================================================================
function Show-Summary {
    param([array]$Containers)

    Write-Host ""
    Write-Host "  ┌──────────────────────────────────────────────────────────┐" -ForegroundColor DarkGray
    Write-Host "  │  FSLogix Container Summary                              │" -ForegroundColor DarkGray
    Write-Host "  ├──────────────────────────────────────────────────────────┤" -ForegroundColor DarkGray
    Write-Host "  │  User:     $env:USERNAME" -ForegroundColor Gray
    Write-Host "  │  Computer: $env:COMPUTERNAME" -ForegroundColor Gray
    Write-Host "  ├──────────────────────────────────────────────────────────┤" -ForegroundColor DarkGray

    foreach ($c in $Containers) {
        $pctColor = switch ($c.Status) {
            "Critical" { "Red" }
            "Warning"  { "Yellow" }
            default    { "Green" }
        }

        Write-Host "  │" -ForegroundColor DarkGray
        Write-Host "  │  $($c.Type) Container" -ForegroundColor White
        Write-Host "  │    Max Size:      $($c.MaxGB) GB (from GPO)" -ForegroundColor Gray
        Write-Host "  │    Used:          $($c.UsedGB) GB" -ForegroundColor Gray
        Write-Host "  │    Remaining:     $($c.RemainingGB) GB" -ForegroundColor Gray
        Write-Host "  │    Utilisation:   $($c.Pct)%" -ForegroundColor $pctColor
        Write-Host "  │    Status:        $($c.Status)" -ForegroundColor $pctColor
    }

    Write-Host "  │" -ForegroundColor DarkGray
    Write-Host "  └──────────────────────────────────────────────────────────┘" -ForegroundColor DarkGray
    Write-Host ""
}

# ============================================================================
# MAIN
# ============================================================================
Write-Log "========== FSLogix Profile Toast Check Started =========="
Write-Log "User: $env:USERNAME | Computer: $env:COMPUTERNAME"
Write-Log "Thresholds: Warning=${WarningThresholdPct}% Critical=${CriticalThresholdPct}%"
if ($DryRun) { Write-Host "`n  === DRY RUN MODE - no toast will be shown ===" -ForegroundColor Cyan }

# Check cooldown
if (Test-Cooldown) {
    Write-Log "Cooldown active, exiting"
    if (-not $DryRun) { exit 0 }
    Write-Host "  [i] Cooldown active but continuing (dry run)" -ForegroundColor Cyan
}

# Detect which container types are enabled
$profileEnabled = Test-ContainerEnabled -ContainerType "Profile"
$odfcEnabled = Test-ContainerEnabled -ContainerType "ODFC"

if (-not $profileEnabled -and -not $odfcEnabled) {
    Write-Log "Neither Profile nor ODFC container is enabled - nothing to check" -Level "WARN"
    exit 0
}

# Get GPO-configured max sizes
$profileMaxMB = if ($profileEnabled) { Get-FSLogixMaxSizeMB -ContainerType "Profile" } else { 0 }
$odfcMaxMB = if ($odfcEnabled) { Get-FSLogixMaxSizeMB -ContainerType "ODFC" } else { 0 }

# Find all mounted FSLogix volumes
$volumes = Get-FSLogixMountedVolumes

if ($volumes.Count -eq 0) {
    Write-Log "No FSLogix mounted volumes found - user may not have an active container" -Level "WARN"
    exit 0
}

# Match volumes to container types
$containerVols = Resolve-ContainerVolumes -Volumes $volumes `
    -ProfileMaxMB $profileMaxMB -ODFCMaxMB $odfcMaxMB `
    -ProfileEnabled $profileEnabled -ODFCEnabled $odfcEnabled

# Evaluate each container
$statuses = @()

if ($containerVols.Profile) {
    $profileStatus = Get-ContainerStatus -ContainerType "Profile" -Volume $containerVols.Profile -MaxSizeMB $profileMaxMB
    Write-Log "Profile: $($profileStatus.UsedGB)GB / $($profileStatus.MaxGB)GB = $($profileStatus.Pct)% [$($profileStatus.Status)]"
    $statuses += $profileStatus
}
elseif ($profileEnabled) {
    Write-Log "Profile container enabled but no matching volume found" -Level "WARN"
}

if ($containerVols.ODFC) {
    $odfcStatus = Get-ContainerStatus -ContainerType "ODFC" -Volume $containerVols.ODFC -MaxSizeMB $odfcMaxMB
    Write-Log "ODFC: $($odfcStatus.UsedGB)GB / $($odfcStatus.MaxGB)GB = $($odfcStatus.Pct)% [$($odfcStatus.Status)]"
    $statuses += $odfcStatus
}
elseif ($odfcEnabled) {
    Write-Log "ODFC container enabled but no matching volume found" -Level "WARN"
}

# Show summary in debug/dry run mode
if (($DryRun -or $ShowDebug) -and $statuses.Count -gt 0) {
    Show-Summary -Containers $statuses
}

# Build toast message for any containers that need attention
$toastParts = @()
$worstLevel = "Healthy"

foreach ($s in $statuses) {
    if ($s.Status -eq "Critical") {
        $toastParts += "$($s.Type): $($s.Pct)% full ($($s.UsedGB)GB of $($s.MaxGB)GB) - only $($s.RemainingGB)GB remaining!"
        $worstLevel = "Critical"
    }
    elseif ($s.Status -eq "Warning") {
        $toastParts += "$($s.Type): $($s.Pct)% full ($($s.UsedGB)GB of $($s.MaxGB)GB) - $($s.RemainingGB)GB remaining."
        if ($worstLevel -ne "Critical") { $worstLevel = "Warning" }
    }
}

if ($toastParts.Count -gt 0) {
    $title = if ($worstLevel -eq "Critical") { "Profile Storage Critical" } else { "Profile Storage Warning" }
    $message = ($toastParts -join "`n") + "`nPlease clear unnecessary files or contact IT."

    Write-Log "Sending $worstLevel toast for $($toastParts.Count) container(s)" -Level "WARN"

    if (-not $DryRun) {
        Show-ToastNotification -Title $title -Message $message -Level $worstLevel
        Set-Cooldown
    }
    else {
        Write-Host "  [DRY RUN] Would show $worstLevel toast:" -ForegroundColor $(if ($worstLevel -eq "Critical") { "Red" } else { "Yellow" })
        foreach ($part in $toastParts) { Write-Host "    $part" -ForegroundColor Gray }
    }
}
else {
    Write-Log "All containers healthy - no notification needed"
    if ($DryRun) { Write-Host "  [DRY RUN] No toast needed - all containers healthy" -ForegroundColor Green }
}

Write-Log "========== FSLogix Profile Toast Check Complete =========="
exit 0