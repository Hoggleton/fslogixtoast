# FSLogix Profile Toast Notification

A PowerShell logon script that warns users when their FSLogix Profile Container or Office Data File Container (ODFC) is approaching capacity.

FSLogix doesn't surface disk usage to end users. When a container hits its GPO-configured maximum, things break silently — Outlook stops caching, profile writes fail, settings reset. This script gives users a heads-up before that happens.

## What It Does

- Detects both **Profile Container** and **ODFC** independently
- Reads the GPO-configured `SizeInMBs` from registry (no hardcoded values)
- Finds FSLogix's **letterless VHDX volumes** via the Storage WMI namespace (`BusType 6`)
- Matches volumes to container types by comparing against GPO max sizes
- Fires a **Windows toast notification** when either container exceeds configurable thresholds
- Cooldown timer prevents nagging users on every reconnect

## Requirements

- Windows 10/11 or Server 2016+
- PowerShell 5.1+
- FSLogix installed and containers active
- Optional: [BurntToast](https://github.com/Windos/BurntToast) module for richer notifications

## Usage

```powershell
# Dry run with full console output (safe for production testing)
.\FSLogixToast.ps1 -DryRun -ShowDebug

# Test with actual toast notification
.\FSLogixToast.ps1 -Force -ShowDebug

# Production (silent, logs to %LOCALAPPDATA%\FSLogix\ProfileToast.log)
.\FSLogixToast.ps1
```

## Parameters

| Parameter | Default | Description |
|---|---|---|
| `-WarningThresholdPct` | 80 | Percentage to trigger a warning toast |
| `-CriticalThresholdPct` | 95 | Percentage to trigger a critical toast |
| `-CooldownHours` | 8 | Suppress repeat notifications within this window |
| `-Force` | — | Ignore cooldown timer |
| `-DryRun` | — | Run detection and show summary, skip toast |
| `-ShowDebug` | — | Stream all log messages to console |
| `-LogPath` | `%LOCALAPPDATA%\FSLogix\ProfileToast.log` | Log file location |

## Deployment

**GPO User Logon Script**
User Configuration → Policies → Windows Settings → Scripts → Logon

**Citrix WEM External Task**
Trigger at logon, filter by user group or delivery group.

**Scheduled Task via GPO Preferences**
Trigger on session connect to catch reconnections as well as fresh logons.

## How It Finds the Containers

FSLogix mounts VHDX containers as letterless volumes on virtual disks. The script walks the Storage WMI namespace:

```
MSFT_Volume (no drive letter) → MSFT_Partition → MSFT_Disk (BusType = 6)
```

When both Profile and ODFC are active, it matches volumes to container types by comparing total volume size against the GPO-configured `SizeInMBs` for each. If sizes are identical, it falls back to volume label matching.

## Toast Delivery

Tries three methods in order:
1. **BurntToast module** (if installed)
2. **Native Windows AppNotification API** (WinRT, no dependencies)
3. **System.Windows.Forms balloon tip** (legacy fallback)

## License

MIT
