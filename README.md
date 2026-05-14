# EventIDPuller

## Description

EventIDPuller is a PowerShell diagnostic script that collects high-value Windows Event IDs for hardware, software, user profile, AD, DNS, DHCP, DFSR, Group Policy, network, and authentication issues, then exports clean CSV reports with optional XLSX output.

## Overview

EventIDPuller is a PowerShell script for pulling useful Windows Event IDs tied to common troubleshooting areas such as hardware, software, user profile, Active Directory, DNS, DHCP, DFSR/SYSVOL, Group Policy, network, and authentication issues.

By default, the script exports matching event rows to a CSV report and prints summary tables in the console. Full event rows are not printed to the console unless requested.

## Requirements

- Windows PowerShell 5.1 or newer
- Run as Administrator for best results
- Administrator rights are usually required for Security log collection
- Remote scans require appropriate permissions and remote event log access
- Microsoft Excel must be installed if using `-ExportXlsx`

## Default Behavior

When run without category flags, EventIDPuller performs a workstation-focused scan for:

- Hardware
- Software
- User Profile
- Group Policy basics

Default output location:

```powershell
C:\Temp
```

Default output type:

```text
CSV
```

The CSV includes a summary section for each computer, followed by a clean event table.

## Basic Usage

Run a default scan for the last 24 hours:

```powershell
.\EventIDPuller.ps1 -HoursBack 24
```

Run a default scan for the last 72 hours:

```powershell
.\EventIDPuller.ps1 -HoursBack 72
```

Run all event categories:

```powershell
.\EventIDPuller.ps1 -HoursBack 24 -IncludeAll
```

Export CSV and XLSX:

```powershell
.\EventIDPuller.ps1 -HoursBack 24 -ExportXlsx
```

Export to a specific folder:

```powershell
.\EventIDPuller.ps1 -HoursBack 24 -CsvPath "C:\Temp"
```

Export to a specific XLSX file:

```powershell
.\EventIDPuller.ps1 -HoursBack 24 -ExportXlsx -XlsxPath "C:\Temp\EventReport.xlsx"
```

Run from Command Prompt:

```cmd
powershell.exe -ExecutionPolicy Bypass -File EventIDPuller.ps1 -HoursBack 24
```

## Common Examples

### Workstation scan

```powershell
.\EventIDPuller.ps1 -HoursBack 24
```

### Full local scan

```powershell
.\EventIDPuller.ps1 -HoursBack 72 -IncludeAll
```

### Domain Controller scan

```powershell
.\EventIDPuller.ps1 -ComputerName DC01,DC02 -HoursBack 72 -IncludeActiveDirectory -IncludeDns -IncludeDfsr -IncludeGroupPolicy -IncludeSecurity
```

### DHCP server scan

```powershell
.\EventIDPuller.ps1 -ComputerName DHCP01 -DaysBack 30 -IncludeDhcp
```

### DNS server scan

```powershell
.\EventIDPuller.ps1 -ComputerName DNS01 -HoursBack 72 -IncludeDns
```

### Show matching events in the console

```powershell
.\EventIDPuller.ps1 -HoursBack 24 -ShowConsoleEvents
```

### Show full event details in the console

```powershell
.\EventIDPuller.ps1 -HoursBack 24 -Detailed
```

## Output Files

### CSV

CSV output is enabled by default.

Default folder:

```powershell
C:\Temp
```

Example output file:

```text
C:\Temp\IssueEvents_20260514_153000.csv
```

### XLSX

XLSX output is optional and requires Excel to be installed.

```powershell
.\EventIDPuller.ps1 -HoursBack 24 -ExportXlsx
```

Example output file:

```text
C:\Temp\IssueEvents_20260514_153000.xlsx
```

## CSV Layout

Each computer section starts with summary lines:

```text
ReportGenerated: 2026-05-14 15:30:00
SearchStart: 2026-05-13 15:30:00
SearchEnd: 2026-05-14 15:30:00
Range: Last 24 hour(s)
ComputerName: PC01
```

Then the event table starts:

```text
TimeCreated
Severity
Level
IssueType
EventId
LogName
ProviderName
Description
Message
```

## Main Flags

| Flag | Description |
|---|---|
| `-HoursBack` | Searches back a specific number of hours. Overrides `-DaysBack`. |
| `-DaysBack` | Searches back a specific number of days. Default is 7. |
| `-ComputerName` | Scans one or more computers. Defaults to the local computer. |
| `-CsvPath` | Folder or CSV file path for CSV output. |
| `-ExportXlsx` | Creates an XLSX copy of the report. Requires Excel. |
| `-XlsxPath` | Folder or XLSX file path for XLSX output. |
| `-NoCsv` | Disables CSV output. |
| `-ShowConsoleEvents` | Prints matching event rows in the console. |
| `-Detailed` | Prints full event details in the console. |
| `-MaxEventsPerQuery` | Limits results per event query group. |

## Category Flags

| Flag | Description |
|---|---|
| `-IncludeHardware` | Hardware, disk, storage, power, PnP, graphics, and driver events. |
| `-IncludeSoftware` | App crashes, service failures, MSI failures, and Windows Update failures. |
| `-IncludeUserProfile` | User Profile Service events. |
| `-IncludeGroupPolicy` | Group Policy processing and operational events. |
| `-IncludeActiveDirectory` | AD DS, replication, Netlogon, and time sync events. |
| `-IncludeDns` | DNS Server and DNS Client events. |
| `-IncludeDhcp` | DHCP Server Admin and Operational events. |
| `-IncludeDfsr` | DFSR, SYSVOL, and legacy FRS events. |
| `-IncludeSecurity` | Failed logons, Kerberos/NTLM failures, lockouts, account changes, and security monitoring events. |
| `-IncludeNetwork` | TCP/IP, NLA, and network-related events. |
| `-IncludeAll` | Enables all categories. |
| `-IncludeNoise` | Adds general/noisy events that are excluded by default. |

## Notes

- `-IncludeSecurity` can produce more data and usually requires Administrator rights.
- `-IncludeNoise` adds common successful/general events and should be used only when needed.
- Some logs only exist on specific server roles. Missing logs are skipped quietly unless `-Verbose` is used.
- XLSX export uses Excel COM automation, so it will not work on systems without Microsoft Excel installed.

## Troubleshooting

Run with verbose logging:

```powershell
.\EventIDPuller.ps1 -HoursBack 24 -Verbose
```

Use a custom output path if the default folder is unavailable:

```powershell
.\EventIDPuller.ps1 -HoursBack 24 -CsvPath "C:\Temp"
```

Check full help:

```powershell
Get-Help .\EventIDPuller.ps1 -Full
```
