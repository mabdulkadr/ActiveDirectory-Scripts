<div align="center">

# 🏢 Active Directory Scripts

**A production-oriented toolkit of 9 PowerShell tools for day-to-day Active Directory administration.**

For IT operations and helpdesk teams that manage users, computers, groups, and Domain Controllers every day.

[![Tools](https://img.shields.io/badge/Tools-9%20Tools-10B981?style=for-the-badge)](#-scripts-included)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.0-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-core-features) • [Scripts](#-scripts-included) • [Requirements](#-requirements) • [License](#-license)

</div>

---

# 📖 Overview

**Active Directory Scripts** is a toolkit of 9 standalone PowerShell tools covering the most common Active Directory tasks: health checks, reporting, bulk user management, computer lifecycle, and group synchronization. Each tool ships in its own folder with a dedicated README, works independently, and depends only on the RSAT **ActiveDirectory** module on Windows PowerShell 5.1+.

---

# ✨ Core Features

### 🔹 Get AD Health (`Get-ADHealth\`)
* Domain Controller health check covering DNS, ICMP, uptime, time sync, disk, services, and `DCDIAG`, delivered as an HTML report.

### 🔹 Active Directory Reports (`ActiveDirectoryReports\`)
* Six menu-driven HTML inventory reports: computers, Domain Controllers, workstations, servers, account status, and OS grouping.

### 🔹 AD Groups Report (`AD-GroupsReport\`)
* Full group inventory CSV with metadata, member-type counts, nesting, and inactivity flags.

### 🔹 Get AD Users Details (`Get-ADUsers-Details\`)
* Ou-scoped user inventory with inactivity calculation and desktop CSV export.

### 🔹 Disable Users in OU (`DisableUsersInOU\`)
* Bulk-disables every enabled user account in an OU, with per-user logging and optional CSV report.

### 🔹 Disable-InactiveComputers (`Disable-InactiveComputers\`)
* Identifies stale computer objects by last logon and optionally disables or deletes them with an HTML audit trail.

### 🔹 Add-UsersFromCSVToADGroup (`Add-UsersFromCSVToADGroup\`)
* Bulk-adds users to a group from a CSV file, skipping existing members.

### 🔹 Shadow Group (`Shadow-Group\`)
* Keeps a group in exact sync with the users of one or more OUs, with dry-run mode and a per-run change cap.

### 🔹 Join Unjoin Computer Tool (`JoinUnjoinComputerTool-GUI\`)
* WPF GUI for helpdesk teams to join, disjoin, and clean up computer objects across AD and Microsoft Entra ID, with light/dark theme.

---

# 📂 Project Structure

```text
ActiveDirectory-Scripts
│
├── ActiveDirectoryReports\
│   └── ActiveDirectoryReports.ps1
├── AD-GroupsReport\
│   └── AD-GroupsReport.ps1
├── Add-UsersFromCSVToADGroup\
│   └── Add-UsersFromCSVToADGroup.ps1
├── Disable-InactiveComputers\
│   └── DisableInactiveComputers.ps1
├── DisableUsersInOU\
│   └── DisableUsersInOU.ps1
├── Get-ADHealth\
│   └── Get-ADHealth.ps1
├── Get-ADUsers-Details\
│   └── Get-ADUsers-Details.ps1
├── JoinUnjoinComputerTool-GUI\
│   ├── JoinUnjoinComputerTool-v1.ps1
│   └── JoinUnjoinComputerTool-v2.ps1
├── Shadow-Group\
│   └── ShadowGroup.ps1
└── README.md
```

---

# 🚀 Scripts Included

## 🔎 Get AD Health (CLI)

**File**
```powershell
Get-ADHealth\Get-ADHealth.ps1
```

Comprehensive Domain Controller health check with HTML report and optional SMTP delivery.

| Code | Status |
| ---- | ------ |
| 0    | Success |
| 1    | Failure |

## 🔎 Active Directory Reports (CLI)

**File**
```powershell
ActiveDirectoryReports\ActiveDirectoryReports.ps1
```

Menu-driven HTML inventory for AD computers, Domain Controllers, workstations, servers, account status, and OS grouping.

| Code | Status |
| ---- | ------ |
| 0    | Success |
| 1    | Failure |

## 🔎 AD Groups Report (CLI)

**File**
```powershell
AD-GroupsReport\AD-GroupsReport.ps1
```

Full domain group inventory with membership, nesting, and inactivity analysis.

| Code | Status |
| ---- | ------ |
| 0    | Success |
| 1    | Failure |

## 🔎 Get AD Users Details (CLI)

**File**
```powershell
Get-ADUsers-Details\Get-ADUsers-Details.ps1
```

OU-scoped user inventory with account attributes and inactivity metrics.

| Code | Status |
| ---- | ------ |
| 0    | Success |
| 1    | Failure |

## 🔎 Disable Users in OU (CLI)

**File**
```powershell
DisableUsersInOU\DisableUsersInOU.ps1
```

Bulk disables every enabled user account in a target OU.

| Code | Status |
| ---- | ------ |
| 0    | Success |
| 1    | Failure |

## 🔎 Disable-InactiveComputers (CLI)

**File**
```powershell
Disable-InactiveComputers\DisableInactiveComputers.ps1
```

Finds stale computer accounts and optionally disables or deletes them.

| Code | Status |
| ---- | ------ |
| 0    | Success |
| 1    | Failure |

## 🔎 Add-UsersFromCSVToADGroup (CLI)

**File**
```powershell
Add-UsersFromCSVToADGroup\Add-UsersFromCSVToADGroup.ps1
```

Bulk-adds users to an AD group from a CSV file, skipping existing members.

| Code | Status |
| ---- | ------ |
| 0    | Success |
| 1    | Failure |

## 🔎 Shadow Group (CLI)

**File**
```powershell
Shadow-Group\ShadowGroup.ps1
```

Synchronizes a group with the users of one or more OUs, with dry-run mode and a change cap.

| Code | Status |
| ---- | ------ |
| 0    | Success |
| 1    | Failure |

## 🖥️ Join Unjoin Computer Tool (GUI)

**File**
```powershell
JoinUnjoinComputerTool-GUI\JoinUnjoinComputerTool-v2.ps1
```

WPF GUI for domain join/disjoin, AD object deletion, Entra ID enrollment, and device snapshots.

| Code | Status |
| ---- | ------ |
| 0    | Normal exit after the user closes the window |

---

# ⚙️ Requirements

### Operating System
* Windows 10 / Windows 11 / Windows Server 2016+ (domain-joined)

### PowerShell
* PowerShell **5.1 or later**

### Permissions
* **RSAT: Active Directory module** (`ActiveDirectory`) on the machine running the tool
* Per-tool permissions are documented in each tool's own README. Most scripts are **read-only**; the disable/delete/add tools require write rights in AD

---

# 🛡 Operational Notes
* Each tool is **self-contained**: copy a folder and run the script without shared dependencies.
* Most scripts are **read-only**; always review the destructive tools (`DisableUsersInOU`, `Disable-InactiveComputers`, `Shadow-Group`) against a staging OU first.
* Apply the per-tool security requirements from each README. Never hard-code credentials in these scripts.
* The GUI tool (`JoinUnjoinComputerTool-GUI`) must be run from an **elevated** PowerShell session.

---

## 👤 Author

**Mohammad Abdelkader Omar**  
GitHub: [@mabdulkadr](https://github.com/mabdulkadr)  
Website: [momar.tech](https://momar.tech)

---

## 📜 License
This project is licensed under the [MIT License](https://opensource.org/licenses/MIT).

---

## ⚠ Disclaimer

This skill and every script it generates are provided as-is with no warranty of any kind. Test generated tools in a staging environment before deploying to production. The authors assume no liability for any damage or data loss resulting from their use.

---

<div align="center">

⭐ **If this collection saves you time, star the repo — it helps others find it.**

[Report an Issue](../../issues) · [momar.tech](https://momar.tech)

[![Buy Me A Coffee](https://img.shields.io/badge/Buy%20Me%20A%20Coffee-FFDD00?style=for-the-badge&logo=buy-me-a-coffee&logoColor=black)](https://www.buymeacoffee.com/mabdulkadrx)

Built with [**PowerShell Enterprise Admin**](https://github.com/mabdulkadr/powershell-enterprise-admin-skill)

</div>