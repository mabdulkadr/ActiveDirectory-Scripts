<div align="center">

# 💻 AD Groups Report

**Exports a detailed CSV report of all Active Directory groups with membership, nesting, and inactivity data.**

For IT teams that need a full inventory of AD groups for auditing, cleanup, and license assessments.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.0-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Usage](#-usage) • [Parameters](#%EF%B8%8F-parameters) • [Requirements](#%EF%B8%8F-requirements) • [License](#-license)

</div>

---

# 📖 Overview

**AD Groups Report** enumerates all groups in an Active Directory domain and exports a timestamped CSV containing group metadata, member-type counts (User, Group, Contact), nested group names, and inactivity flags computed from the last group modification date.

The script optionally self-installs the RSAT **ActiveDirectory** tools when the module is missing, so the report runs on a fresh admin workstation without manual prerequisites.

---

# ✨ Features

* **Full group inventory** — name, DN, group type, email, description, and locality attributes
* **Inactivity flags** — `NeverModified` and `IsLikelyInactive` derived from modification timestamps
* **Member-type counts** — breakdown of members by User, Group, and Contact for each group
* **Nesting details** — nested group names and detail strings included per group
* **CSV + log export** — timestamped report and a matching execution log

---

# 📂 Project Structure

```text
AD-GroupsReport
│
├── AD-GroupsReport.ps1
└── README.md
```

---

# 🚀 Usage

### Basic Usage
```powershell
.\AD-GroupsReport.ps1
```

### Self-Installed RSAT Warning
```powershell
# First run may attempt to install RSAT ActiveDirectory tools - requires elevation
```

---

# ⚙️ Parameters

This script takes **no parameters**. Replace the hardcoded configuration in the variables block at the top of the script when you need to change paths.

| Variable | Type | Default | Description |
|----------|------|---------|-------------|
| `$ReportFolder` | String | `C:\ADGroupsReport` | Folder for the CSV report and log file |

---

# ⚙️ Requirements

### Operating System
* Windows Server 2012 R2+ or Windows 10/11 (domain-joined)

### PowerShell
* PowerShell **5.1 or later**

### Permissions
* **ActiveDirectory** module (RSAT) — the script attempts to install it when missing
* Read access to Active Directory groups
* Run from an **elevated** PowerShell session

---

## 🧰 Core Commands

The PowerShell cmdlets that drive the group report.

| Command | Purpose | Docs |
|---|---|---|
| `Get-ADGroup` | Enumerates all groups in the domain with metadata attributes | [Microsoft Docs](https://learn.microsoft.com/en-us/powershell/module/activedirectory/get-adgroup) |
| `Get-ADObject` | Resolves group member objects to compute member-type counts and nesting | [Microsoft Docs](https://learn.microsoft.com/en-us/powershell/module/activedirectory/get-adobject) |
| `Export-Csv` | Writes the timestamped CSV report | [Microsoft Docs](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/export-csv) |

---

# 🛡 Operational Notes
* Output lands in `C:\ADGroupsReport\ADGroupsReport_<timestamp>.csv` with a matching `.log` file.
* The script is **read-only** — it queries groups and members but never modifies objects.
* In very large domains the first pass over all groups can take several minutes; run it outside business hours.
* Module installation requires elevation — run elevated or pre-install RSAT to avoid a stalled run.

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

⭐ **If this skill saves you time, star the repo — it helps others find it.**

[Report an Issue](../../issues) · [momar.tech](https://momar.tech)

[![Buy Me A Coffee](https://img.shields.io/badge/Buy%20Me%20A%20Coffee-FFDD00?style=for-the-badge&logo=buy-me-a-coffee&logoColor=black)](https://www.buymeacoffee.com/mabdulkadrx)

Built with [**PowerShell Enterprise Admin**](https://github.com/mabdulkadr/powershell-enterprise-admin-skill)

</div>