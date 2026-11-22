<div align="center">

# 💻 Active Directory Reports

**Builds HTML reports for your whole Active Directory environment from one menu-driven console script.**

For administrators who want a timestamped snapshot of computers, domain controllers, workstations, servers, and account status without installing a reporting suite.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.0-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Usage](#-usage) • [Requirements](#%EF%B8%8F-requirements) • [License](#-license)

</div>

---

# 📖 Overview

Active Directory Reports is one PowerShell script that builds six HTML reports about an AD environment. A console menu lets you pick a report or run them all. Each run writes its output to a timestamped folder, which makes it easy to keep a history for audits.

The script runs in PowerShell and PowerShell ISE. Its only dependency is the RSAT Active Directory module.

---

# ✨ Features

* Six reports: computers, domain controllers, workstations, servers, account status, and OS grouping
* A numbered console menu, plus an option to run every report
* HTML with card layouts and sticky headers that also reads well on a phone
* Timestamped output under `C:\ADReports`, so earlier runs are kept
* Only the RSAT Active Directory module is required

---

# 📂 Project Structure

```text
ActiveDirectoryReports
│
├── ActiveDirectoryReports.ps1
└── README.md
```

---

# 🚀 Usage

### Basic Usage
```powershell
.\ActiveDirectoryReports.ps1
```

### Interactive Menu
```text
1 - Complete Computer Object Report
2 - Domain Controllers Report
3 - Workstations Report
4 - Servers Report
5 - Computer Account Status Report
6 - OS Based Report
7 - Run ALL Reports
```

---

# ⚙️ Parameters

This script takes **no parameters**. Report selection is driven through the interactive console menu.

| Variable | Type | Default | Description |
|----------|------|---------|-------------|
| `$CSS` / `Build-HtmlReport` | Scriptblock | *(embedded template)* | Edit to add your organization logo or brand colors |

---

# ⚙️ Requirements

### Operating System
* Windows Server or Windows 10/11 (domain-joined)

### PowerShell
* PowerShell **5.1 or later**

### Permissions
* **RSAT** with the **Active Directory PowerShell** module enabled
* Read permissions on Active Directory (**Domain Admin** recommended for full information)
* Run as Administrator for best results

---

## 🧰 Core Commands

The PowerShell cmdlets that drive each report type.

| Command | Purpose | Docs |
|---|---|---|
| `Get-ADComputer` | Enumerates computer objects for the computers, workstations, servers, and OS-based reports | [Microsoft Docs](https://learn.microsoft.com/en-us/powershell/module/activedirectory/get-adcomputer) |
| `Get-ADDomainController` | Discovers Domain Controllers for the DC report | [Microsoft Docs](https://learn.microsoft.com/en-us/powershell/module/activedirectory/get-addomaincontroller) |
| `ConvertTo-Html` | Renders each report into a styled, timestamped HTML file | [Microsoft Docs](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/convertto-html) |

---

# 🛡 Operational Notes
* All reports are saved to `C:\ADReports` and are timestamped, so multiple runs are preserved.
* The script is **read-only**: it never modifies Active Directory objects.
* Data freshness depends on AD replication; the report reflects the DC that answered the query.
* You can add an organization logo to reports by editing the `$CSS` and `Build-HtmlReport` section.

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