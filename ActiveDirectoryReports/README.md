<div align="center">

# 💻 Active Directory Reports

**Generates modern, responsive HTML reports for your entire Active Directory environment from a single menu-driven console script.**

For Administrators who need a clean, timestamped snapshot of computers, domain controllers, workstations, servers, and account status — without installing reporting suites.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.0-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Usage](#-usage) • [Requirements](#%EF%B8%8F-requirements) • [License](#-license)

</div>

---

# 📖 Overview

**Active Directory Reports** is a single PowerShell script that produces modern, responsive, card-based HTML reports for an Active Directory environment. A console menu lets you pick one of six built-in reports — or run all of them — with every report saved into a styled, timestamped folder for audit and documentation purposes.

The script works in both classic PowerShell and PowerShell ISE, depends only on the RSAT Active Directory module, and needs no additional tooling.

---

# ✨ Features

* **Six built-in reports** — computers, domain controllers, workstations, servers, account status, and OS-based grouping
* **Menu-driven interface** — type a number to pick a report, or select "Run ALL Reports"
* **Modern HTML output** — responsive, mobile-friendly, card-based layouts with sticky headers
* **Timestamped results** — every run is preserved in `C:\ADReports`
* **Minimal prerequisites** — only the RSAT Active Directory PowerShell module

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

This script takes **no parameters** — all report selection is driven through the interactive console menu.

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
* The script is **read-only** — it never modifies Active Directory objects.
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