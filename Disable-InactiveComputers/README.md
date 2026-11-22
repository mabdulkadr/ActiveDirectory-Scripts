<div align="center">

# 💻 Disable Inactive Computers

**Identifies stale computer accounts by last logon and optionally disables or deletes them, producing a color-coded HTML audit report.**

For Administrators who need to clean up old computer objects in Active Directory with a visible audit trail.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.0-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Usage](#-usage) • [Parameters](#%EF%B8%8F-parameters) • [Requirements](#%EF%B8%8F-requirements) • [License](#-license)

</div>

---

# 📖 Overview

**Disable Inactive Computers** scans a configurable Organizational Unit, finds computer accounts that have not logged on within a set number of days, and optionally disables and then deletes them. Every run produces a timestamped HTML report with three color-coded tables (**Inactive**, **Disabled**, and **Deleted**) and color-coded console output for immediate review.

---

# ✨ Features

* Searches only the Organizational Unit you specify
* Sets the inactivity window in days with `DaysInactive`
* Uses two prompts: the first disables stale computers, the second deletes them
* Writes three color-coded HTML tables with counts and timestamps
* Logs color-coded actions to the console and, optionally, to a file

---

# 📂 Project Structure

```text
Disable-InactiveComputers
│
├── Disable-ADInactiveComputer.ps1
└── README.md
```

---

# 🚀 Usage

### Basic Usage
```powershell
.\Disable-ADInactiveComputer.ps1
```

### With Parameters
```powershell
.\Disable-ADInactiveComputer.ps1 -DaysInactive 180 -SearchBaseOU "OU=Domain Computers,DC=contoso,DC=local"
```

---

# ⚙️ Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `DaysInactive` | Int | No | `180` | Number of days of inactivity used to flag stale computer accounts |
| `SearchBaseOU` | String | No | `OU=Domain Computers,DC=ABC,DC=local` | Distinguished name of the OU to scan |

---

# ⚙️ Requirements

### Operating System
* Windows PowerShell host on a domain-joined machine

### PowerShell
* PowerShell **5.1 or later**

### Permissions
* **ActiveDirectory** module (RSAT)
* Permissions to query, disable, and delete computer objects in AD

---

## 🧰 Core Commands

The PowerShell cmdlets that drive the cleanup workflow.

| Command | Purpose | Docs |
|---|---|---|
| `Get-ADComputer` | Discovers computer accounts in the target OU with `lastLogon` | [Microsoft Docs](https://learn.microsoft.com/en-us/powershell/module/activedirectory/get-adcomputer) |
| `Disable-ADAccount` | Disables computer accounts flagged as inactive | [Microsoft Docs](https://learn.microsoft.com/en-us/powershell/module/activedirectory/disable-adaccount) |
| `Remove-ADComputer` | Deletes computer accounts in the final cleanup phase | [Microsoft Docs](https://learn.microsoft.com/en-us/powershell/module/activedirectory/remove-adcomputer) |

---

# 🛡 Operational Notes
* The HTML report is saved to `C:\Computers_Report_yyyy-MM-dd_HH-mm-ss.html` with tables for inactive, disabled, and deleted computers.
* Run against a **test/staging OU** before using in production.
* Schedule the script regularly to keep stale computer accounts in check.
* Back up Active Directory before running the delete phase; removal is irreversible.

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

Built with [**PowerShell Enterprise Admin**](https://github.com/mabdulkadr/powershell-enterprise-admin)

</div>