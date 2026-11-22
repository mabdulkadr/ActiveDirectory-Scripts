<div align="center">

# 💻 Get AD Users Details

**Retrieves detailed attributes for all users in an Organizational Unit, calculates inactivity in days, and exports a desktop CSV.**

For Administrators identifying stale or inactive accounts and producing user inventory reports.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.0-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Usage](#-usage) • [Requirements](#%EF%B8%8F-requirements) • [License](#-license)

</div>

---

# 📖 Overview

Get AD Users Details reads every user in a chosen OU and reports their name, logon name, enabled state, creation and modification dates, last logon, days inactive, and password age. The console lists the most inactive users first, and the same data goes to a CSV on the desktop for reporting or stale-account cleanup.

---

# ✨ Features

* An inventory limited to the OU you choose, with names, logon names, enabled state, and dates
* Days since last logon, including accounts that never logged in
* Console output sorted so the most inactive accounts appear first
* A desktop CSV with the full set of attributes
* Read-only: one query per user, with no changes to AD

---

# 📂 Project Structure

```text
Get-ADUsers-Details
│
├── Get-ADUsers-Details.ps1
└── README.md
```

---

# 🚀 Usage

### Configure the OU
```powershell
$OU = "OU=Operation Dept,DC=QassimU,DC=local"
```

### Run the Script
```powershell
.\Get-ADUsers-Details.ps1
```

---

# ⚙️ Parameters

This script takes **no parameters**. Set the `$OU` variable inside the script before running.

| Variable | Type | Default | Description |
|----------|------|---------|-------------|
| `$OU` | String | *(set by operator)* | Distinguished name of the Organizational Unit to inventory |

---

# ⚙️ Requirements

### Operating System
* Windows PowerShell host on a domain-joined machine

### PowerShell
* PowerShell **5.1 or later**

### Permissions
* **Active Directory** PowerShell module (RSAT Active Directory Tools)
* Read access to the target OU

---

## 🧰 Core Commands

The PowerShell cmdlets that drive the user inventory.

| Command | Purpose | Docs |
|---|---|---|
| `Get-ADUser` | Reads account attributes and logon timestamps per OU user | [Microsoft Docs](https://learn.microsoft.com/en-us/powershell/module/activedirectory/get-aduser) |
| `Get-ADGroup` | Resolves group membership details for each user | [Microsoft Docs](https://learn.microsoft.com/en-us/powershell/module/activedirectory/get-adgroup) |
| `Export-Csv` | Writes the desktop CSV report | [Microsoft Docs](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/export-csv) |

---

# 🛡 Operational Notes
* Output CSV: `C:\Users\<YourName>\Desktop\OperationDept_UserDetails.csv`.
* The script is **read-only**: it collects attributes and exports, and changes no accounts.
* Adjust the `$OU` variable for your environment before each run.
* Large OUs may take time on the first pass; results repeat identically until AD data changes.

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