<div align="center">

# 💻 Add-UsersFromCSVToADGroup

**Bulk-adds users to an Active Directory group from a CSV file, skipping users who are already members.**

For Administrators who need to onboard batch users into a group without duplicates or manual ADUC clicks.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.0-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Usage](#-usage) • [Parameters](#%EF%B8%8F-parameters) • [Requirements](#%EF%B8%8F-requirements) • [License](#-license)

</div>

---

# 📖 Overview

**Add-UsersFromCSVToADGroup** reads a list of user IDs from a CSV file, resolves each ID against Active Directory, verifies it has not already been added to the target group, and adds only the missing members. A color-coded console summary reports **Added**, **Already in Group**, and **Failed** results after the run.

---

# ✨ Features

* **CSV input** — reads a `UserID` column from the supplied CSV file
* **UPN resolution** — appends the configured domain suffix to build a full user principal name
* **Duplicate avoidance** — checks existing membership before every add
* **Pre-add verification** — confirms the user object exists in AD before attempting the add
* **Color-coded summary** — Added / Already in Group / Failed counts on completion

---

# 📂 Project Structure

```text
Add-UsersFromCSVToADGroup
│
├── Add-UsersFromCSVToADGroup.ps1
├── users.csv
└── README.md
```

---

# 🚀 Usage

### Basic Usage
```powershell
.\Add-UsersFromCSVToADGroup.ps1
```

### CSV Format
```text
UserID
john.doe
jane.smith
```

---

# ⚙️ Parameters

This script takes **no parameters**. Edit the configuration variables at the top of the script before running.

| Variable | Type | Default | Description |
|----------|------|---------|-------------|
| `$GroupName` | String | *(set by operator)* | Name of the target AD group |
| `$CsvPath` | String | *(set by operator)* | Path to the CSV file containing the `UserID` column |
| `$Domain` | String | *(set by operator)* | Domain suffix appended to user IDs to form UPNs |

---

# ⚙️ Requirements

### Operating System
* Windows Server 2012 R2+ or Windows 10/11 (domain-joined)

### PowerShell
* PowerShell **5.1 or later**

### Permissions
* **RSAT-AD-PowerShell** module
* **Administrator** privileges
* Write access to the target AD group

---

## 🧰 Core Commands

The PowerShell cmdlets that drive the bulk-add workflow.

| Command | Purpose | Docs |
|---|---|---|
| `Import-Csv` | Reads the `UserID` column from the input CSV | [Microsoft Docs](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/import-csv) |
| `Get-ADUser` | Resolves each user ID and verifies the account exists | [Microsoft Docs](https://learn.microsoft.com/en-us/powershell/module/activedirectory/get-aduser) |
| `Get-ADGroupMember` | Reads existing membership to avoid duplicates | [Microsoft Docs](https://learn.microsoft.com/en-us/powershell/module/activedirectory/get-adgroupmember) |
| `Add-ADGroupMember` | Adds only the users not already in the group | [Microsoft Docs](https://learn.microsoft.com/en-us/powershell/module/activedirectory/add-adgroupmember) |

---

# 🛡 Operational Notes
* The CSV **must** contain a `UserID` column; any other column name produces zero matched users.
* Set `$GroupName`, `$CsvPath`, and `$Domain` before running — the script terminates on a missing group.
* Verify the summary counts match expectations; users reported as **Failed** are never partially added.
* The operation is **additive only** — no members are ever removed by this script.

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