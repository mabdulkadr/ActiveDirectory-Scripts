<div align="center">

# 💻 Shadow Group

**Keeps an Active Directory group in sync with one or more OUs, adding new users and removing those who left.**

For Administrators enforcing group-based policies (such as Fine-Grained Password Policies) that require exact OU-to-group alignment.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-3.0.0-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Usage](#-usage) • [Parameters](#%EF%B8%8F-parameters) • [Requirements](#%EF%B8%8F-requirements) • [License](#-license)

</div>

---

# 📖 Overview

**Shadow Group** keeps an Active Directory shadow group in exact sync with the users of one or more Organizational Units. It adds users who entered the OUs and removes users who no longer belong, writes every change to a detailed execution log, and caps bulk modifications per run to prevent excessive AD operations.

Run it first in log-only mode to preview changes, then enable modifications and schedule it as a Windows Task on a Domain Controller for hands-free reconciliation.

---

# ✨ Features

* Mirrors users from several OUs into one shadow group
* Can include only enabled accounts
* Includes or excludes child OUs based on the configuration
* A dry-run mode logs planned changes without touching AD
* A cap limits how many users are added or removed per run
* Built to run as a scheduled task on a domain controller

---

# 📂 Project Structure

```text
Shadow-Group
│
├── Sync-ShadowGroup.ps1
└── README.md
```

---

# 🚀 Usage

### Test Run (No Changes Applied)
```powershell
.\Sync-ShadowGroup.ps1
```

### Apply Changes
```powershell
# 1. Set $Update = $true inside the script
# 2. Then run:
.\Sync-ShadowGroup.ps1
```

### Schedule on a Domain Controller
```powershell
powershell.exe -ExecutionPolicy Bypass -File "C:\Path\To\Sync-ShadowGroup.ps1"
```

---

# ⚙️ Parameters

All settings live in the **PARAMETERS** section at the top of the script; the script takes no command-line parameters.

| Variable | Type | Description |
|----------|------|-------------|
| `$Server` | String | Domain Controller handling AD queries |
| `$LogFile` | String | Path for execution logs (default `C:\ShadowGroups\ShadowGroup.log`) |
| `$Update` | Bool | `$true` applies changes, `$false` logs only |
| `$EnabledOnly` | Bool | Include only enabled users in the shadow group |
| `$OUDNs` | String[] | Array of OUs to monitor |
| `$ChildOUs` | Bool | Whether to include child OUs |
| `$GroupDN` | String | Distinguished Name of the shadow group |
| `$MaxChanges` | Int | Maximum users added/removed per execution |

---

# ⚙️ Requirements

### Operating System
* Windows Server **Domain Controller** (recommended) or domain-joined admin host

### PowerShell
* PowerShell **5.1 or later**

### Permissions
* **Active Directory Module** for Windows PowerShell
* **ADWS** (Active Directory Web Services) on the target DC
* Rights to manage group membership in AD (run elevated)

---

## 🧰 Core Commands

The PowerShell cmdlets that drive the shadow-group sync.

| Command | Purpose | Docs |
|---|---|---|
| `Get-ADUser` | Enumerates users in the monitored OUs | [Microsoft Docs](https://learn.microsoft.com/en-us/powershell/module/activedirectory/get-aduser) |
| `Get-ADGroupMember` | Reads the current shadow group membership for comparison | [Microsoft Docs](https://learn.microsoft.com/en-us/powershell/module/activedirectory/get-adgroupmember) |
| `Add-ADGroupMember` | Adds users who entered the monitored OUs | [Microsoft Docs](https://learn.microsoft.com/en-us/powershell/module/activedirectory/add-adgroupmember) |
| `Remove-ADGroupMember` | Removes users who no longer belong to the OUs | [Microsoft Docs](https://learn.microsoft.com/en-us/powershell/module/activedirectory/remove-adgroupmember) |

---

# 🛡 Operational Notes
* Create the shadow group before the first run. The script stops with an error if it is missing
* Changes beyond `$MaxChanges` wait for the next run
* The log records added and removed users, timestamps, and any warnings or errors
* For a scheduled task, set "Run whether user is logged on or not" and "Run with highest privileges"

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