<div align="center">

# 💻 Disable Users in OU

**Disables every enabled user account in a target Active Directory OU, logging each action and producing a summary plus optional CSV report.**

For Administrators executing bulk access revocations, offboarding, or OU cleanups.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.0-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Usage](#-usage) • [Requirements](#%EF%B8%8F-requirements) • [License](#-license)

</div>

---

# 📖 Overview

**Disable Users in OU** processes every **enabled** user account inside a configured Organizational Unit and disables it. It writes a formatted log for each operation, handles errors gracefully, and finishes with a summary block showing total accounts found, successfully disabled users, and failures. Set `$ExportCSV = $true` to also emit a CSV report of all attempted actions.

---

# ✨ Features

* **Bulk disable** — disables all enabled user accounts in the target OU
* **Per-user logging** — success and failure recorded per account with formatted output
* **Run summary** — totals for accounts found, disabled, and failed
* **Optional CSV export** — full report of attempted actions
* **Simple configuration** — set the `$OU` variable and run

---

# 📂 Project Structure

```text
DisableUsersInOU
│
├── DisableUsersInOU.ps1
└── README.md
```

---

# 🚀 Usage

### Configure
```powershell
$OU = "OU=Employees,OU=QU,DC=ABC,DC=local"
$ExportCSV = $true
```

### Run as Administrator
```powershell
.\DisableUsersInOU.ps1
```

---

# ⚙️ Parameters

This script takes **no parameters** — configure the variables at the top of the script.

| Variable | Type | Default | Description |
|----------|------|---------|-------------|
| `$OU` | String | *(set by operator)* | Distinguished name of the OU whose users will be disabled |
| `$ExportCSV` | Bool | `$false` | Toggle CSV report export to `C:\Reports` |

---

# ⚙️ Requirements

### Operating System
* Windows PowerShell host on a domain-joined machine

### PowerShell
* PowerShell **5.1 or later**

### Permissions
* **ActiveDirectory** module (RSAT)
* Permissions to disable user accounts in the target OU

---

# 🛡 Operational Notes
* Log file: `C:\Logs\DisableUsers_yyyy-MM-dd_HH-mm-ss.log`.
* CSV report (when enabled): `C:\Reports\DisabledUsers_yyyyMMdd_HHmmss.csv`.
* The action is **irreversible** for an affected login — users must be re-enabled manually if disabled by mistake.
* Review the OU variable carefully before running in production; consider a staging OU first.

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