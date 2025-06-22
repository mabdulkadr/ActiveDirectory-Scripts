
# AD Groups Report

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![PowerShell](https://img.shields.io/badge/powershell-5.1%2B-blue.svg)
![Version](https://img.shields.io/badge/version-1.1-green.svg)

## Overview

**AD-GroupsReport.ps1** is a powerful PowerShell script that generates a detailed CSV report of all Active Directory groups in your environment. It includes comprehensive metadata, member statistics, nested group insights, and logs all operations. The script automatically checks and installs the required **ActiveDirectory** module if missing.

---

## Features

- ✅ Auto-installs **RSAT Active Directory Tools** if needed
- ✅ Retrieves **all AD group details** with:
  - Name, Display Name, Distinguished Name
  - Email, Description
  - Group Type (Scope + Category)
  - Creation and Modification Dates
  - Inactivity Flags
  - Member counts by type (Users, Groups, Contacts)
  - Nested Group names and details

- 🪵 Logs each step to a timestamped log file
- 📁 Creates output folder automatically
- 🎨 Color-coded console output (INFO, SUCCESS, WARNING, ERROR)

---

## Requirements

- Windows PowerShell 5.1 or later
- ActiveDirectory PowerShell module (installed via RSAT)
- Run as Administrator
- Domain-joined device with AD read access

---

## Usage

```powershell
# Run in elevated PowerShell session
.\AD-GroupsReport.ps1
````

### Output Location

```
C:\ADGroupsReport\
├── ADGroupsReport_<timestamp>.csv  # Main report
└── ADGroupsExport_<timestamp>.log  # Execution log
```

---

## Report Fields

| Column                | Description                                               |
| --------------------- | --------------------------------------------------------- |
| GroupName             | sAMAccountName of the group                               |
| DisplayName           | Display name (if configured)                              |
| DistinguishedName     | Full DN of the group                                      |
| GroupType             | Combination of GroupScope and GroupCategory               |
| GroupEmailAddress     | Email address (if configured)                             |
| Description           | Group description                                         |
| WhenCreated           | Creation timestamp                                        |
| WhenChanged           | Last modification timestamp                               |
| NeverModified         | Yes if `WhenChanged == WhenCreated`                       |
| IsLikelyInactive      | Yes if empty and not modified for over a year             |
| MembersCount          | Total number of direct members                            |
| MembersCountByType    | Breakdown by User, Group, Contact                         |
| NestedGroupCount      | Count of directly nested groups                           |
| NestedGroupNames      | Comma-separated nested group names                        |
| NestedGroups\_Details | Detailed view of nested groups (Name, Type, Email, Count) |

---

## Sample Row

```csv
"HR-Global","HR Team","CN=HR-Global,OU=Groups,DC=contoso,DC=com","Global-Security","hr@contoso.com","HR Staff","2022-01-01","2024-01-01","No","No",15,"User - 10; Group - 5; Contact - 0",2,"IT-Admins, Finance-Users","IT-Admins:Global-Security:it@contoso.com:8 | Finance-Users:DomainLocal-Security:fin@contoso.com:6"
```

---

## Troubleshooting

* ❗ **ActiveDirectory module not found**
  → The script will attempt automatic installation. If it fails, install RSAT manually from:
  `Settings → Apps → Optional Features → Add a feature → RSAT: Active Directory DS and LDS Tools`

* ❗ **Access Denied**
  → Ensure you're running PowerShell **as Administrator**

---

## License

Licensed under the [MIT License](https://opensource.org/licenses/MIT)

---

## Disclaimer

> ⚠️ This script is provided as-is. Test it in a development environment before using in production. The author is not responsible for unintended consequences.
