
# AD GroupReport

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![PowerShell](https://img.shields.io/badge/powershell-5.1%2B-blue.svg)
![Version](https://img.shields.io/badge/version-1.0-green.svg)

## Overview

`AD-GroupReport.ps1` is a comprehensive PowerShell script that exports a detailed report of **all Active Directory groups**. The report includes group metadata, nested group analysis, member type breakdowns, and timestamps. It logs each step and outputs results to a timestamped CSV file and log file.

---

## Features

- Exports a **single unified CSV** report of all AD groups
- Includes:
  - Group name, DN, display name, type, description
  - Email, creation/change timestamps
  - Nested group names and member counts
  - Member type counts (users, groups, contacts)
  - Activity indicators (`IsLikelyInactive`, `NeverModified`)
- Robust logging for success, warnings, and errors
- Automatically creates output directory

---

## Requirements

- Windows PowerShell 5.1+
- Active Directory PowerShell Module (`RSAT-AD-PowerShell`)
- Run with domain user privileges that can read AD group membership

---

## How to Use

```powershell
# Open PowerShell as Administrator
.\AD-GroupReport.ps1
````

> 📁 Output location: `C:\ADGroupReports\`
>
> * CSV report: `ADGroupUnifiedReport_<timestamp>.csv`
> * Log file: `ADGroupExport_<timestamp>.log`

---

## CSV Output Columns

| Column Name           | Description                                              |
| --------------------- | -------------------------------------------------------- |
| GroupName             | Name of the group                                        |
| DisplayName           | Friendly display name (if set)                           |
| DistinguishedName     | Full DN of the group                                     |
| GroupType             | Scope and category (e.g., Global-Security)               |
| GroupEmailAddress     | Group email (if set)                                     |
| Description           | AD description field                                     |
| WhenCreated           | Creation timestamp                                       |
| WhenChanged           | Last modified timestamp                                  |
| NeverModified         | `Yes` if unchanged since creation                        |
| IsLikelyInactive      | `Yes` if no members and unchanged for over 1 year        |
| MembersCount          | Total number of direct members                           |
| MembersCountByType    | Breakdown by User / Group / Contact                      |
| NestedGroupCount      | Number of directly nested groups                         |
| NestedGroupNames      | Comma-separated names of nested groups                   |
| NestedGroups\_Details | Name, type, email, and count of members in nested groups |

---

## Example Output

```csv
GroupName,DisplayName,GroupType,MembersCount,IsLikelyInactive
"HR-Global","HR Global Team","Global-Security",5,"No"
"IT-Archive","","DomainLocal-Security",0,"Yes"
...
```

---

## License

This project is licensed under the [MIT License](https://opensource.org/licenses/MIT).

---

**Disclaimer**: Always test scripts in a development environment before deploying them in production. The author is not responsible for any unintended consequences.
