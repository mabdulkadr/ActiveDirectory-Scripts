
# Disable and Delete Inactive AD Computers

![PowerShell](https://img.shields.io/badge/powershell-5.1%2B-blue.svg)
![Status](https://img.shields.io/badge/status-stable-brightgreen.svg)
![License](https://img.shields.io/badge/license-MIT-blue.svg)

## Overview

This PowerShell script identifies **inactive computer objects** in Active Directory based on their last logon date, and optionally:
- Disables them
- Deletes them

It then generates a **timestamped HTML report** showing:
- Inactive computers
- Disabled computers (if any)
- Deleted computers (if any)

The report is styled and saved locally for audit and documentation purposes.

---

## Features

- Search computers in a specified Organizational Unit (OU)
- Identify computers inactive for a configurable number of days
- Prompt before disabling or deleting
- Generate a full HTML report with clean formatting and color-coded tables
- Optional console and log file output

---

## Parameters

| Name           | Type   | Description                                                                 |
|----------------|--------|-----------------------------------------------------------------------------|
| `DaysInactive` | `int`  | Number of days of inactivity used to determine stale computer accounts.     |
| `SearchBaseOU` | `string` | The distinguished name of the OU to scan (e.g., `OU=Domain Computers,DC=contoso,DC=local`). |

Default values:
```powershell
$DaysInactive    = 180
$SearchBaseOU    = "OU=Domain Computers,DC=ABC,DC=local"
````

---

## How It Works

1. Retrieves all computers from the specified OU.
2. Filters out computers that haven't logged on within the past `DaysInactive`.
3. Displays and optionally disables inactive computers.
4. Prompts again and optionally deletes previously disabled computers.
5. Generates a styled HTML report showing:

   * Inactive computers
   * Disabled computers
   * Deleted computers

---

## Output

* ✅ **HTML Report**

  * Path: `C:\Computers_Report_YYYY-MM-DD_HH-mm-ss.html`
  * Contains 3 tables (Inactive, Disabled, Deleted)
  * Includes counts and timestamps

* ✅ **Console Output**

  * Shows each action and prompt
  * Color-coded for easy readability

---

## Example

```powershell
.\DisableInactiveComputers.ps1 -DaysInactive 180 -SearchBaseOU "OU=Domain Computers,DC=ABC,DC=local"
```

---

## Requirements

* Windows PowerShell 5.1+
* ActiveDirectory module (RSAT)
* Domain-joined machine
* Permissions to query, disable, and delete computer objects in AD

---

## Recommended Usage

* Run in a test/staging OU before using in production
* Schedule regularly to manage stale computer accounts
* Back up AD before deletion (optional)

---

## License

This project is licensed under the [MIT License](https://opensource.org/licenses/MIT).

---

## ⚠️ Disclaimer

This script is provided **as-is** without warranty.
The author is **not responsible** for unintended modifications or data loss.
Always test thoroughly before deploying in production.