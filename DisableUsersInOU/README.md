
# Disable Users in Active Directory OU

![PowerShell](https://img.shields.io/badge/powershell-5.1%2B-blue.svg)
![Status](https://img.shields.io/badge/status-stable-brightgreen.svg)
![License](https://img.shields.io/badge/license-MIT-blue.svg)

## Overview

This PowerShell script disables all **enabled** user accounts within a specified **Organizational Unit (OU)** in Active Directory.  
It logs every action, handles errors gracefully, and generates a detailed summary at the end of the run.  
Optionally, it can export a report of all processed accounts to a `.csv` file.

---

## Features

- Disables all **enabled** user accounts in the target OU
- Logs each operation (success/failure)
- Summarizes:
  - Total accounts found
  - Successfully disabled users
  - Failed operations
- Optionally exports a CSV report
- Easy to configure — just set the `$OU` variable

---

## How to Use

1. **Edit the Script**
   Open the script file and set the following variable:

   ```powershell
   $OU = "OU=Employees,OU=QU,DC=ABC,DC=local"
````

Optionally, toggle CSV export:

```powershell
$ExportCSV = $true
```

2. **Run as Administrator**

   Run the script in an elevated PowerShell session on a domain-joined system with RSAT tools installed:

   ```powershell
   .\DisableUsersInOU.ps1
   ```

---

## Output

### Console Output:

* Formatted logs for each user
* Summary block with totals

### Files Generated:

* `C:\Logs\DisableUsers_yyyy-MM-dd_HH-mm-ss.log` – Full operation log
* `C:\Reports\DisabledUsers_yyyyMMdd_HHmmss.csv` – (if `$ExportCSV = $true`) CSV report of all attempted actions

---

## Script Requirements

* PowerShell 5.1+
* RSAT: Active Directory PowerShell module (`ActiveDirectory`)
* Domain-joined machine with permissions to disable AD users

---

## Example Output

```text
[INFO] Starting user disable process in OU: OU=Employees,OU=QU,DC=ABC,DC=local
[SUCCESS] Disabled user: jdoe
[ERROR] Failed to disable user: test.user. Error: Access is denied.

======================= SUMMARY =======================
OU:                  OU=Employees,OU=QU,DC=ABC,DC=local
Total Users Found:   27
Successfully Disabled: 26
Failed to Disable:     1
CSV Report Saved To:  C:\Reports\DisabledUsers_20250625_130045.csv
=======================================================
```

---

## License

This project is licensed under the [MIT License](https://opensource.org/licenses/MIT).

---

## ⚠️ Disclaimer

This script is provided **as-is** without warranty.
The author is **not responsible** for unintended modifications or data loss.
Always test thoroughly before deploying in production.