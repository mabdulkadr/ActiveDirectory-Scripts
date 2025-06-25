
# Get-ADUsers-Details.ps1

![PowerShell](https://img.shields.io/badge/powershell-5.1%2B-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Version](https://img.shields.io/badge/version-1.0-blue.svg)

## Overview

`Get-ADUsers-Details.ps1` is a PowerShell script that retrieves comprehensive information about all Active Directory users within a specified Organizational Unit (OU). It calculates inactivity in days and displays key user account attributes such as creation, modification, last login, and password last set date.

This script is ideal for identifying stale or inactive accounts and for reporting purposes.

---

## Features

- 🔍 Collects detailed user attributes from a specified OU:
  - `Name`, `SamAccountName`, `Enabled`
  - `whenCreated`, `whenChanged`
  - `LastLogonDate`, `DaysInactive`
  - `PasswordLastSet`
- 📊 Calculates the number of **inactive days** since last logon
- 📁 Exports results to CSV on the desktop
- 🖥 Displays a sorted table in the console (most inactive first)

---

## Requirements

- Windows PowerShell 5.1+
- Active Directory PowerShell module (`RSAT: Active Directory Tools`)
- Domain-joined machine with read access to AD

---

## How to Use

### Step 1: Set OU Path

Edit the `$OU` variable inside the script:

```powershell
$OU = "OU=Operation Dept,DC=QassimU,DC=local"
````

> 💡 Adjust the OU as needed for your environment.

### Step 2: Run the Script

```powershell
.\Get-ADUsers-Details.ps1
```

---

## Output

### Console Output

Displays a table of users sorted by `DaysInactive` (descending):

```
Name         SamAccountName  Enabled  LastLogonDate  DaysInactive  PasswordLastSet
----         --------------  -------  -------------- ------------- ----------------
John Smith   jsmith          True     2024-11-01      233           2024-10-30
Jane Doe     jdoe            True     Never Logged In Never Logged In 2023-08-15
...
```

### CSV Export

A report file is saved to:

```
C:\Users\<YourName>\Desktop\OperationDept_UserDetails.csv
```

> ✅ Includes all attributes listed above.

---

## Example

```powershell
.\Get-ADUsers-Details.ps1
```

This command will generate a full user report for the OU `"OU=Operation Dept,DC=QassimU,DC=local"` and export it to your desktop.

---

## License

This script is licensed under the [MIT License](https://opensource.org/licenses/MIT)

---

## Disclaimer

> ⚠️ Use responsibly in production. Always test scripts in a staging environment first.
