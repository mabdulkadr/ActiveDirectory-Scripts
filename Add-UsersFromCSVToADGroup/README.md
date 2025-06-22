
# Add Users From CSV To Group in AD

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![PowerShell](https://img.shields.io/badge/powershell-5.1%2B-blue.svg)
![Version](https://img.shields.io/badge/version-1.0-green.svg)

## Overview

`Add-UsersFromCSVToGroup.ps1` is a PowerShell script that automates the process of adding multiple users to an Active Directory group. It reads user IDs from a CSV file, appends a defined domain suffix to create UPNs, verifies each user in AD, checks their membership, and adds them to the specified group only if not already present.

---

## Features

- ✅ Reads a list of users from a CSV file
- ✅ Appends a domain suffix to form each user's UPN
- ✅ Verifies if each user exists in Active Directory
- ✅ Checks if the user is already a group member
- ✅ Adds only non-members to the group
- ✅ Displays group metadata and detailed summary
- 🎨 Color-coded output (Added, Already Exists, Failed)

---

## Requirements

- Windows PowerShell 5.1+
- Active Directory PowerShell module (`RSAT-AD-PowerShell`)
- Administrator privileges
- Access to the target AD group

---

## CSV Format

The CSV must contain a single column named `UserID`:

```csv
UserID
m.omar
a.ahmed
````

---

## How to Use

1. Open PowerShell as Administrator
2. Customize these parameters in the script:

   ```powershell
   $GroupName = "AD-Group-Name"
   $CsvPath = "C:\Add-UsersFromCSVToGroup\users.csv"
   $Domain   = "@abc.local"
   ```
3. Run the script:

   ```powershell
   .\Add-UsersFromCSVToGroup.ps1
   ```

---

## Output

Displays:

* ✅ Added users
* ⚠️ Users already in the group
* ❌ Failed additions

### Example Summary:

```
============= Summary =============
✅ Total Users Added      : 3
⚠️  Already in Group       : 2
❌ Failed Attempts         : 1
====================================
```

---

## License

Licensed under the [MIT License](https://opensource.org/licenses/MIT)

---

## Disclaimer

> ⚠️ Always test scripts in a lab environment before production use. The author is not responsible for any misuse or data loss.
