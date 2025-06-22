
# Activate-Windows-KMS.ps1

![PowerShell](https://img.shields.io/badge/powershell-5.1%2B-blue.svg)
![Version](https://img.shields.io/badge/version-1.0-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)

## Overview

`Activate-Windows-KMS.ps1` is a PowerShell script that automates the activation of Windows using a specified **KMS (Key Management Service) server**. It checks server connectivity, configures the system to use the KMS host, initiates activation, and displays detailed license and activation status.

This tool is ideal for system administrators managing volume activation across domain-joined environments.

---

## Features

- 🔍 Checks TCP connectivity to the KMS server (port 1688)
- ⚙️ Configures the system to use a specified KMS server
- 🔐 Triggers Windows activation via KMS (`slmgr.vbs /ato`)
- 📊 Displays detailed activation status (`slmgr.vbs /dlv`)
- ✅ Validates administrative privileges before running

---

## Requirements

- PowerShell 5.1+
- Administrator privileges
- Access to the KMS server on port `1688`
- Supported Windows OS with a valid volume license key

---

## How to Use

### Step 1: Modify the KMS Server Name (if needed)

```powershell
$KMS_Server = "KMS.abc.local"
````

> 🔧 Change this value to match your environment's KMS host FQDN or IP address.

---

### Step 2: Run the Script as Administrator

```powershell
.\Activate-Windows-KMS.ps1
```

> 💡 If not run as Administrator, the script will exit with a warning.

---

## Sample Output

```
=== Step 1: Checking Connectivity to KMS Server ===
Success: Able to connect to KMS server 'KMS16.QassimU.local' on port 1688.

=== Step 2: Configuring Windows to Use the KMS Server ===
Executing: slmgr.vbs /skms KMS16.QassimU.local : 1688

=== Step 3: Activating Windows via KMS Server ===
Executing: slmgr.vbs /ato

=== Step 4: Retrieving Detailed Activation Information ===
[Activation ID, License Status, Remaining Grace, etc.]

=== Activation Process Completed Successfully ===
```

---

## Troubleshooting

| Issue                                    | Resolution                                                           |
| ---------------------------------------- | -------------------------------------------------------------------- |
| ❌ `Unable to connect to KMS server`      | Ensure the server is reachable and port 1688 is open.                |
| ⚠️ `Script must be run as Administrator` | Right-click PowerShell and select **Run as Administrator**.          |
| 🔒 Activation fails                      | Confirm the system has a valid KMS-compatible license key installed. |

---

## License

This script is licensed under the [MIT License](https://opensource.org/licenses/MIT)

---

## Disclaimer

> ⚠️ Use at your own risk. This script modifies Windows activation settings. Test in a lab before using in production environments.