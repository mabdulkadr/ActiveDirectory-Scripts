
# Active Directory Health Check

![PowerShell](https://img.shields.io/badge/powershell-5.1%2B-blue.svg)
![Version](https://img.shields.io/badge/version-2.10-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)

## Overview

`Get-ADHealth.ps1` is a comprehensive PowerShell script for auditing and validating the health of all Domain Controllers in a specific Active Directory domain or the entire forest.

The script runs a series of diagnostics such as service availability, time synchronization, disk space, replication issues, and DCDIAG test results. The output is a structured, color-coded **HTML report** that provides at-a-glance status for each DC, and optionally sends it via **email with an HTML body and attachment**.

---

## Features

- 🔍 Health checks for all DCs in a domain or forest
- ✅ Tests include:
  - DNS resolution (Resolve-DnsName)
  - Ping status
  - Server uptime
  - Time offset check via w32tm
  - Disk space (GB and percentage)
  - Key service status (DNS, NTDS, NetLogon)
  - Full DCDIAG test suite
- 📄 HTML report with color-coded status
- 📧 SMTP email support with the report as both **attachment** and **HTML body**
- 📁 Saves reports to `C:\Reports` automatically
- 🧠 Modular functions and clean output formatting
- 🛠 Ready for scheduled tasks or automation tools

---

## Prerequisites

- ✅ PowerShell 5.1+
- ✅ RSAT installed (Active Directory module, DCDIAG)
- ✅ Domain Admin permissions (or delegated rights)
- ✅ Network access to all DCs
- ✅ SMTP access if using `-SendEmail`

---

## How to Use

### Basic Syntax

```powershell
.\Get-ADHealth.ps1 [-DomainName <domain>] [-ReportFile] [-SendEmail]
````

### Examples

```powershell
# Run forest-wide health check and save report
.\Get-ADHealth.ps1 -ReportFile

# Check a specific domain
.\Get-ADHealth.ps1 -DomainName contoso.com -ReportFile

# Check domain and send report via email
.\Get-ADHealth.ps1 -DomainName contoso.com -SendEmail
```

---

## Parameters

| Parameter     | Description                                                               |
| ------------- | ------------------------------------------------------------------------- |
| `-DomainName` | (Optional) Target a specific domain. If omitted, all domains are checked. |
| `-ReportFile` | Save the HTML report locally to `C:\Reports`.                             |
| `-SendEmail`  | Send the report via email (requires SMTP configuration).                  |

---

## SMTP Configuration

Update this block in the script:

```powershell
$smtpsettings = @{
    To         = 'admin@domain.com'
    From       = 'adhealth@yourdomain.com'
    Subject    = "Domain Controller Health Report - $(Get-Date -Format 'yyyy-MM-dd')"
    SmtpServer = "mail.domain.com"
    Port       = 25
    #Credential = (Get-Credential)
    #UseSsl     = $true
}
```

> 💡 You can enable authentication and SSL by uncommenting `Credential` and `UseSsl`.

---

## Output

* 📁 Saved to: `C:\Reports\ADHealthReport_<timestamp>.html`
* 📧 Sent via SMTP with:

  * Inline HTML body
  * Attached report file

Report statuses are styled:

| Status    | Meaning       |
| --------- | ------------- |
| ✅ Green   | Passed check  |
| ⚠️ Yellow | Warning state |
| ❌ Red     | Failed check  |

---

## Screenshot

> A screenshot preview of the HTML table can be added here in your GitHub repo if desired.

---

## License

Licensed under the [MIT License](https://opensource.org/licenses/MIT).

---

## Disclaimer

> ⚠️ Use this script at your own risk. Test thoroughly in a lab environment before deploying in production. The author assumes no responsibility for system changes or failures.
