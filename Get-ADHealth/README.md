<div align="center">

# 💻 Get AD Health

**Runs a comprehensive health check against all Domain Controllers — DNS, reachability, replication, disk, time sync — and delivers a card-style HTML report, optionally by email.**

For IT operations teams that need a daily heartbeat of their Active Directory fleet.

[![Mode](https://img.shields.io/badge/Mode-CLI-334155?style=for-the-badge)](#-usage)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0F172A?style=for-the-badge)](#%EF%B8%8F-requirements)
[![License](https://img.shields.io/badge/License-MIT-F59E0B?style=for-the-badge)](#-license)
[![Version](https://img.shields.io/badge/Version-1.0.0-334155?style=for-the-badge)](#-overview)

[Overview](#-overview) • [Features](#-features) • [Usage](#-usage) • [Parameters](#%EF%B8%8F-parameters) • [Requirements](#%EF%B8%8F-requirements) • [License](#-license)

</div>

---

# 📖 Overview

**Get AD Health** enumerates every Domain Controller in one or more domains (or the entire forest), probes each one across DNS, ICMP, uptime, time synchronization, OS disk space, critical services (DNS, NTDS, NetLogon), and the `DCDIAG` suite, then summarizes each DC as **Healthy**, **Warning**, or **Critical**.

Results are rendered as modern, email-safe HTML cards — saved to disk and optionally sent inline via SMTP so the on-call team sees the same layout in the report file and the email body.

---

# ✨ Features

* **Per-DC collectors** — DNS, ping, uptime, time drift, disk %, services, and `DCDIAG`
* **Healthy / Warning / Critical** verdicts from binary checks and tunable numeric thresholds
* **Forest or targeted scope** — scan the whole forest or specific domains
* **Card-style HTML** — email-safe inline CSS with badges and meters
* **SMTP delivery** — same layout inline in email and on disk, ready for Task Scheduler

---

# 📂 Project Structure

```text
Get-ADHealth
│
├── Get-ADHealth.ps1
├── Screenshot.png
└── README.md
```

---

# 🚀 Usage

### Scan the Entire Forest
```powershell
.\Get-ADHealth.ps1 -ReportPath 'C:\Reports'
```

### Scope to Specific Domains
```powershell
.\Get-ADHealth.ps1 -DomainName 'contoso.com','emea.contoso.com' -ReportPath 'C:\Reports'
```

### Email the Report Inline
```powershell
$cred = Get-Credential
.\Get-ADHealth.ps1 -SendEmail -Subject 'AD DC Health Check' `
  -UserFrom 'op-smtp@yourtenant.onmicrosoft.com' `
  -UserTo 'it-admins@yourdomain.com' `
  -SmtpServer 'smtp.office365.com' -Port 587 -Credential $cred
```

---

# ⚙️ Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `DomainName` | String[] | No | *(all forest domains)* | One or more domain DNS names to scan; omit to scan the entire forest |
| `ReportPath` | String | No | `C:\Reports` | Folder to save the HTML report |
| `SendEmail` | Switch | No | `False` | Sends the report inline via SMTP |
| `Subject` | String | No | `Domain Controller Health Check` | Email subject |
| `UserFrom` | String | No | `smtp-reports@yourdomain.com` | SMTP From address |
| `UserTo` | String[] | No | `it-admins@yourdomain.com` | One or more recipients |
| `SmtpServer` | String | No | `smtp.office365.com` | SMTP server hostname |
| `Port` | Int | No | `587` | SMTP port |
| `Credential` | PSCredential | No | — | SMTP credential; built from `UserFrom` + `Password` when omitted |
| `Password` | SecureString | No | — | SMTP password used when `Credential` is not supplied |

---

# ⚙️ Requirements

### Operating System
* Windows PowerShell 5.1 host (Windows Server or domain-joined client)

### PowerShell
* PowerShell **5.1**

### Permissions
* **RSAT: Active Directory** PowerShell module
* `dcdiag.exe` available (RSAT or run on a DC)
* Remote WMI/CIM access to DCs for uptime and disk metrics
* Optional SMTP server for email delivery

---

## 🧰 Core Commands

The PowerShell cmdlets that drive the health checks.

| Command | Purpose | Docs |
|---|---|---|
| `Get-ADForest` | Enumerates all domains in the forest to build the DC list | [Microsoft Docs](https://learn.microsoft.com/en-us/powershell/module/activedirectory/get-adforest) |
| `Get-ADDomainController` | Discovers every Domain Controller per domain | [Microsoft Docs](https://learn.microsoft.com/en-us/powershell/module/activedirectory/get-addomaincontroller) |
| `Test-Connection` | Probes ICMP reachability for each DC | [Microsoft Docs](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.management/test-connection) |
| `Get-CimInstance` | Reads uptime, OS disk space, and time sync via remote WMI | [Microsoft Docs](https://learn.microsoft.com/en-us/powershell/module/cimcmdlets/get-ciminstance) |
| `Send-MailMessage` | Delivers the inline HTML report via SMTP | [Microsoft Docs](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/send-mailmessage) |

---

# 🛡 Operational Notes
* Output: `ADHealthReport_yyyyMMdd_HHmmss.html` in `ReportPath`.
* Tunable thresholds in the script (uptime warn 24h, disk free 5/30%, time drift 0.5/1.0s).
* Never hard-code credentials — prefer `Get-Credential`, Windows Credential Manager, or SecretManagement.
* The script runs cleanly under Task Scheduler as a service account with AD/WMI permissions.
* `Send-MailMessage` is deprecated upstream; a future revision may migrate to Graph SendMail.

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