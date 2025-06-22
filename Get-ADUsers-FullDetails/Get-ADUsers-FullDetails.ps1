<#
.SYNOPSIS
    Get all AD user details including last logon, inactivity days, creation date, modification date, and last password set.

.DESCRIPTION
    This script retrieves all users in the specified OU from on-prem Active Directory, and includes:
    - Name
    - SamAccountName
    - Enabled
    - LastLogonDate
    - DaysInactive
    - Created (whenCreated)
    - Modified (whenChanged)
    - PasswordLastSet

.EXAMPLE
    .\Get-ADUsers-FullDetails.ps1

#>

# Load AD module
Import-Module ActiveDirectory

# Set the OU
$OU = "OU=Operation Dept,DC=QassimU,DC=local"

# Get today’s date
$Today = Get-Date

# Fetch users and enrich with additional properties
$Users = Get-ADUser -SearchBase $OU -Filter * -Properties DisplayName, LastLogonDate, Enabled, SamAccountName, whenCreated, whenChanged, PasswordLastSet |
    Select-Object `
        Name,
        SamAccountName,
        Enabled,
        whenCreated,
        whenChanged,
        LastLogonDate,
        @{Name = "DaysInactive"; Expression = {
            if ($_.LastLogonDate) {
                ($Today - $_.LastLogonDate).Days
            } else {
                "Never Logged In"
            }
        }},
        PasswordLastSet

# Display results
$Users | Sort-Object DaysInactive -Descending | Format-Table -AutoSize

# Export to CSV
$ExportPath = "$env:USERPROFILE\Desktop\OperationDept_UserDetails.csv"
$Users | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8

Write-Host "`n✅ Report saved to: $ExportPath" -ForegroundColor Green
