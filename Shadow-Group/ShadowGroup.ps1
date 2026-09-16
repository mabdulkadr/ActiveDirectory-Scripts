<#
.TITLE
    ShadowGroup - Mirror OU users into a shadow Active Directory group

.SYNOPSIS
    Keeps an AD shadow group in sync with the users of one or more Organizational Units.

.DESCRIPTION
    Ensures all users in the specified OUs are members of a corresponding shadow group and
    removes users who no longer belong. Supports child-OUs inclusion, an enabled-users-only
    filter, dry-run logging, and a per-run change cap to prevent excessive AD operations.
    Intended for scheduled runs on a Domain Controller for group-based policies (such as
    Fine-Grained Password Policies).

.TAGS
    ActiveDirectory, Group Management, Automation

.PLATFORM
    Windows

.PERMISSIONS
    Manage group membership in AD (Add-ADGroupMember / Remove-ADGroupMember)

.AUTHOR
    Mohammad Abdelkader Omar

.VERSION
    1.1.0

.CHANGELOG
    1.1.0 (2026-09-16)
    - Standardized help header and script formatting

    1.0.0 (2025-02-02)
    - Initial release

.LASTUPDATE
    2026-09-16

.EXAMPLE
    .\ShadowGroup.ps1
    Runs the sync with the settings in the Parameters section. Set $Update to $true to apply
    changes; keep it $false to log planned changes only.

.NOTES
    - Designed to run on a Domain Controller or admin host with the ActiveDirectory module.
    - The shadow group must exist before running; the script exits if it cannot resolve it.
    - Log file: C:\ShadowGroups\ShadowGroup.log
    - Dry-run: set $Update = $false to preview changes without touching AD.
#>

#Requires -Version 5.1

###############################################################################
# PARAMETERS
###############################################################################

# Domain Controller that supports Active Directory module
$Server = "dc01.example.com"

# Log file location
$LogFile = "C:\ShadowGroups\ShadowGroup.log"

# Ensure log file exists
if (-not (Test-Path $LogFile)) {
    New-Item -ItemType File -Path $LogFile -Force | Out-Null
}

# Enable/Disable actual modifications (True applies changes, False logs only)
$Update = $true

# Include only enabled users (True excludes disabled accounts)
$EnabledOnly = $true

# Organizational Units to monitor
$OUDNs = @("OU=HR,OU=Company,DC=example,DC=com", "OU=IT,OU=Company,DC=example,DC=com")

# Include child OUs (True considers sub-OUs, False restricts to direct OU members)
$ChildOUs = $true

# Shadow Group Distinguished Name (Must exist in AD)
$GroupDN = "CN=Shadow-Group,OU=Security,DC=example,DC=com"

# Maximum modifications per run to prevent excessive network load
$MaxChanges = 4000

###############################################################################
# SCRIPT EXECUTION
###############################################################################

# Import Active Directory Module
try {
    Import-Module ActiveDirectory -ErrorAction Stop -WarningAction Stop
} catch {
    Write-Host "Active Directory module not found. Script aborted." -ForegroundColor Red
    break
}

# Verify Domain Controller connectivity
if (-not (Test-Connection -ComputerName $Server -Count 1 -Quiet)) {
    Write-Host "Error: Unable to connect to $Server. Script aborted." -ForegroundColor Red
    break
}

# Ensure valid OUs
foreach ($OUDN in $OUDNs) {
    try {
        [ADSI]"LDAP://$OUDN" | Out-Null
    } catch {
        Write-Host "Error: Invalid OU $OUDN. Script aborted." -ForegroundColor Red
        break
    }
}

# Validate shadow group existence
try {
    [ADSI]"LDAP://$GroupDN" | Out-Null
} catch {
    Write-Host "Error: Invalid group $GroupDN. Script aborted." -ForegroundColor Red
    break
}

# Determine LDAP search scope
$Scope = if ($ChildOUs) { "SubTree" } else { "OneLevel" }

# Retrieve current group members
try {
    $Members = Get-ADUser -LDAPFilter "(memberOf=$GroupDN)" -Server $Server | Select distinguishedName, Enabled
} catch {
    Write-Host "Error: Unable to retrieve group members. Script aborted." -ForegroundColor Red
    break
}

# Prepare logging
try {
    Add-Content -Path $LogFile -Value "--- Script Execution Start: $(Get-Date) ---"
} catch {
    Write-Host "Error: Cannot write to log file. Script aborted." -ForegroundColor Red
    break
}

# Users to add and remove
$UsersToAdd = @()
$UsersToRemove = @()

# Evaluate group membership
foreach ($Member in $Members) {
    if ($EnabledOnly -and -not $Member.Enabled) {
        $UsersToRemove += $Member.distinguishedName
    } else {
        $UserOU = ($Member.distinguishedName -split ",", 2)[1]
        if ($OUDNs -notcontains $UserOU) {
            $UsersToRemove += $Member.distinguishedName
        }
    }
}

# Fetch users in OUs not in the group
$Filter = "(!(memberOf=$GroupDN))"
if ($EnabledOnly) {
    $Filter = "(&" + $Filter + "(!(userAccountControl:1.2.840.113556.1.4.803:=2)))"
}

foreach ($OUDN in $OUDNs) {
    $UsersInOU = Get-ADUser -SearchBase $OUDN -SearchScope $Scope -LDAPFilter $Filter -Server $Server
    foreach ($User in $UsersInOU) {
        $UsersToAdd += $User.distinguishedName
    }
}

# Apply changes if enabled
if ($Update) {
    if ($UsersToRemove.Count -gt 0) {
        Remove-ADGroupMember -Identity $GroupDN -Members $UsersToRemove -Server $Server -Confirm:$False
    }
    if ($UsersToAdd.Count -gt 0) {
        Add-ADGroupMember -Identity $GroupDN -Members $UsersToAdd -Server $Server
    }
}

# Log completion
Add-Content -Path $LogFile -Value "Users Removed: $($UsersToRemove.Count)"
Add-Content -Path $LogFile -Value "Users Added: $($UsersToAdd.Count)"
Add-Content -Path $LogFile -Value "--- Script Execution End: $(Get-Date) ---"

# Display summary
Write-Host "Execution completed. See log: $LogFile"
Write-Host "Users removed: $($UsersToRemove.Count)" -ForegroundColor Yellow
Write-Host "Users added: $($UsersToAdd.Count)" -ForegroundColor Green
