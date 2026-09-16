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
    1.2.0

.CHANGELOG
    1.2.0 (2026-09-16)
    - Added canonical CLI logging, $ErrorActionPreference = 'Stop', Write-Progress, and exit codes
    - Log path now anchors in %ProgramData% (Law 12)

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
    - Log file: %ProgramData%\ShadowGroup\Logs\ShadowGroup_<timestamp>.log
    - Dry-run: set $Update = $false to preview changes without touching AD.
    - Exit code 0 on completion; 1 on fatal error.
#>

#Requires -Version 5.1

$SolutionName = 'ShadowGroup'
$ScriptMode   = 'run'

$ErrorActionPreference = 'Stop'
# --- Logging (CLI Configuration) --------------------------------------------
$script:SystemDrive = if ($env:SystemDrive) { $env:SystemDrive.TrimEnd('\') } else {
    [System.IO.Path]::GetPathRoot($env:SystemRoot).TrimEnd('\')
}
$script:LogRoot  = $null
$script:LogFile  = $null
$script:LogReady = $false

function Initialize-Log {
    [CmdletBinding()]
    param(
        [string]$SolutionName = 'EnterpriseAdminTool',
        [string]$ScriptMode = 'run',
        [ValidateSet('Intune', 'General')]
        [string]$Type = 'General'
    )

    try {
        if ($Type -eq 'Intune') {
            $script:LogRoot = Join-Path $script:SystemDrive "IntuneLogs\$SolutionName"
            $script:LogFile = Join-Path $script:LogRoot "$SolutionName-$ScriptMode.txt"
        } else {
            $script:LogRoot = Join-Path $env:ProgramData "$SolutionName\Logs"
            $script:LogFile = Join-Path $script:LogRoot "$SolutionName`_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
        }

        if (-not (Test-Path -LiteralPath $script:LogRoot)) {
            $null = [System.IO.Directory]::CreateDirectory($script:LogRoot)
        }
        if (-not (Test-Path -LiteralPath $script:LogFile)) {
            $null = [System.IO.File]::Create($script:LogFile).Dispose()
        }

        $script:LogReady = $true
        return $true
    }
    catch {
        Write-Host "Log initialization failed: $($_.Exception.Message)" -ForegroundColor Red
        $script:LogReady = $false
        return $false
    }
}

function Write-Banner {
    [CmdletBinding()]
    [Alias('Show-Banner')]
    param()

    $title      = '{0} | {1} | {2}' -f $SolutionName, $ScriptMode, (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    $bannerLine = '=' * 78
    $lines      = @('', $bannerLine, $title, $bannerLine)

    foreach ($line in $lines) {
        if ($line -eq $title) {
            Write-Host $line -ForegroundColor White
        } else {
            Write-Host $line -ForegroundColor DarkGray
        }

        if ($script:LogReady -and $script:LogFile) {
            Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue -WhatIf:$false
        }
    }
}

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$Message = "",
        [ValidateSet("INFO", "SUCCESS", "WARNING", "ERROR", "DEBUG")]
        [string]$Level = "INFO"
    )

    if ([string]::IsNullOrEmpty($Message)) { return }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $fileLine  = "[$timestamp] [$Level] $Message"

    $color = switch ($Level) {
        "DEBUG"   { "DarkGray" }
        "INFO"    { "Cyan" }
        "SUCCESS" { "Green" }
        "WARNING" { "Yellow" }
        "ERROR"   { "Red" }
    }
    Write-Host $Message -ForegroundColor $color

    if ($script:LogReady -and $script:LogFile) {
        Add-Content -LiteralPath $script:LogFile -Value $fileLine -Encoding UTF8 -ErrorAction SilentlyContinue -WhatIf:$false
    }
}

function Finish-Script {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int]$ExitCode,
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [ValidateSet("INFO", "SUCCESS", "WARNING", "ERROR", "DEBUG")]
        [string]$Level = "INFO",
        [switch]$NoExit
    )

    Write-Log -Message $Message -Level $Level
    if (-not $NoExit) {
        exit $ExitCode
    }
}

$null = Initialize-Log -SolutionName $SolutionName -ScriptMode $ScriptMode -Type 'General'
Write-Banner

###############################################################################
# PARAMETERS
###############################################################################

# Domain Controller that supports Active Directory module
$Server = "dc01.example.com"

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
    Write-Log -Message "Active Directory module not found. Script aborted." -Level "ERROR"
    Finish-Script -ExitCode 1 -Message "Active Directory module not found." -Level "ERROR"
}

# Verify Domain Controller connectivity
if (-not (Test-Connection -ComputerName $Server -Count 1 -Quiet)) {
    Write-Log -Message "Unable to connect to $Server. Script aborted." -Level "ERROR"
    Finish-Script -ExitCode 1 -Message "Unable to connect to $Server." -Level "ERROR"
}

# Ensure valid OUs
foreach ($OUDN in $OUDNs) {
    try {
        [ADSI]"LDAP://$OUDN" | Out-Null
    } catch {
        Write-Log -Message "Invalid OU $OUDN. Script aborted." -Level "ERROR"
        Finish-Script -ExitCode 1 -Message "Invalid OU $OUDN." -Level "ERROR"
    }
}

# Validate shadow group existence
try {
    [ADSI]"LDAP://$GroupDN" | Out-Null
} catch {
    Write-Log -Message "Invalid group $GroupDN. Script aborted." -Level "ERROR"
    Finish-Script -ExitCode 1 -Message "Invalid group $GroupDN." -Level "ERROR"
}

# Determine LDAP search scope
$Scope = if ($ChildOUs) { "SubTree" } else { "OneLevel" }

# Retrieve current group members
$Members = Get-ADUser -LDAPFilter "(memberOf=$GroupDN)" -Server $Server | Select distinguishedName, Enabled

Write-Log -Message "Found $($Members.Count) current member(s) in the shadow group." -Level "INFO"

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

$processedCount = 0
$totalOUs = @($OUDNs).Count

foreach ($OUDN in $OUDNs) {
    $processedCount++
    # Activity = persistent label; Status = current OU; PercentComplete = 0..100
    Write-Progress -Activity 'Collecting OU users' `
        -Status "OU $processedCount of $totalOUs : $OUDN" `
        -PercentComplete (($processedCount / [Math]::Max($totalOUs, 1)) * 100)

    $UsersInOU = Get-ADUser -SearchBase $OUDN -SearchScope $Scope -LDAPFilter $Filter -Server $Server
    foreach ($User in $UsersInOU) {
        $UsersToAdd += $User.distinguishedName
    }
}

# Always close the progress bar when the loop exits (success OR failure)
Write-Progress -Activity 'Collecting OU users' -Completed

# Apply the per-run change cap
if (($UsersToAdd.Count + $UsersToRemove.Count) -gt $MaxChanges) {
    Write-Log -Message "Change count ($($UsersToAdd.Count) add + $($UsersToRemove.Count) remove) exceeds the max of $MaxChanges. Aborting before applying changes." -Level "ERROR"
    Finish-Script -ExitCode 1 -Message "Change cap exceeded. No changes applied." -Level "ERROR"
}

# Apply changes if enabled
if ($Update) {
    if ($UsersToRemove.Count -gt 0) {
        Remove-ADGroupMember -Identity $GroupDN -Members $UsersToRemove -Server $Server -Confirm:$False
        Write-Log -Message "Removed $($UsersToRemove.Count) user(s) from the shadow group." -Level "SUCCESS"
    }
    if ($UsersToAdd.Count -gt 0) {
        Add-ADGroupMember -Identity $GroupDN -Members $UsersToAdd -Server $Server
        Write-Log -Message "Added $($UsersToAdd.Count) user(s) to the shadow group." -Level "SUCCESS"
    }
    if ($UsersToAdd.Count -eq 0 -and $UsersToRemove.Count -eq 0) {
        Write-Log -Message "Shadow group is already in sync. No changes required." -Level "INFO"
    }
} else {
    Write-Log -Message "Dry-run: $($UsersToRemove.Count) to remove, $($UsersToAdd.Count) to add (no changes applied)." -Level "WARNING"
}

# Log completion
Write-Log -Message "Users Removed: $($UsersToRemove.Count)" -Level "INFO"
Write-Log -Message "Users Added: $($UsersToAdd.Count)" -Level "INFO"

Finish-Script -ExitCode 0 -Message "Execution completed. See log: $script:LogFile" -Level "SUCCESS"