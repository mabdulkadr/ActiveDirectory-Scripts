<#
.TITLE
    Get-ADUsers-Details - AD user inventory with activity metrics for an OU

.SYNOPSIS
    Exports detailed AD user attributes from an OU, sorted by inactivity, to CSV.

.DESCRIPTION
    Retrieves all users in a configured Organizational Unit and reports display and account
    info, password settings and lockout status, group memberships, OU location, and last
    activity. Results are sorted by inactivity (most inactive first), shown in the console,
    and exported to a timestamped CSV in the Reports folder beside the script. Read-only.

.TAGS
    ActiveDirectory, Reporting

.PLATFORM
    Windows

.PERMISSIONS
    Read access to the target OU (ActiveDirectory module)

.AUTHOR
    Mohammad Abdelkader Omar

.VERSION
    2.4.0

.CHANGELOG
    2.4.0 (2026-09-16)
    - Added canonical CLI logging, $ErrorActionPreference = 'Stop', and exit codes
    - Report path now anchors beside the script (Law 12)

    2.3.0 (2026-09-16)
    - Standardized help header and section formatting

    2.2.0 (2025-06-25)
    - Initial release

.LASTUPDATE
    2026-09-16

.EXAMPLE
    .\Get-ADUsers-Details.ps1
    Inventories all users in the OU defined by $OU and writes the CSV report to
    <script-folder>\Reports\ADUsers-Details_<timestamp>.csv.

.NOTES
    - Requires RSAT (ActiveDirectory module) and read access to the target OU.
    - Configure the $OU variable in the Configuration section before running.
    - Output: <script-folder>\Reports\ADUsers-Details_<timestamp>.csv
    - Read-only: no directory objects are modified.
    - Exit code 0 on completion; 1 on fatal error.
#>

#Requires -Version 5.1

$SolutionName = 'GetADUsersDetails'
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

#====================[ Configuration ]====================#
$OU = "OU=Employees,DC=abc,DC=local"

# Report path law: seed output beside the script, never from Get-Location alone.
$scriptBase = if ($PSScriptRoot) { $PSScriptRoot }
elseif ($PSCommandPath) { Split-Path -Parent $PSCommandPath }
elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path }
else { (Get-Location).Path }
$reports = Join-Path $scriptBase "Reports"
if (-not (Test-Path -LiteralPath $reports)) {
    $null = [System.IO.Directory]::CreateDirectory($reports)
}

$Today = Get-Date
$TimeStamp = $Today.ToString("yyyyMMdd-HHmmss")
$ExportPath = Join-Path $reports "ADUsers-Details-$TimeStamp.csv"

#====================[ Load AD Module ]====================#
try {
    Import-Module ActiveDirectory -ErrorAction Stop
} catch {
    Write-Log -Message "ActiveDirectory module is required. Please install RSAT and try again." -Level "ERROR"
    Finish-Script -ExitCode 1 -Message "ActiveDirectory module is missing." -Level "ERROR"
}

#====================[ Collect Users ]====================#
$Users = Get-ADUser -SearchBase $OU -Filter * -Properties DisplayName, LastLogonDate, Enabled, SamAccountName, whenCreated, whenChanged, PasswordLastSet, MemberOf, PasswordNeverExpires, LockedOut, CannotChangePassword, UserPrincipalName, Title, Department

$processedCount = 0
$total = @($Users).Count

# Builds one result object per user with enriched account details.
$Results = foreach ($User in $Users) {
    $processedCount++
    # Activity = persistent label; Status = current item; PercentComplete = 0..100
    Write-Progress -Activity 'Collecting AD user details' `
        -Status "User $processedCount of $total : $($User.SamAccountName)" `
        -PercentComplete (($processedCount / [Math]::Max($total, 1)) * 100)

    # Get OU from DN
    $UserOU = ($User.DistinguishedName -split ",(?=OU=)")[1..99] -join ","

    # Resolve group names
    $GroupNames = foreach ($dn in $User.MemberOf) {
        try {
            (Get-ADGroup $dn -ErrorAction Stop).Name
        } catch {
            "[Unknown Group]"
        }
    }
    $GroupsJoined = $GroupNames -join ", "

    # Build result object
    [PSCustomObject]@{
        DisplayName          = $User.DisplayName
        SamAccountName       = $User.SamAccountName
        UserPrincipalName    = $User.UserPrincipalName
        Title                = $User.Title
        Department           = $User.Department
        Enabled              = $User.Enabled
        OU                   = $UserOU
        Groups               = $GroupsJoined
        Created              = $User.whenCreated
        Modified             = $User.whenChanged
        LastLogonDate        = $User.LastLogonDate
        DaysInactive         = if ($User.LastLogonDate) { ($Today - $User.LastLogonDate).Days } else { "Never Logged In" }
        PasswordLastSet      = $User.PasswordLastSet
        PasswordNeverExpires = $User.PasswordNeverExpires
        LockedOut            = $User.LockedOut
        CannotChangePassword = $User.CannotChangePassword
    }
}

# Always close the progress bar when the loop exits (success OR failure)
Write-Progress -Activity 'Collecting AD user details' -Completed

#====================[ Console Output ]====================#
# Show users sorted by inactivity, most inactive first.
$Results | Sort-Object DaysInactive -Descending | Format-Table -AutoSize

#====================[ Export to CSV ]====================#
try {
    $Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Log -Message "Report saved to: $ExportPath" -Level "SUCCESS"
} catch {
    Write-Log -Message "Failed to export CSV: $_" -Level "ERROR"
    Finish-Script -ExitCode 1 -Message "Failed to export CSV: $_" -Level "ERROR"
}

Finish-Script -ExitCode 0 -Message "Report generated: $ExportPath" -Level "SUCCESS"