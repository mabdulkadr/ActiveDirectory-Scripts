<#
.TITLE
    Add-UsersFromCSVToADGroup - Bulk-add users from CSV to an Active Directory group

.SYNOPSIS
    Adds user IDs from a CSV file to an Active Directory group, skipping existing members.

.DESCRIPTION
    Reads user IDs from a CSV file, appends the configured UPN domain suffix, verifies
    each user exists in Active Directory, checks the user is not already a group member,
    and adds only the missing users to the specified group. Prints a color-coded summary
    of added, already-existing, and failed users. The operation is additive only and never
    removes group members.

.TAGS
    ActiveDirectory, Bulk Operations

.PLATFORM
    Windows

.PERMISSIONS
    Write access to the target Active Directory group (Add-ADGroupMember)

.AUTHOR
    Mohammad Abdelkader Omar

.VERSION
    1.1.0

.CHANGELOG
    1.1.0 (2026-09-16)
    - Added canonical CLI logging, $ErrorActionPreference = 'Stop', Write-Progress, and exit codes

    1.0.0 (2026-09-16)
    - Initial release

.LASTUPDATE
    2026-09-16

.EXAMPLE
    .\Add-UsersFromCSVToADGroup.ps1
    Adds all users from the configured CSV to the target group. The group name, CSV path,
    and UPN domain suffix are set in the Configuration section at the top of the script.

.NOTES
    - Requires RSAT (ActiveDirectory module) and a domain-joined host.
    - The CSV must contain a "UserID" column; any other column name yields zero matches.
    - Configure $GroupName, $CsvPath, and $Domain before the first run.
    - Additive only: users already in the group are skipped, never removed.
    - Exit code 0 on completion; 1 on fatal error.
#>

#Requires -Version 5.1

$SolutionName = 'Add-UsersFromCSVToADGroup'
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

$GroupName = "Blind-UserLogin-Exception"
$CsvPath = "C:\Add-UsersFromCSVToGroup\users.csv"
$Domain = "@abc.local"

$SuccessCount = 0
$FailureCount = 0
$AlreadyCount = 0

#====================[ Display Group Details ]====================#

try {
    $Group = Get-ADGroup -Identity $GroupName -Properties Description
    Write-Log -Message "Target Group         : $($Group.Name)" -Level "INFO"
    Write-Log -Message "Distinguished Name   : $($Group.DistinguishedName)" -Level "INFO"
    if ($Group.Description) {
        Write-Log -Message "Description          : $($Group.Description)" -Level "INFO"
    }
    Write-Log -Message "" -Level "INFO"
}
catch {
    Write-Log -Message "Failed to retrieve group details: $_" -Level "ERROR"
    Finish-Script -ExitCode 1 -Message "Failed to retrieve group details: $_" -Level "ERROR"
}

#====================[ Import and Process Users ]====================#

try {
    $UserIDs = Import-Csv -Path $CsvPath | Select-Object -ExpandProperty UserID
}
catch {
    Write-Log -Message "Failed to read CSV file: $_" -Level "ERROR"
    Finish-Script -ExitCode 1 -Message "Failed to read CSV file: $_" -Level "ERROR"
}

$processedCount = 0
$total = @($UserIDs).Count

foreach ($id in $UserIDs) {
    $processedCount++
    # Activity = persistent label; Status = current item; PercentComplete = 0..100
    Write-Progress -Activity 'Processing users from CSV' `
        -Status "User $processedCount of $total : $id" `
        -PercentComplete (($processedCount / [Math]::Max($total, 1)) * 100)

    $UPN = "$id$Domain"

    try {
        $User = Get-ADUser -Filter "UserPrincipalName -eq '$UPN'" -ErrorAction Stop

        # Check if user is already a member
        $isMember = Get-ADGroupMember -Identity $GroupName -Recursive | Where-Object {
            $_.DistinguishedName -eq $User.DistinguishedName
        }

        if ($isMember) {
            Write-Log -Message "Already a member: $UPN" -Level "WARNING"
            $AlreadyCount++
        }
        else {
            Add-ADGroupMember -Identity $GroupName -Members $User -ErrorAction Stop
            Write-Log -Message "Added: $UPN" -Level "SUCCESS"
            $SuccessCount++
        }
    }
    catch {
        Write-Log -Message "Failed to process $UPN : $_" -Level "ERROR"
        $FailureCount++
    }
}

# Always close the progress bar when the loop exits (success OR failure)
Write-Progress -Activity 'Processing users from CSV' -Completed

#====================[ Summary ]====================#

Write-Log -Message "Total added       : $SuccessCount" -Level "SUCCESS"
Write-Log -Message "Already in group  : $AlreadyCount" -Level "WARNING"
Write-Log -Message "Failed attempts   : $FailureCount" -Level "ERROR"

if ($FailureCount -gt 0) {
    Finish-Script -ExitCode 1 -Message "Script completed with $FailureCount error(s)." -Level "ERROR"
} else {
    Finish-Script -ExitCode 0 -Message "Script completed successfully." -Level "SUCCESS"
}