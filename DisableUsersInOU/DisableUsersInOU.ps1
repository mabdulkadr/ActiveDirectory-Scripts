<#
.TITLE
    DisableUsersInOU - Disable all enabled user accounts in a target OU

.SYNOPSIS
    Disables every enabled user account in a specified Active Directory Organizational Unit.

.DESCRIPTION
    Searches a configured OU for all enabled user accounts and disables them one by one.
    Each action is written to a formatted execution log; a colored summary shows the total
    users found, successfully disabled, and failed. Set $ExportCSV to also emit a CSV report
    of every attempted action. The operation is irreversible for the affected logins.

.TAGS
    ActiveDirectory, User Management

.PLATFORM
    Windows

.PERMISSIONS
    Write access to user objects in the target OU (Disable-ADAccount)

.AUTHOR
    Mohammad Abdelkader Omar

.VERSION
    1.3.0

.CHANGELOG
    1.3.0 (2026-09-16)
    - Added canonical CLI logging, $ErrorActionPreference = 'Stop', Write-Progress, and exit codes
    - Report path now anchors beside the script (Law 12)

    1.2.0 (2026-09-16)
    - Standardized help header and script formatting

    1.1.0 (2025-06-25)
    - Initial release

.LASTUPDATE
    2026-09-16

.EXAMPLE
    .\DisableUsersInOU.ps1
    Disables all enabled users in the OU defined by $OU. Requires the ActiveDirectory
    module and permissions to disable accounts in that OU.

.NOTES
    - Requires RSAT (ActiveDirectory module) on a domain-joined host.
    - Configure $OU (and optionally $ExportCSV) in the Configuration section first.
    - Log file: %ProgramData%\DisableUsersInOU\Logs\
    - CSV report (when enabled): <script-folder>\Reports\
    - Exit code 0 on success; 1 on fatal error.
#>

#Requires -Version 5.1

$SolutionName = 'DisableUsersInOU'
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

$OU = "OU=IT,DC=ABC,DC=local"
$ExportCSV = $true

# Report path law: seed output beside the script, never from Get-Location alone.
$scriptBase = if ($PSScriptRoot) { $PSScriptRoot }
elseif ($PSCommandPath) { Split-Path -Parent $PSCommandPath }
elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path }
else { (Get-Location).Path }
$reports = Join-Path $scriptBase "Reports"
if (-not (Test-Path -LiteralPath $reports)) {
    $null = [System.IO.Directory]::CreateDirectory($reports)
}

#====================[ Main Execution ]====================#

try {
    Write-Log -Message "Starting user disable process in OU: $OU" -Level "INFO"
    Import-Module ActiveDirectory -ErrorAction Stop

    $UsersToDisable = Get-ADUser -Filter 'Enabled -eq $true' -SearchBase $OU -Properties SamAccountName, DistinguishedName

    if ($UsersToDisable.Count -eq 0) {
        Write-Log -Message "No enabled users found in the specified OU." -Level "WARNING"
        Finish-Script -ExitCode 0 -Message "No enabled users found in OU: $OU." -Level "SUCCESS"
    } else {
        Write-Log -Message "$($UsersToDisable.Count) enabled user(s) found. Proceeding to disable them..." -Level "INFO"

        $processedCount = 0
        $total = @($UsersToDisable).Count

        $DisabledResults = foreach ($user in $UsersToDisable) {
            $processedCount++
            # Activity = persistent label; Status = current item; PercentComplete = 0..100
            Write-Progress -Activity 'Disabling users' `
                -Status "User $processedCount of $total : $($user.SamAccountName)" `
                -PercentComplete (($processedCount / [Math]::Max($total, 1)) * 100)

            try {
                Disable-ADAccount -Identity $user.DistinguishedName -ErrorAction Stop
                Write-Log -Message "Disabled user: $($user.SamAccountName)" -Level "SUCCESS"
                [PSCustomObject]@{
                    SamAccountName     = $user.SamAccountName
                    DistinguishedName  = $user.DistinguishedName
                    Status             = "Disabled"
                    Timestamp          = Get-Date
                }
            } catch {
                Write-Log -Message "Failed to disable user: $($user.SamAccountName). Error: $($_.Exception.Message)" -Level "ERROR"
                [PSCustomObject]@{
                    SamAccountName     = $user.SamAccountName
                    DistinguishedName  = $user.DistinguishedName
                    Status             = "Failed"
                    Timestamp          = Get-Date
                }
            }
        }

        # Always close the progress bar when the loop exits (success OR failure)
        Write-Progress -Activity 'Disabling users' -Completed

        # Export results to CSV
        $CsvFile = $null
        if ($ExportCSV) {
            $CsvFile = Join-Path $reports "DisabledUsers_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
            $DisabledResults | Export-Csv -Path $CsvFile -NoTypeInformation -Encoding UTF8
            Write-Log -Message "Exported results to: $CsvFile" -Level "SUCCESS"
        }

        #====================[ Summary Section ]====================#
        $Total   = $DisabledResults.Count
        $Success = ($DisabledResults | Where-Object { $_.Status -eq 'Disabled' }).Count
        $Failed  = ($DisabledResults | Where-Object { $_.Status -eq 'Failed' }).Count

        Write-Log -Message "Total: $Total | Disabled: $Success | Failed: $Failed" -Level "INFO"

        if ($CsvFile) {
            Write-Log -Message "CSV Report Saved To: $CsvFile" -Level "INFO"
        }

        if ($Failed -gt 0) {
            Finish-Script -ExitCode 1 -Message "Completed with $Failed failure(s)." -Level "ERROR"
        } else {
            Finish-Script -ExitCode 0 -Message "Completed - $Success user(s) disabled." -Level "SUCCESS"
        }
    }
} catch {
    Write-Log -Message "Fatal error: $($_.Exception.Message)" -Level "ERROR"
    Finish-Script -ExitCode 1 -Message "Fatal error: $($_.Exception.Message)" -Level "ERROR"
}