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
    1.2.0

.CHANGELOG
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
    - Log file: C:\Logs\DisableUsers_<timestamp>.log
    - CSV report (when enabled): C:\Reports\DisabledUsers_<timestamp>.csv
    - Exit code 0 on success; 1 on fatal error.
#>

#Requires -Version 5.1

#====================[ Configuration ]====================#

$OU = "OU=IT,DC=ABC,DC=local"
$ExportCSV = $true

#====================[ Logging Setup ]====================#

$LogPath = "C:\Logs"
if (-not (Test-Path $LogPath)) { New-Item -Path $LogPath -ItemType Directory -Force }
$LogFile = Join-Path $LogPath "DisableUsers_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"

# Writes a timestamped entry to the console and the execution log file.
function Write-Log {
    param (
        [string]$Message,
        [ValidateSet("INFO", "SUCCESS", "ERROR", "WARNING")]
        [string]$Level = "INFO"
    )
    $Color = switch ($Level) {
        "INFO"    { "Cyan" }
        "SUCCESS" { "Green" }
        "WARNING" { "Yellow" }
        "ERROR"   { "Red" }
    }
    Write-Host "[$Level] $Message" -ForegroundColor $Color
    Add-Content -Path $LogFile -Value "[$Level] $(Get-Date -Format 'u') - $Message"
}

#====================[ Main Execution ]====================#

try {
    Write-Log "Starting user disable process in OU: $OU" -Level "INFO"
    Import-Module ActiveDirectory -ErrorAction Stop

    $UsersToDisable = Get-ADUser -Filter 'Enabled -eq $true' -SearchBase $OU -Properties SamAccountName, DistinguishedName

    if ($UsersToDisable.Count -eq 0) {
        Write-Log "No enabled users found in the specified OU." -Level "WARNING"
    } else {
        Write-Log "$($UsersToDisable.Count) enabled user(s) found. Proceeding to disable them..." -Level "INFO"

        $DisabledResults = foreach ($user in $UsersToDisable) {
            try {
                Disable-ADAccount -Identity $user.DistinguishedName -ErrorAction Stop
                Write-Log "Disabled user: $($user.SamAccountName)" -Level "SUCCESS"
                [PSCustomObject]@{
                    SamAccountName     = $user.SamAccountName
                    DistinguishedName  = $user.DistinguishedName
                    Status             = "Disabled"
                    Timestamp          = Get-Date
                }
            } catch {
                Write-Log "Failed to disable user: $($user.SamAccountName). Error: $($_.Exception.Message)" -Level "ERROR"
                [PSCustomObject]@{
                    SamAccountName     = $user.SamAccountName
                    DistinguishedName  = $user.DistinguishedName
                    Status             = "Failed"
                    Timestamp          = Get-Date
                }
            }
        }

        # Export results to CSV
        $CsvFile = $null
        if ($ExportCSV) {
            $ExportPath = "C:\Reports"
            if (-not (Test-Path $ExportPath)) { New-Item -Path $ExportPath -ItemType Directory -Force }
            $CsvFile = Join-Path $ExportPath "DisabledUsers_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
            $DisabledResults | Export-Csv -Path $CsvFile -NoTypeInformation -Encoding UTF8
            Write-Log "Exported results to: $CsvFile" -Level "SUCCESS"
        }

        #====================[ Summary Section ]====================#
        $Total   = $DisabledResults.Count
        $Success = ($DisabledResults | Where-Object { $_.Status -eq 'Disabled' }).Count
        $Failed  = ($DisabledResults | Where-Object { $_.Status -eq 'Failed' }).Count

        Write-Host "`n======================= SUMMARY =======================" -ForegroundColor Yellow
        Write-Host "OU:                  $OU" -ForegroundColor Cyan
        Write-Host "Total Users Found:   $Total" -ForegroundColor Cyan
        Write-Host "Successfully Disabled: $Success" -ForegroundColor Green
        Write-Host "Failed to Disable:     $Failed" -ForegroundColor Red
        if ($CsvFile) {
            Write-Host "CSV Report Saved To:  $CsvFile" -ForegroundColor Cyan
        }
        Write-Host "=======================================================`n"
    }
} catch {
    Write-Log "Fatal error: $($_.Exception.Message)" -Level "ERROR"
    exit 1
}
