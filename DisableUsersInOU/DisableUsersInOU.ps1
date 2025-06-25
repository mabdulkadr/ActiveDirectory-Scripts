<#
.SYNOPSIS
    Disables all user accounts in a specific Active Directory Organizational Unit (OU).

.DESCRIPTION
    This script searches for all enabled user accounts within a given OU and disables them.
    It logs the actions taken and exports the affected user list to a CSV file for auditing.

.EXAMPLE
    Run the script directly after setting the correct $OU value inside the script.

.NOTES
    Author  : Mohammad Abdelkader
    Website : momar.tech
    Date    : 2025-06-25
    Version : 1.1
#>

# =============================
# Configuration
# =============================

$OU = "OU=External Users,OU=Resaerch Project,OU=QU,DC=ABC,DC=local"
$ExportCSV = $true

# =============================
# Logging Setup
# =============================

$LogPath = "C:\Logs"
if (-not (Test-Path $LogPath)) { New-Item -Path $LogPath -ItemType Directory -Force }
$LogFile = Join-Path $LogPath "DisableUsers_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"

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

# =============================
# Main Execution
# =============================

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

        # =============================
        # Summary Section
        # =============================
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
