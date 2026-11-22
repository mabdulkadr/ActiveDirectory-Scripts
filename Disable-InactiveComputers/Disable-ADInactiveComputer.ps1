<#
.TITLE
    Disable-ADInactiveComputer - Disable or delete stale computer accounts with HTML audit

.SYNOPSIS
    Finds inactive computer accounts in an OU by last logon and optionally disables or deletes them.

.DESCRIPTION
    Searches a configurable Organizational Unit for computer accounts that have not logged
    on within a threshold number of days, lists them in the console, prompts whether to disable
    them, then prompts whether to delete the disabled accounts. Every run produces a timestamped
    HTML report with color-coded tables for inactive, disabled, and deleted computers. Write actions
    are confirmed interactively and can be safety-skipped.

.TAGS
    ActiveDirectory, Computer Management, Cleanup

.PLATFORM
    Windows

.PERMISSIONS
    Full control over computer objects in the target OU (disable + delete)

.AUTHOR
    Mohammad Abdelkader Omar

.VERSION
    1.2.0

.CHANGELOG
    1.2.0 (2026-09-16)
    - Added canonical CLI logging, $ErrorActionPreference = 'Stop', Write-Progress, and exit codes
    - Report path now anchors beside the script (Law 12)

    1.1.0 (2026-09-16)
    - Standardized help header and script formatting

    1.0.0
    - Initial release

.LASTUPDATE
    2026-09-16

.EXAMPLE
    .\Disable-ADInactiveComputer.ps1 -DaysInactive 180 -SearchBaseOU "OU=Domain Computers,DC=ABC,DC=local"
    Flags any computer that has not logged on for more than 180 days and offers to disable
    and then delete those accounts, writing a timestamped audit report to
    <script-folder>\Reports\Computers_Report_<timestamp>.html.

.NOTES
    - Requires RSAT (ActiveDirectory module) and permissions to disable/delete computer objects.
    - HTML report: <script-folder>\Reports\Computers_Report_<timestamp>.html
    - Log file: %ProgramData%\DisableInactiveComputers\Logs\
    - Test against a staging OU before running in production; deletion is irreversible.
    - Exit code 0 on completion; 1 on fatal error.
#>

#Requires -Version 5.1

param (
    [int]$DaysInactive     = 180,                                            # Number of days of inactivity for filtering inactive computers
    [string]$SearchBaseOU  = "OU=Domain Computers,DC=ABC,DC=local"           # Organizational Unit (OU) for searching computers
)

$SolutionName = 'DisableInactiveComputers'
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

# Builds a single timestamped HTML report combining inactive, disabled, and deleted tables.
function Write-CombinedHTMLReport {
    param (
        [string]$HtmlPath,        # Path to save the final HTML report
        [PSObject]$InactiveData,  # Data for inactive computers
        [int]$TotalInactive,      # Total number of inactive computers
        [PSObject]$DisabledData,  # Data for disabled computers
        [int]$TotalDisabled,      # Total number of disabled computers
        [PSObject]$DeletedData,   # Data for deleted computers
        [int]$TotalDeleted        # Total number of deleted computers
    )

    # Basic HTML structure with inline CSS for formatting the report
    # This includes styling for tables, headers, and alternating row colors.
    $htmlHeader = @"
<!DOCTYPE html>
<html lang='en'>
<head>
    <meta charset='UTF-8'>
    <meta name='viewport' content='width=device-width, initial-scale=1.0'>
    <title>Inactive, Disabled, and Deleted Computers Report</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 20px;
        }
        h1, h2, h3 {
            color: #2c3e50;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            margin-top: 20px;
        }
        th, td {
            border: 1px solid #dddddd;
            text-align: left;
            padding: 8px;
        }
        th {
            background-color: #2c3e50;
            color: white;
        }
        tr:nth-child(even) {
            background-color: #f2f2f2;
        }
        .summary {
            background-color: #ecf0f1;
            padding: 10px;
            border-radius: 8px;
            margin-top: 20px;
        }
    </style>
</head>
<body>
<h1>Inactive, Disabled, and Deleted Computers Report</h1>
"@

    # Section for Inactive Computers
    $htmlInactiveSection = @"
<h2>Inactive Computers</h2>
<p class='summary'><strong>Total Inactive Computers:</strong> $TotalInactive</p>
<table>
<thead>
<tr>
    <th>Computer Name</th>
    <th>Last Logon Date</th>
    <th>Distinguished Name</th>
</tr>
</thead>
<tbody>
"@

    # Loop through the inactive computers and create HTML rows
    $htmlInactiveRows = foreach ($item in $InactiveData) {
        "<tr>
            <td>$($item.ComputerName)</td>
            <td>$($item.LastLogonDate)</td>
            <td>$($item.DistinguishedName)</td>
        </tr>"
    }

    $htmlInactiveFooter = "</tbody></table>"

    # Section for Disabled Computers
    $htmlDisabledSection = @"
<h2>Disabled Computers</h2>
<p class='summary'><strong>Total Disabled Computers:</strong> $TotalDisabled</p>
<table>
<thead>
<tr>
    <th>Computer Name</th>
    <th>Last Logon Date</th>
    <th>Distinguished Name</th>
</tr>
</thead>
<tbody>
"@

    # Loop through the disabled computers and create HTML rows
    $htmlDisabledRows = foreach ($item in $DisabledData) {
        "<tr>
            <td>$($item.ComputerName)</td>
            <td>$($item.LastLogonDate)</td>
            <td>$($item.DistinguishedName)</td>
        </tr>"
    }

    $htmlDisabledFooter = "</tbody></table>"

    # Section for Deleted Computers
    $htmlDeletedSection = @"
<h2>Deleted Computers</h2>
<p class='summary'><strong>Total Deleted Computers:</strong> $TotalDeleted</p>
<table>
<thead>
<tr>
    <th>Computer Name</th>
    <th>Last Logon Date</th>
    <th>Distinguished Name</th>
</tr>
</thead>
<tbody>
"@

    # Loop through the deleted computers and create HTML rows
    $htmlDeletedRows = foreach ($item in $DeletedData) {
        "<tr>
            <td>$($item.ComputerName)</td>
            <td>$($item.LastLogonDate)</td>
            <td>$($item.DistinguishedName)</td>
        </tr>"
    }

    $htmlDeletedFooter = "</tbody></table>"

    # Combine all sections (Inactive, Disabled, Deleted) and write to the HTML file
    $htmlFooter = "</body></html>"
    $fullHtml = $htmlHeader + $htmlInactiveSection + ($htmlInactiveRows -join "`n") + $htmlInactiveFooter + "`n" +
                $htmlDisabledSection + ($htmlDisabledRows -join "`n") + $htmlDisabledFooter + "`n" +
                $htmlDeletedSection + ($htmlDeletedRows -join "`n") + $htmlDeletedFooter + $htmlFooter

    # Save the full HTML report to the specified path
    Set-Content -Path $HtmlPath -Value $fullHtml
}

#===============[ Main Execution ]===============#

try {
    # Report path law: seed output beside the script, never from Get-Location alone.
    $scriptBase = if ($PSScriptRoot) { $PSScriptRoot }
    elseif ($PSCommandPath) { Split-Path -Parent $PSCommandPath }
    elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path }
    else { (Get-Location).Path }
    $reports = Join-Path $scriptBase "Reports"
    if (-not (Test-Path -LiteralPath $reports)) {
        $null = [System.IO.Directory]::CreateDirectory($reports)
    }

    # Generate a timestamp for the report filename
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $reportFileName = Join-Path $reports "Computers_Report_$timestamp.html"

    # Start processing inactive computers
    Write-Log -Message "Processing inactive computers based on LastLogonDate and $DaysInactive days of inactivity..." -Level "INFO"

    # Get the threshold date for inactivity (current date minus DaysInactive)
    $InactiveThreshold = (Get-Date).AddDays(-$DaysInactive)

    # Fetch all computers from the specified Organizational Unit (OU)
    $AllComputers = Get-ADComputer -Filter * -SearchBase $SearchBaseOU -Property Name, LastLogonDate -ErrorAction Stop

    # Filter inactive computers: Those with no LastLogonDate or last logon date older than the threshold
    $InactiveComputers = foreach ($Computer in $AllComputers) {
        $LastLogonDate = if ($Computer.LastLogonDate) { $Computer.LastLogonDate } else { "No Logon Information" }

        if ($LastLogonDate -eq "No Logon Information" -or [DateTime]$LastLogonDate -lt $InactiveThreshold) {
            [PSCustomObject]@{
                ComputerName      = $Computer.Name
                LastLogonDate     = $LastLogonDate
                DistinguishedName = $Computer.DistinguishedName
            }
        }
    }

    Write-Log -Message "Found $($InactiveComputers.Count) inactive computer(s)." -Level "INFO"

    # Initialize arrays for Disabled and Deleted computers
    $DisabledComputers = @()
    $DeletedComputers = @()

    # Check if there are any inactive computers to process
    if ($InactiveComputers.Count -gt 0) {
        # List inactive computers in the console
        Write-Log -Message "Inactive Computers List:" -Level "INFO"
        $InactiveComputers | Format-Table -Property ComputerName, LastLogonDate -AutoSize

        # Ask if you want to disable the inactive computers
        $DisableInput = Read-Host "`nDo you want to disable all the inactive computers? (Yes/No)"

        if ($DisableInput -eq "Yes") {
            # Loop through inactive computers and disable them
            $processedCount = 0
            $maxComputers = @($InactiveComputers).Count
            foreach ($Computer in $InactiveComputers) {
                $processedCount++
                # Activity = persistent label; Status = current item; PercentComplete = 0..100
                Write-Progress -Activity 'Disabling inactive computers' `
                    -Status "Computer $processedCount of $maxComputers : $($Computer.ComputerName)" `
                    -PercentComplete (($processedCount / [Math]::Max($maxComputers, 1)) * 100)

                try {
                    Write-Log -Message "Disabling computer: $($Computer.ComputerName)" -Level "WARNING"
                    Disable-ADAccount -Identity $Computer.DistinguishedName -ErrorAction Stop
                    Write-Log -Message "Disabled: $($Computer.ComputerName)" -Level "SUCCESS"

                    # Add the computer to the DisabledComputers array
                    $DisabledComputers += $Computer
                } catch {
                    Write-Log -Message "Failed to disable computer $($Computer.ComputerName): $($_.Exception.Message)" -Level "ERROR"
                }
            }
            # Always close the progress bar when the loop exits (success OR failure)
            Write-Progress -Activity 'Disabling inactive computers' -Completed
        } else {
            Write-Log -Message "No computers were disabled." -Level "WARNING"
        }

        # Ask if you want to delete the disabled computers
        $DeleteInput = Read-Host "`nDo you want to delete all disabled computers? (Yes/No)"

        if ($DeleteInput -eq "Yes" -and $DisabledComputers.Count -gt 0) {
            # Loop through disabled computers and delete them
            $processedCount = 0
            $maxComputers = @($DisabledComputers).Count
            foreach ($Computer in $DisabledComputers) {
                $processedCount++
                # Activity = persistent label; Status = current item; PercentComplete = 0..100
                Write-Progress -Activity 'Deleting disabled computers' `
                    -Status "Computer $processedCount of $maxComputers : $($Computer.ComputerName)" `
                    -PercentComplete (($processedCount / [Math]::Max($maxComputers, 1)) * 100)

                try {
                    Write-Log -Message "Deleting computer: $($Computer.ComputerName)" -Level "WARNING"
                    Remove-ADComputer -Identity $Computer.DistinguishedName -Confirm:$false -ErrorAction Stop
                    Write-Log -Message "Deleted: $($Computer.ComputerName)" -Level "SUCCESS"

                    # Add the computer to the DeletedComputers array
                    $DeletedComputers += $Computer
                } catch {
                    Write-Log -Message "Failed to delete computer $($Computer.ComputerName): $($_.Exception.Message)" -Level "ERROR"
                }
            }
            # Always close the progress bar when the loop exits (success OR failure)
            Write-Progress -Activity 'Deleting disabled computers' -Completed
        } else {
            Write-Log -Message "No computers were deleted." -Level "WARNING"
        }

        # Generate the combined HTML report and save it to the specified path
        Write-CombinedHTMLReport -HtmlPath $reportFileName -InactiveData $InactiveComputers -TotalInactive $InactiveComputers.Count `
            -DisabledData $DisabledComputers -TotalDisabled $DisabledComputers.Count `
            -DeletedData $DeletedComputers -TotalDeleted $DeletedComputers.Count

        Write-Log -Message "Report has been generated at $reportFileName" -Level "SUCCESS"
    } else {
        Write-Log -Message "No computers have been inactive for more than $DaysInactive days." -Level "WARNING"
    }

    Finish-Script -ExitCode 0 -Message "Script completed." -Level "SUCCESS"
} catch {
    Write-Log -Message "Fatal error: $($_.Exception.Message)" -Level "ERROR"
    Finish-Script -ExitCode 1 -Message "Fatal error: $($_.Exception.Message)" -Level "ERROR"
}