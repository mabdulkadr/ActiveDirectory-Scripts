<#
.TITLE
    DisableInactiveComputers - Disable or delete stale computer accounts with HTML audit

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
    1.1.0

.CHANGELOG
    1.1.0 (2026-09-16)
    - Standardized help header and script formatting

    1.0.0
    - Initial release

.LASTUPDATE
    2026-09-16

.EXAMPLE
    .\DisableInactiveComputers.ps1 -DaysInactive 180 -SearchBaseOU "OU=Domain Computers,DC=ABC,DC=local"
    Flags any computer that has not logged on for more than 180 days and offers to disable
    and then delete those accounts, writing a timestamped audit report to C:\.

.NOTES
    - Requires RSAT (ActiveDirectory module) and permissions to disable/delete computer objects.
    - HTML report: C:\Computers_Report_<timestamp>.html
    - Optional action log: C:\InactiveComputers.log (used by the Log-Message helper).
    - Test against a staging OU before running in production; deletion is irreversible.
#>

#Requires -Version 5.1

param (
    [int]$DaysInactive     = 180,                                            # Number of days of inactivity for filtering inactive computers
    [string]$SearchBaseOU  = "OU=Domain Computers,DC=ABC,DC=local"           # Organizational Unit (OU) for searching computers
)

# Writes a timestamped line to the optional action log file.
function Log-Message {
    param (
        [string]$Message,      # The message to log
        [string]$LogPath = "C:\InactiveComputers.log"  # Default log file path
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp - $Message"
    Add-Content -Path $LogPath -Value $logEntry
}

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

# Generate a timestamp for the report filename
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$reportFileName = "C:\Computers_Report_$timestamp.html"

# Start processing inactive computers
Write-Host "Processing inactive computers based on LastLogonDate and $DaysInactive days of inactivity..." -ForegroundColor Cyan

# Get the threshold date for inactivity (current date minus DaysInactive)
$InactiveThreshold = (Get-Date).AddDays(-$DaysInactive)

# Fetch all computers from the specified Organizational Unit (OU)
try {
    $AllComputers = Get-ADComputer -Filter * -SearchBase $SearchBaseOU -Property Name, LastLogonDate -ErrorAction Stop
} catch {
    Write-Error "Failed to retrieve computers from the specified OU: $_"
    Exit
}

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

# Initialize arrays for Disabled and Deleted computers
$DisabledComputers = @()
$DeletedComputers = @()

# Check if there are any inactive computers to process
if ($InactiveComputers.Count -gt 0) {
    # List inactive computers in the console
    Write-Host "`nInactive Computers List:" -ForegroundColor Cyan
    $InactiveComputers | Format-Table -Property ComputerName, LastLogonDate -AutoSize

    # Ask if you want to disable the inactive computers
    $DisableInput = Read-Host "`nDo you want to disable all the inactive computers? (Yes/No)"
    
    if ($DisableInput -eq "Yes") {
        # Loop through inactive computers and disable them
        foreach ($Computer in $InactiveComputers) {
            try {
                Write-Host "Disabling computer: $($Computer.ComputerName)" -ForegroundColor Yellow
                Disable-ADAccount -Identity $Computer.DistinguishedName -ErrorAction Stop
                Write-Host "Disabled: $($Computer.ComputerName)" -ForegroundColor Green
                
                # Add the computer to the DisabledComputers array
                $DisabledComputers += $Computer
            } catch {
                Write-Error "Failed to disable computer $($Computer.ComputerName): $_"
            }
        }
    } else {
        Write-Host "`nNo computers were disabled." -ForegroundColor Yellow
    }

    # Ask if you want to delete the disabled computers
    $DeleteInput = Read-Host "`nDo you want to delete all disabled computers? (Yes/No)"
    
    if ($DeleteInput -eq "Yes" -and $DisabledComputers.Count -gt 0) {
        # Loop through disabled computers and delete them
        foreach ($Computer in $DisabledComputers) {
            try {
                Write-Host "Deleting computer: $($Computer.ComputerName)" -ForegroundColor Red
                Remove-ADComputer -Identity $Computer.DistinguishedName -Confirm:$false -ErrorAction Stop
                Write-Host "Deleted: $($Computer.ComputerName)" -ForegroundColor Green

                # Add the computer to the DeletedComputers array
                $DeletedComputers += $Computer
            } catch {
                Write-Error "Failed to delete computer $($Computer.ComputerName): $_"
            }
        }
    } else {
        Write-Host "`nNo computers were deleted." -ForegroundColor Yellow
    }

    # Generate the combined HTML report and save it to the specified path
    Write-CombinedHTMLReport -HtmlPath $reportFileName -InactiveData $InactiveComputers -TotalInactive $InactiveComputers.Count `
        -DisabledData $DisabledComputers -TotalDisabled $DisabledComputers.Count `
        -DeletedData $DeletedComputers -TotalDeleted $DeletedComputers.Count

    Write-Host "`nReport has been generated at $reportFileName" -ForegroundColor Green

} else {
    Write-Host "`nNo computers have been inactive for more than $DaysInactive days." -ForegroundColor Yellow
}

Write-Host "`nScript completed." -ForegroundColor Cyan
