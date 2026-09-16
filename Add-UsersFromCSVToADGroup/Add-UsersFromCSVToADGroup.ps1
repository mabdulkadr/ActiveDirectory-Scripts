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
    1.0.0

.CHANGELOG
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
#>

#Requires -Version 5.1

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
    Write-Host "`n==========================================" -ForegroundColor Cyan
    Write-Host "Target Group         : $($Group.Name)" -ForegroundColor Cyan
    Write-Host "Distinguished Name   : $($Group.DistinguishedName)" -ForegroundColor Cyan
    if ($Group.Description) {
        Write-Host "Description          : $($Group.Description)" -ForegroundColor Cyan
    }
    Write-Host "==========================================`n" -ForegroundColor Cyan
}
catch {
    Write-Host "❌ Failed to retrieve group details: $_" -ForegroundColor Red
    exit
}

#====================[ Import and Process Users ]====================#

$UserIDs = Import-Csv -Path $CsvPath | Select-Object -ExpandProperty UserID

foreach ($id in $UserIDs) {
    $UPN = "$id$Domain"

    try {
        $User = Get-ADUser -Filter "UserPrincipalName -eq '$UPN'" -ErrorAction Stop

        # Check if user is already a member
        $isMember = Get-ADGroupMember -Identity $GroupName -Recursive | Where-Object {
            $_.DistinguishedName -eq $User.DistinguishedName
        }

        if ($isMember) {
            Write-Host "⚠️ Already a member: $UPN" -ForegroundColor Yellow
            $AlreadyCount++
        }
        else {
            Add-ADGroupMember -Identity $GroupName -Members $User -ErrorAction Stop
            Write-Host "✅ Added: $UPN" -ForegroundColor Green
            $SuccessCount++
        }
    }
    catch {
        Write-Host "❌ Failed to process $UPN : $_" -ForegroundColor Red
        $FailureCount++
    }
}

#====================[ Summary ]====================#

Write-Host "`n============= Summary =============" -ForegroundColor Cyan
Write-Host "✅ Total Users Added      : $SuccessCount" -ForegroundColor Green
Write-Host "⚠️  Already in Group       : $AlreadyCount" -ForegroundColor Yellow
Write-Host "❌ Failed Attempts         : $FailureCount" -ForegroundColor Red
Write-Host "=====================================`n" -ForegroundColor Cyan
