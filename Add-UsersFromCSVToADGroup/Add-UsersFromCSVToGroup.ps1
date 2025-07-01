<#
.SYNOPSIS
    Adds a list of users from a CSV file to an Active Directory group.

.DESCRIPTION
    This script reads user IDs from a CSV file, appends the UPN domain (@qu.edu.sa),
    verifies if the user exists in Active Directory, checks if the user is already in the group,
    and adds them to the specified AD group if not already present.

.EXAMPLE
    PS> .\Add-UsersFromCSVToGroup.ps1

.NOTES
    Author  : Mohammed Omar
    Requires: RSAT (ActiveDirectory module)
#>

# ========================= Configuration =========================

$GroupName = "Blind-UserLogin-Exception"
$CsvPath = "C:\Add-UsersFromCSVToGroup\users.csv"
$Domain = "@abc.local"

$SuccessCount = 0
$FailureCount = 0
$AlreadyCount = 0

# ========================= Display Group Details =========================

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

# ========================= Import and Process Users =========================

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

# ========================= Summary =========================

Write-Host "`n============= Summary =============" -ForegroundColor Cyan
Write-Host "✅ Total Users Added      : $SuccessCount" -ForegroundColor Green
Write-Host "⚠️  Already in Group       : $AlreadyCount" -ForegroundColor Yellow
Write-Host "❌ Failed Attempts         : $FailureCount" -ForegroundColor Red
Write-Host "=====================================`n" -ForegroundColor Cyan
