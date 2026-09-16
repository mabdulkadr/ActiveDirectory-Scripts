<#
.TITLE
    Get-ADUsers-Details - AD user inventory with activity metrics for an OU

.SYNOPSIS
    Exports detailed AD user attributes from an OU, sorted by inactivity, to CSV.

.DESCRIPTION
    Retrieves all users in a configured Organizational Unit and reports display and account
    info, password settings and lockout status, group memberships, OU location, and last
    activity. Results are sorted by inactivity (most inactive first), shown in the console,
    and exported to a timestamped CSV on C:\Reports. The script is read-only.

.TAGS
    ActiveDirectory, Reporting

.PLATFORM
    Windows

.PERMISSIONS
    Read access to the target OU (ActiveDirectory module)

.AUTHOR
    Mohammad Abdelkader Omar

.VERSION
    2.3.0

.CHANGELOG
    2.3.0 (2026-09-16)
    - Standardized help header and section formatting

    2.2.0 (2025-06-25)
    - Initial release

.LASTUPDATE
    2026-09-16

.EXAMPLE
    .\Get-ADUsers-Details.ps1
    Inventories all users in the OU defined by $OU and writes the CSV report to C:\Reports.

.NOTES
    - Requires RSAT (ActiveDirectory module) and read access to the target OU.
    - Configure the $OU variable in the Configuration section before running.
    - Output: C:\Reports\ADUsers-Details_<timestamp>.csv
    - Read-only: no directory objects are modified.
#>

#Requires -Version 5.1

#====================[ Configuration ]====================#
$OU = "OU=Employees,DC=abc,DC=local"
$Today = Get-Date
$TimeStamp = $Today.ToString("yyyyMMdd-HHmmss")
$ExportFolder = "C:\Reports"

if (-not (Test-Path $ExportFolder)) {
    New-Item -Path $ExportFolder -ItemType Directory -Force | Out-Null
}

$ExportPath = "$ExportFolder\ADUsers-Details-$TimeStamp.csv"

#====================[ Load AD Module ]====================#
Import-Module ActiveDirectory

#====================[ Collect Users ]====================#
$Users = Get-ADUser -SearchBase $OU -Filter * -Properties DisplayName, LastLogonDate, Enabled, SamAccountName, whenCreated, whenChanged, PasswordLastSet, MemberOf, PasswordNeverExpires, LockedOut, CannotChangePassword, UserPrincipalName, Title, Department

# Builds one result object per user with enriched account details.
$Results = foreach ($User in $Users) {
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

#====================[ Console Output ]====================#
# Show users sorted by inactivity, most inactive first.
$Results | Sort-Object DaysInactive -Descending | Format-Table -AutoSize

#====================[ Export to CSV ]====================#
$Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8

Write-Host "`n✅ Report saved to: $ExportPath" -ForegroundColor Green
