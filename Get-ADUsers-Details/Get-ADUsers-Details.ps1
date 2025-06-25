<#!
.SYNOPSIS
    Get full AD user details from a specific OU including group memberships, password settings, and activity data.

.DESCRIPTION
    This script retrieves all users in the specified OU from Active Directory with detailed account info:
    - Display and account info
    - Password settings and lockout status
    - Group memberships
    - OU and last activity info
    - Sorted by inactivity days

.EXAMPLE
    .\Get-ADUsers-Details.ps1

.NOTES
    Author  : Mohammad Abdelkader
    Website : momar.tech
    Date    : 2025-06-25
    Version : 2.2
#>


#====================[ Configuration ]====================#
$OU = "OU=Employees,OU=QU,DC=QassimU,DC=local"
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
$Results | Sort-Object DaysInactive -Descending | Format-Table -AutoSize

#====================[ Export to CSV ]====================#
$Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8

Write-Host "`n✅ Report saved to: $ExportPath" -ForegroundColor Green
