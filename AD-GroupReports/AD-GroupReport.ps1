<#
.SYNOPSIS
    Export a single detailed report of all AD groups with nested info and member stats.

.DESCRIPTION
    Retrieves complete group details including name, DN, timestamps, email, 
    member type counts, nested group names and metadata, and exports all into one CSV.
    Logs to a single log file.

.NOTES
    Author: Mohammed Omar
    Date  : 2025-06-19
#>

Import-Module ActiveDirectory

# =======================
# Configuration
# =======================
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$BasePath  = "C:\ADGroupReports"
$LogFile   = "$BasePath\ADGroupExport_$Timestamp.log"
$CsvOutput = "$BasePath\ADGroupUnifiedReport_$Timestamp.csv"

if (!(Test-Path $BasePath)) {
    New-Item -Path $BasePath -ItemType Directory | Out-Null
}

# =======================
# Logging Function
# =======================
function Write-Log {
    param (
        [string]$Message,
        [ValidateSet("INFO", "SUCCESS", "ERROR", "WARNING")]
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp][$Level] $Message"

    $color = switch ($Level) {
        "INFO"     { "Cyan" }
        "SUCCESS"  { "Green" }
        "WARNING"  { "Yellow" }
        "ERROR"    { "Red" }
    }

    Write-Host $logEntry -ForegroundColor $color
    Add-Content -Path $LogFile -Value $logEntry
}

Write-Log "🔄 Starting unified Active Directory group export..." -Level "INFO"

# =======================
# Main Logic
# =======================
$GroupReport = @()

$Groups = Get-ADGroup -Filter * -Properties mail, description, GroupScope, GroupCategory, Members, whenCreated, whenChanged, displayName
Write-Log "📦 Found $($Groups.Count) groups in Active Directory." -Level "INFO"

foreach ($Group in $Groups) {
    $GroupName   = $Group.Name
    $DisplayName = $Group.DisplayName
    $GroupDN     = $Group.DistinguishedName
    $GroupType   = "$($Group.GroupScope)-$($Group.GroupCategory)"
    $GroupEmail  = $Group.Mail
    $Description = $Group.Description
    $Created     = $Group.whenCreated
    $Changed     = $Group.whenChanged
    $NeverModified    = if ($Changed -eq $Created) { "Yes" } else { "No" }
    $IsLikelyInactive = if (($Group.Members.Count -eq 0) -and ($Changed -lt (Get-Date).AddYears(-1))) { "Yes" } else { "No" }

    $CountUser     = 0
    $CountGroup    = 0
    $CountContact  = 0
    $NestedGroupNames = @()
    $NestedGroupCount = 0
    $NestedDetails     = @()

    try {
        $GroupFull = Get-ADGroup -Identity $Group.DistinguishedName -Properties Member
        foreach ($MemberDN in $GroupFull.Member) {
            try {
                $Member = Get-ADObject -Identity $MemberDN -Properties objectClass, name
                switch ($Member.objectClass) {
                    'user'    { $CountUser++ }
                    'group'   {
                        $CountGroup++
                        $NestedGroupCount++
                        $NestedGroupNames += $Member.Name

                        try {
                            $NestedGroup = Get-ADGroup -Identity $Member.DistinguishedName -Properties mail, Members, GroupScope, GroupCategory
                            $NestedDetails += "$($NestedGroup.Name):$($NestedGroup.GroupScope)-$($NestedGroup.GroupCategory):$($NestedGroup.Mail):$($NestedGroup.Members.Count)"
                        } catch {
                            Write-Log "⚠️ Could not get nested group details: $($Member.Name)" -Level "WARNING"
                        }
                    }
                    'contact' { $CountContact++ }
                }
            } catch {
                Write-Log "⚠️ Failed to resolve member DN: $MemberDN" -Level "WARNING"
            }
        }

        $GroupReport += [pscustomobject]@{
            'GroupName'              = $GroupName
            'DisplayName'            = $DisplayName
            'DistinguishedName'      = $GroupDN
            'GroupType'              = $GroupType
            'GroupEmailAddress'      = $GroupEmail
            'Description'            = $Description
            'WhenCreated'            = $Created
            'WhenChanged'            = $Changed
            'NeverModified'          = $NeverModified
            'IsLikelyInactive'       = $IsLikelyInactive
            'MembersCount'           = ($CountUser + $CountGroup + $CountContact)
            'MembersCountByType'     = "User - $CountUser; Group - $CountGroup; Contact - $CountContact"
            'NestedGroupCount'       = $NestedGroupCount
            'NestedGroupNames'       = ($NestedGroupNames -join ', ')
            'NestedGroups_Details'   = ($NestedDetails -join ' | ')
        }

        Write-Log "✅ Processed group: $GroupName (Total Members: $($CountUser + $CountGroup + $CountContact))" -Level "SUCCESS"
    } catch {
        Write-Log "❌ Failed to process group: $GroupName - $_" -Level "ERROR"
    }
}

# =======================
# Export CSV
# =======================
try {
    $GroupReport | Export-Csv -Path $CsvOutput -NoTypeInformation -Encoding UTF8
    Write-Log "📁 Exported unified group report to: $CsvOutput" -Level "SUCCESS"
} catch {
    Write-Log "❌ Failed to export CSV: $_" -Level "ERROR"
}

Write-Log "🏁 Script execution completed." -Level "INFO"
