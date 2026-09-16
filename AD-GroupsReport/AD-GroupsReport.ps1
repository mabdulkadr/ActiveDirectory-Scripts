<#
.TITLE
    AD-GroupsReport - Full AD group inventory with nesting and member statistics

.SYNOPSIS
    Exports a detailed CSV report of all AD groups with nested info and member stats.

.DESCRIPTION
    Retrieves complete group details for every group in the domain — name, DN, timestamps,
    email, member-type counts (User / Group / Contact), nested group names and metadata, and
    inactivity flags — and exports everything into one timestamped CSV file. Also writes a
    matching execution log. If the ActiveDirectory module is missing, the script attempts a
    self-install of the RSAT tools (elevation required) before failing cleanly.

.TAGS
    ActiveDirectory, Reporting

.PLATFORM
    Windows

.PERMISSIONS
    Read access to groups and members (ActiveDirectory module); elevation only for the
    optional RSAT self-install

.AUTHOR
    Mohammad Abdelkader Omar

.VERSION
    1.2.0

.CHANGELOG
    1.2.0 (2026-09-16)
    - Added canonical CLI logging, elevated RSAT install check, Write-Progress, and exit codes
    - Report path now anchors beside the script (Law 12)

    1.1.0 (2026-09-16)
    - Standardized help header and section formatting

    1.0.0 (2025-06-19)
    - Initial release

.LASTUPDATE
    2026-09-16

.EXAMPLE
    .\AD-GroupsReport.ps1
    Exports the full group inventory to <script-folder>\Reports\ADGroupsReport_<timestamp>.csv.

.NOTES
    - Requires RSAT (ActiveDirectory module); the script attempts to install it when missing.
    - Output folder: <script-folder>\Reports (created on first run).
    - Run from an elevated PowerShell session for the optional module install.
    - Read-only: queries groups and members but never modifies objects.
    - Exit code 0 on completion; 1 on module install/import failure.
#>

#Requires -Version 5.1

$SolutionName = 'AD-GroupsReport'
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

#====================[ Module Check ]====================#

# Returns $true only when the process runs from an elevated (admin) context.
function Test-IsElevated {
    return ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Log -Message "ActiveDirectory module not found. Checking for elevation..." -Level "WARNING"

    if (-not (Test-IsElevated)) {
        Write-Log -Message "RSAT install requires an elevated session. Re-run as Administrator." -Level "ERROR"
        Finish-Script -ExitCode 1 -Message "Elevation required to install the ActiveDirectory module." -Level "ERROR"
    }

    # Attempt installation for Windows 10/11 and Server 2019+
    try {
        Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" -ErrorAction Stop
        Write-Log -Message "RSAT: Active Directory tools installed successfully." -Level "SUCCESS"
    } catch {
        Write-Log -Message "Failed to install ActiveDirectory module. Please install RSAT manually." -Level "ERROR"
        Finish-Script -ExitCode 1 -Message "Failed to install the ActiveDirectory module." -Level "ERROR"
    }
}

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Log -Message "ActiveDirectory module imported successfully." -Level "SUCCESS"
} catch {
    Write-Log -Message "Failed to import ActiveDirectory module. Ensure RSAT tools are correctly installed." -Level "ERROR"
    Finish-Script -ExitCode 1 -Message "Failed to import the ActiveDirectory module." -Level "ERROR"
}

#====================[ Configuration ]====================#

# Report path law: seed output beside the script, never from Get-Location alone.
$scriptBase = if ($PSScriptRoot) { $PSScriptRoot }
elseif ($PSCommandPath) { Split-Path -Parent $PSCommandPath }
elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path }
else { (Get-Location).Path }
$BasePath = Join-Path $scriptBase "Reports"
if (-not (Test-Path -LiteralPath $BasePath)) {
    $null = [System.IO.Directory]::CreateDirectory($BasePath)
    Write-Log -Message "Created output directory: $BasePath" -Level "INFO"
}

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$CsvOutput = Join-Path $BasePath "ADGroupsReport_$Timestamp.csv"

#====================[ Main Logic ]====================#

$GroupReport = @()

$Groups = Get-ADGroup -Filter * -Properties mail, description, GroupScope, GroupCategory, Members, whenCreated, whenChanged, displayName
Write-Log -Message "Found $($Groups.Count) groups in Active Directory." -Level "INFO"

$processedCount = 0
$total = @($Groups).Count

foreach ($Group in $Groups) {
    $processedCount++
    # Activity = persistent label; Status = current item; PercentComplete = 0..100
    Write-Progress -Activity 'Processing AD groups' `
        -Status "Group $processedCount of $total : $($Group.Name)" `
        -PercentComplete (($processedCount / [Math]::Max($total, 1)) * 100)

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
                            Write-Log -Message "Could not get nested group details: $($Member.Name)" -Level "WARNING"
                        }
                    }
                    'contact' { $CountContact++ }
                }
            } catch {
                Write-Log -Message "Failed to resolve member DN: $MemberDN" -Level "WARNING"
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

        if ($processedCount % 10 -eq 0) {
            Write-Log -Message "Processed $processedCount of $total groups..." -Level "INFO"
        }
    } catch {
        Write-Log -Message "Failed to process group: $GroupName - $($_.Exception.Message)" -Level "ERROR"
    }
}

# Always close the progress bar when the loop exits (success OR failure)
Write-Progress -Activity 'Processing AD groups' -Completed

#====================[ Export CSV ]====================#
try {
    $GroupReport | Export-Csv -Path $CsvOutput -NoTypeInformation -Encoding UTF8
    Write-Log -Message "Exported unified group report to: $CsvOutput" -Level "SUCCESS"
} catch {
    Write-Log -Message "Failed to export CSV: $($_.Exception.Message)" -Level "ERROR"
    Finish-Script -ExitCode 1 -Message "Failed to export the group report." -Level "ERROR"
}

Finish-Script -ExitCode 0 -Message "Script execution completed. $($GroupReport.Count) group(s) exported." -Level "SUCCESS"