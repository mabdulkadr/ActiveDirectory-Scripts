<#
.SYNOPSIS
    Get-ADHealth.ps1 — Domain Controller Health Check (card-style HTML + email)

.DESCRIPTION
    Performs a comprehensive health check of all domain controllers in a specified domain
    or across the entire Active Directory forest, then outputs a modern, responsive
    HTML report composed of per-DC "cards" grouped by Domain → Site, with badges,
    progress bars, and a built-in explainer for each check.

    Collectors (per DC):
      - DNS resolution (Resolve-DnsName)
      - ICMP reachability (Test-Connection)
      - Uptime and last reboot (Win32_OperatingSystem)
      - Time synchronization offset (w32tm /stripchart)
      - OS drive free space (% + GB) (Win32_LogicalDisk)
      - Service status (DNS, NTDS, Netlogon)
      - DCDIAG test suite (Connectivity, Replication, SysVol, FSMO, etc.)

    Output:
      - Card-based HTML (responsive, easy to read)
      - Optional email (inline + attachment)

.PARAMETER DomainName
    One or more domain DNS names to scope the check. If omitted, scans all domains in the forest.

.PARAMETER ReportPath
    Folder to save the HTML report. Default: C:\Reports

.PARAMETER SendEmail
    If specified, sends the report via SMTP (inline HTML + file attachment).

.PARAMETER Subject
    Email subject. Default: "Domain Controller Health Check"

.PARAMETER UserFrom
    SMTP From address (e.g., smtp-reports@yourdomain.com)

.PARAMETER UserTo
    One or more recipients.

.PARAMETER SmtpServer
    SMTP server hostname. Default: smtp.office365.com

.PARAMETER Port
    SMTP port. Default: 587

.PARAMETER Credential
    PSCredential for SMTP authentication. If not supplied, you can also pass -Password to be combined with -UserFrom.

.PARAMETER Password
    SecureString password. If provided (and -Credential not provided), it will be combined with -UserFrom.

.LINK
    Resolve-DnsName (Microsoft Docs): https://learn.microsoft.com/powershell/module/dnsclient/resolve-dnsname
    DCDIAG reference:                 https://learn.microsoft.com/windows-server/administration/windows-commands/dcdiag
    Windows Time (W32Time):           https://learn.microsoft.com/windows-server/networking/windows-time-service/windows-time-service-tools-and-settings
    Win32_OperatingSystem:            https://learn.microsoft.com/windows/win32/cimwin32prov/win32-operatingsystem
    Win32_LogicalDisk:                https://learn.microsoft.com/windows/win32/cimwin32prov/win32-logicaldisk
#>

[CmdletBinding()]
param(
    [string[]]$DomainName,
    [string]$ReportPath = 'C:\Reports',

    [switch]$SendEmail,
    [string]$Subject    = 'Domain Controller Health Check',
    [string]$UserFrom   = 'smtp-reports@yourdomain.com',
    [string[]]$UserTo   = @("it-admins@yourdomain.com"),
    [string]$SmtpServer = 'smtp.office365.com', # Your SMTP server (EOP/Exchange Online)
    [int]$Port          = 587,

	[securestring]$Password,
    [pscredential]$Credential
    
)

$Password   = ConvertTo-SecureString "<REPLACE_WITH_SMTP_PASSWORD>" -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential ($UserFrom, $Password)

#region Pre-flight & Globals
# -------------------------

# Require AD module
try {
    Import-Module ActiveDirectory -ErrorAction Stop
} catch {
    Write-Error "ActiveDirectory module not found. Install RSAT/AD tools and try again."
    return
}

# Check for DCDIAG
if (-not (Get-Command dcdiag.exe -ErrorAction SilentlyContinue)) {
    Write-Warning "DCDIAG not found in PATH. Install RSAT or run on a DC for full test coverage."
}

# Email credential fallback if only -Password was provided
if ($SendEmail -and -not $Credential -and $Password) {
    $Credential = [pscredential]::new($UserFrom, $Password)
}

# Collections & timestamps
$allTestedDomainControllers = [System.Collections.Generic.List[hashtable]]::new()
$now                        = Get-Date
$reportTime                 = $now
$reportFileNameTime         = $now.ToString('yyyyMMdd_HHmmss')

# Health thresholds (tune here)
$Thresholds = @{
    UptimeWarnHours  = 24
    FreePctFail      = 5
    FreePctWarn      = 30
    FreeGBFail       = 5
    FreeGBWarn       = 10
    TimeWarnSeconds  = 0.5
    TimeFailSeconds  = 1.0
}

# ---- Severity policy for overall card state ----
# Anything in $CriticalBinaryProps that fails ⇒ Critical immediately.
# Anything in $WarningBinaryProps that fails (and no criticals) ⇒ Warning.
$DCDiagCritical = @(
  'DCDIAG: Connectivity',
  'DCDIAG: SysVolCheck',
  'DCDIAG: Replications',
  'DCDIAG: ObjectsReplicated',
  'DCDIAG: NetLogons',
  'DCDIAG: MachineAccount',
  'DCDIAG: FSMO Check',
  'DCDIAG: FSMO KnowsOfRoleHolders'
)

# Explicitly keep these as WARNING (your request)
$DCDiagWarning = @(
  'DCDIAG: DFSREvent',
  'DCDIAG: SystemLog'
)

# Binary probes treated as Critical on failure
$CriticalBinaryProps = @(
  'DNS','Ping','DNS Service','NTDS Service','NetLogon Service'
) + $DCDiagCritical

# Binary probes treated as Warning on failure
$WarningBinaryProps  = @() + $DCDiagWarning

#endregion Pre-flight & Globals

#region Discovery
# ---------------

function Get-AllDomains {
    (Get-ADForest).Domains
}

function Get-AllDomainControllers {
    param([Parameter(Mandatory)][string]$ComputerName)
    Get-ADDomainController -Filter * -Server $ComputerName | Sort-Object HostName
}

#endregion Discovery

#region Probes (per-DC collectors)
# --------------------------------

function Get-DomainControllerNSLookup {
    param([Parameter(Mandatory)][string]$ComputerName)
    try {
        $null = Resolve-DnsName -Name $ComputerName -Type A -ErrorAction Stop
        'Success'
    } catch { 'Fail' }
}

function Get-DomainControllerPingStatus {
    param([Parameter(Mandatory)][string]$ComputerName)
    if (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) { 'Success' } else { 'Fail' }
}

function Get-DomainControllerUpTime {
    param([Parameter(Mandatory)][string]$ComputerName)
    if (-not (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet)) { return 'Fail' }
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $ComputerName -ErrorAction Stop
        $hours = [int](((Get-Date) - $os.LastBootUpTime).TotalHours)
        $hours
    } catch { 'CIM Failure' }
}

function Get-TimeDifference {
    param([Parameter(Mandatory)][string]$ComputerName)
    if (-not (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet)) { return 'Fail' }
    try {
        $out = & w32tm /stripchart /computer:$ComputerName /samples:1 /dataonly 2>$null
        # Typical last line: "xx:xx:xx, 0.0012345s"
        $line = $out | Select-Object -Last 1
        if ($line -match '(-?\d+(?:[.,]\d+)?)s') {
            $num = ($matches[1] -replace ',', '.')
            [math]::Round([double]$num, 1, [System.MidpointRounding]::AwayFromZero)
        } else { 'Fail' }
    } catch { 'Fail' }
}

function Get-DomainControllerServices {
    param([Parameter(Mandatory)][string]$ComputerName)
    $o = [ordered]@{ DNSService=$null; NTDSService=$null; NETLOGONService=$null }

    if (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {
        $svc = Get-Service -ComputerName $ComputerName -Name 'DNS' -ErrorAction SilentlyContinue
        if ($svc -and $svc.Status -eq 'Running') { $o.DNSService = 'Success' } else { $o.DNSService = 'Fail' }

        $svc = Get-Service -ComputerName $ComputerName -Name 'NTDS' -ErrorAction SilentlyContinue
        if ($svc -and $svc.Status -eq 'Running') { $o.NTDSService = 'Success' } else { $o.NTDSService = 'Fail' }

        $svc = Get-Service -ComputerName $ComputerName -Name 'NetLogon' -ErrorAction SilentlyContinue
        if ($svc -and $svc.Status -eq 'Running') { $o.NETLOGONService = 'Success' } else { $o.NETLOGONService = 'Fail' }
    } else {
        $o.DNSService = 'Fail'
        $o.NTDSService = 'Fail'
        $o.NETLOGONService = 'Fail'
    }
    [pscustomobject]$o
}

function Get-DomainControllerDCDiagTestResults {
    param([Parameter(Mandatory)][string]$ComputerName)

    $result = [ordered]@{
        ServerName         = $ComputerName
        Connectivity       = $null
        Advertising        = $null
        FrsEvent           = $null
        DFSREvent          = $null
        SysVolCheck        = $null
        KccEvent           = $null
        KnowsOfRoleHolders = $null
        MachineAccount     = $null
        NCSecDesc          = $null
        NetLogons          = $null
        ObjectsReplicated  = $null
        Replications       = $null
        RidManager         = $null
        Services           = $null
        SystemLog          = $null
        VerifyReferences   = $null
        CheckSDRefDom      = $null
        CrossRefValidation = $null
        LocatorCheck       = $null
        Intersite          = $null
        FSMOCheck          = $null
    }

    if (-not (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet)) {
        foreach ($k in $result.Keys | Where-Object { $_ -ne 'ServerName' }) { $result[$k] = 'Failed' }
        return [pscustomobject]$result
    }

    if (-not (Get-Command dcdiag.exe -ErrorAction SilentlyContinue)) {
        foreach ($k in $result.Keys | Where-Object { $_ -ne 'ServerName' }) { $result[$k] = 'Unknown' }
        return [pscustomobject]$result
    }

    $tests = @(
        'Connectivity','Advertising','FrsEvent','DFSREvent','SysVolCheck','KccEvent','KnowsOfRoleHolders',
        'MachineAccount','NCSecDesc','NetLogons','ObjectsReplicated','Replications','RidManager',
        'Services','SystemLog','VerifyReferences','CheckSDRefDom','CrossRefValidation','LocatorCheck','Intersite','FSMOCheck'
    )
    $params = @("/s:$ComputerName") + ($tests | ForEach-Object { "/test:$($_)" })
    $lines  = (dcdiag.exe @params) -split "(`r`n|`n)"

    $current = $null
    foreach ($line in $lines) {
        if ($line -match 'Starting test:\s*(.+)$') {
            $current = ($matches[1]).Trim()
            continue
        }
        if ($null -ne $current -and ($line -match 'passed test' -or $line -match 'failed test')) {
            if ($line -match 'passed test') {
                $result[$current] = 'Passed'
            } else {
                $result[$current] = 'Failed'
            }
            $current = $null
        }
    }

    [pscustomobject]$result
}

function Get-DomainControllerOSDriveFreeSpace {
    param([Parameter(Mandatory)][string]$ComputerName)
    if (-not (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet)) { return 'Fail' }
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $ComputerName -ErrorAction Stop
        $sys = $os.SystemDrive
        $disk = Get-CimInstance -ClassName Win32_LogicalDisk -ComputerName $ComputerName -Filter "DeviceID='$sys'" -ErrorAction Stop
        if ($disk.Size -gt 0) { [math]::Round(($disk.FreeSpace / $disk.Size) * 100) } else { 'CIM Failure' }
    } catch { 'CIM Failure' }
}

function Get-DomainControllerOSDriveFreeSpaceGB {
    param([Parameter(Mandatory)][string]$ComputerName)
    if (-not (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet)) { return 'Fail' }
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $ComputerName -ErrorAction Stop
        $sys = $os.SystemDrive
        $disk = Get-CimInstance -ClassName Win32_LogicalDisk -ComputerName $ComputerName -Filter "DeviceID='$sys'" -ErrorAction Stop
        [math]::Round(($disk.FreeSpace / 1GB), 2)
    } catch { 'CIM Failure' }
}

#endregion Probes

#region HTML helpers & builder (card UI)
# --------------------------------------

Add-Type -AssemblyName System.Web | Out-Null

function HtmlEncode {
    param([object]$Value)
    if ($null -eq $Value) { return '' }
    [System.Web.HttpUtility]::HtmlEncode([string]$Value)
}

function Get-StatusClass {
    param([string]$Metric, [object]$Value)

    $v = [string]$Value
    if ($v -in @('Success','Passed','Pass')) { return 'pass' }
    if ($v -in @('Fail','Failed','CIM Failure','Could not test server uptime.')) { return 'fail' }

    switch ($Metric) {
        'Uptime (hours)' {
            if ($v -eq 'Fail') { return 'fail' }
            if ($v -eq 'CIM Failure') { return 'warn' }
            $h = [int]$v
            if ($h -le $Thresholds.UptimeWarnHours) { 'warn' } else { 'pass' }
        }
        'OS Free Space (%)' {
            if ($v -eq 'Fail' -or $v -eq 'CIM Failure') { return 'fail' }
            $p = [int]$v
            if     ($p -le $Thresholds.FreePctFail) { 'fail' }
            elseif ($p -le $Thresholds.FreePctWarn) { 'warn' }
            else { 'pass' }
        }
        'OS Free Space (GB)' {
            if ($v -eq 'Fail' -or $v -eq 'CIM Failure') { return 'fail' }
            $g = [double]$v
            if     ($g -lt $Thresholds.FreeGBFail) { 'fail' }
            elseif ($g -lt $Thresholds.FreeGBWarn) { 'warn' }
            else { 'pass' }
        }
        'Time offset (seconds)' {
            if ($v -eq 'Fail') { return 'fail' }
            $d = [double]$v
            if     ($d -ge $Thresholds.TimeFailSeconds) { 'fail' }
            elseif ($d -ge $Thresholds.TimeWarnSeconds) { 'warn' }
            else { 'pass' }
        }
        default { '' }
    }
}

function New-ProgressBar {
    param([int]$Percent = 0, [string]$Class = 'pass')
@"
<div class='meter $Class'><div class='fill' style='width:$([Math]::Max(0,[Math]::Min(100,$Percent)))%'></div></div>
"@
}

function Get-DCHealthState {
    param([hashtable]$DC)

    $failTokens = @('Fail','Failed','CIM Failure','Could not test server uptime.')

    # 1) Any critical binary failure? -> Critical
    foreach ($p in $CriticalBinaryProps) {
        if ($failTokens -contains ([string]$DC[$p])) { return 'Critical' }
    }

    # 2) Track warnings from warning-class binary props
    $warnTriggered = $false
    foreach ($p in $WarningBinaryProps) {
        if ($failTokens -contains ([string]$DC[$p])) { $warnTriggered = $true }
    }

    # 3) Numeric thresholds (can escalate to Critical)
    try {
        $gb = [double]$DC['OS Free Space (GB)']
        if ($gb -lt $Thresholds.FreeGBFail) { return 'Critical' }
        if ($gb -lt $Thresholds.FreeGBWarn) { $warnTriggered = $true }
    } catch {}

    try {
        $pct = [int]$DC['OS Free Space (%)']
        if ($pct -le $Thresholds.FreePctFail) { return 'Critical' }
        if ($pct -le $Thresholds.FreePctWarn) { $warnTriggered = $true }
    } catch {}

    try {
        $ofs = [double]$DC['Time offset (seconds)']
        if ($ofs -ge $Thresholds.TimeFailSeconds) { return 'Critical' }
        if ($ofs -ge $Thresholds.TimeWarnSeconds) { $warnTriggered = $true }
    } catch {}

    try {
        $upt = [int]$DC['Uptime (hours)']
        if ($upt -le $Thresholds.UptimeWarnHours) { $warnTriggered = $true }
    } catch {}

    if ($warnTriggered) { return 'Warning' }
    return 'Healthy'
}


function Build-EmailSafeHtmlReport {
    param(
        [Parameter(Mandatory)][System.Collections.IEnumerable]$Items,
        [string]$Forest     = (Get-ADForest).Name,
        [datetime]$Generated = (Get-Date),
        [string]$Title       = 'Domain Controller Health – (Email View)',
        [string]$Description = 'Lightweight email-safe view (inline styles, table layout).'
    )

    # Reuse your existing helpers:
    # - HtmlEncode
    # - Get-StatusClass
    # - Get-DCHealthState

    # Precompute states
    $states = @{'Healthy'=0; 'Warning'=0; 'Critical'=0}
    foreach ($i in $Items) {
        $s = Get-DCHealthState -DC $i
        $states[$s]++
        $i['__State'] = $s
    }
    $total = @($Items).Count

    # Simple colors for chips/badges (email-safe)
    $chipMap = @{
        pass = @{ fg = '#1e4620'; bg = '#e6f4ea'; bd = '#cbe9d3' }
        warn = @{ fg = '#735f1e'; bg = '#fff4d6'; bd = '#ffe7a3' }
        fail = @{ fg = '#6b1f2c'; bg = '#fde2e6'; bd = '#f7c2cb' }
    }
    function NewChip([string]$text,[string]$cls){
        $m = $chipMap[$cls]; if(-not $m){ $m = $chipMap['pass'] }
        "<span style=""display:inline-block;margin:2px 4px 2px 0;padding:3px 8px;border:1px solid $($m.bd);background:$($m.bg);color:$($m.fg);font:12px 'Segoe UI',Arial;border-radius:999px;"">$text</span>"
    }
    function NewBadge([string]$label,[string]$cls){
        $m = $chipMap[$cls]; if(-not $m){ $m = $chipMap['pass'] }
        "<span style=""display:inline-block;margin-left:6px;padding:6px 10px;border:1px solid $($m.bd);background:$($m.bg);color:$($m.fg);font:12px 'Segoe UI',Arial;border-radius:999px;font-weight:600;"">$label</span>"
    }
    function NewMeter([int]$pct,[string]$cls){
        $pct = [Math]::Max(0,[Math]::Min(100,$pct))
        $bar = switch($cls){ 'fail' {'#ef476f'} 'warn' {'#ffd166'} default {'#20c997'} }
        @"
<div style="width:100%;height:8px;border:1px solid #e0e0e0;border-radius:6px;background:#f4f4f4;">
  <div style="width:${pct}%;height:8px;background:${bar};border-radius:6px;"></div>
</div>
"@
    }

    $sb = New-Object System.Text.StringBuilder
    $who = [Security.Principal.WindowsIdentity]::GetCurrent().Name

    # Email-safe wrapper table (centered 820px)
    [void]$sb.AppendLine(@"
<!doctype html><html><body style="margin:0;padding:0;background:#ffffff;">
<table role="presentation" cellpadding="0" cellspacing="0" border="0" align="center" width="100%" style="background:#ffffff;">
  <tr><td align="center" style="padding:16px;">
    <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="820" style="width:820px;max-width:100%;font-family:'Segoe UI',Arial,sans-serif;color:#0f172a;">
      <tr>
        <td style="padding:8px 0 4px 0;">
          <div style="font-size:22px;font-weight:700;color:#0f172a;margin-top:5px;">$(HtmlEncode $Title)</div>
          <div style="font-size:13px;color:#334155;margin-top:5px;">$(HtmlEncode $Description)</div>
          <div style="font-size:12px;color:#475569;margin:7px 0;"><b>Run as : </b> $who <br><b>Host : </b> $($env:COMPUTERNAME) <br><b>Generated : </b> $(HtmlEncode ($Generated.ToString('yyyy-MM-dd HH:mm:ss')))  </div>
        </td>
      </tr>
      <tr>
        <td style="padding:6px 0 10px 0;">
          $(NewBadge "Total DCs: $total" 'pass')$(NewBadge "Healthy: $($states['Healthy'])" 'pass')$(NewBadge "Warnings: $($states['Warning'])" 'warn')$(NewBadge "Critical: $($states['Critical'])" 'fail')
        </td>
      </tr>
"@)

    # Domain -> Site
    $byDomain = $Items | Group-Object Domain
    foreach ($d in $byDomain) {
        $bySite = $d.Group | Group-Object Site
        foreach ($s in $bySite) {
            foreach ($dc in $s.Group) {
                # Classes/values
                $dnsClass  = Get-StatusClass 'DNS'                   $dc['DNS']
                $pingClass = Get-StatusClass 'Ping'                  $dc['Ping']
                $dnsSvc    = Get-StatusClass 'DNS Service'           $dc['DNS Service']
                $ntdsSvc   = Get-StatusClass 'NTDS Service'          $dc['NTDS Service']
                $netlogSvc = Get-StatusClass 'NetLogon Service'      $dc['NetLogon Service']
                $uptClass  = Get-StatusClass 'Uptime (hours)'        $dc['Uptime (hours)']
                $gbClass   = Get-StatusClass 'OS Free Space (GB)'    $dc['OS Free Space (GB)']
                $pctClass  = Get-StatusClass 'OS Free Space (%)'     $dc['OS Free Space (%)']
                $ofsClass  = Get-StatusClass 'Time offset (seconds)' $dc['Time offset (seconds)']
                $pct = 0; try { $pct = [int]$dc['OS Free Space (%)'] } catch {}

                $state = $dc['__State']
                $stateBadge = switch($state){ 'Critical' {NewBadge 'Critical' 'fail'} 'Warning' {NewBadge 'Warning' 'warn'} default {NewBadge 'Healthy' 'pass'} }

                # FSMO chips
                $roles = ''
                if ($dc['Operation Master Roles'] -and $dc['Operation Master Roles'].Count -gt 0) {
                    foreach($r in $dc['Operation Master Roles']){ $roles += (NewChip (HtmlEncode $r) 'pass') }
                } else { $roles = (NewChip 'None' 'warn') }

                # DCDIAG chips
                $tests = @(
                    'DCDIAG: Connectivity','DCDIAG: Advertising','DCDIAG: FrsEvent','DCDIAG: DFSREvent','DCDIAG: SysVolCheck',
                    'DCDIAG: KccEvent','DCDIAG: FSMO KnowsOfRoleHolders','DCDIAG: MachineAccount','DCDIAG: NCSecDesc',
                    'DCDIAG: NetLogons','DCDIAG: ObjectsReplicated','DCDIAG: Replications','DCDIAG: RidManager',
                    'DCDIAG: Services','DCDIAG: SystemLog','DCDIAG: VerifyReferences','DCDIAG: CheckSDRefDom',
                    'DCDIAG: CrossRefValidation','DCDIAG: LocatorCheck','DCDIAG: Intersite','DCDIAG: FSMO Check'
                )
                $testHtml = ''
                foreach($t in $tests){
                    $cls = Get-StatusClass $t $dc[$t]
                    $val = HtmlEncode $dc[$t]
                    $name = HtmlEncode ($t -replace '^DCDIAG:\s*','')
                    $testHtml += (NewChip "${name}: $val" $cls)
                }

                # “Card” table (email-safe)
                [void]$sb.AppendLine(@"
<tr><td style="padding:6px 0 10px 0;">
  <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border:1px solid #e5e7eb;border-radius:10px;">
    <tr>
      <td style="padding:12px;">
        <div style="font-weight:700;color:#0f172a;font-size:16px;text-transform:uppercase;">$(HtmlEncode $dc['Server'])</div>
        <div style="font-size:12px;color:#475569;margin:4px 0 6px 0;">
          <b>Domain:</b> $(HtmlEncode $dc['Domain']) &nbsp;•&nbsp;
          <b>IP:</b> $(HtmlEncode $dc['IPv4 Address']) &nbsp;•&nbsp;
          <b>OS:</b> $(HtmlEncode $dc['OS Version'])
        </div>
        <div style="margin-bottom:6px;"><b style="font-size:12px;color:#0f172a;">FSMO:</b> $roles</div>
        <div style="margin-bottom:6px;">$stateBadge</div>

        <div style="margin:6px 0;">
          $(NewChip "DNS: $(HtmlEncode $dc['DNS'])"               $dnsClass)
          $(NewChip "Ping: $(HtmlEncode $dc['Ping'])"              $pingClass)
          $(NewChip "DNS Service: $(HtmlEncode $dc['DNS Service'])" $dnsSvc)
          $(NewChip "NTDS: $(HtmlEncode $dc['NTDS Service'])"       $ntdsSvc)
          $(NewChip "NetLogon: $(HtmlEncode $dc['NetLogon Service'])" $netlogSvc)
          $(NewChip "Time Offset (s): $(HtmlEncode $dc['Time offset (seconds)'])" $ofsClass)
          $(NewChip "Uptime (h): $(HtmlEncode $dc['Uptime (hours)'])" $uptClass)
          $(NewChip "Free (GB): $(HtmlEncode $dc['OS Free Space (GB)'])" $gbClass)
        </div>

        <div style="display:flex;justify-content:space-between;align-items:center;font-size:12px;color:#475569;margin:4px 0;">
          <div>OS Free Space (%)</div>
          <div style="font-weight:700;color:#0f172a;">$pct%</div>
        </div>
        $(NewMeter $pct $pctClass)


        <div style="font-size:12px;color:#334155;margin-top:8px;">DCDIAG checks</div>
        <div>$testHtml</div>

        <div style="font-size:12px;color:#64748b;margin-top:6px;">Processing time: $(HtmlEncode $dc['Processing Time (seconds)']) s</div>
      </td>
    </tr>
  </table>
</td></tr>
"@)
            }
        }
    }

    # -----------------------
    # Glossary (email-safe)
    # -----------------------
    $glossary = @(
        @{ Term = 'DNS (Resolve-DnsName)'; Meaning = 'Resolves the DC hostname to A/AAAA records.'; Link = 'https://learn.microsoft.com/powershell/module/dnsclient/resolve-dnsname' },
        @{ Term = 'Ping'; Meaning = 'ICMP reachability to the DC.'; Link = '' },
        @{ Term = 'Uptime (h)'; Meaning = 'Hours since last reboot (Win32_OperatingSystem.LastBootUpTime).'; Link = 'https://learn.microsoft.com/windows/win32/cimwin32prov/win32-operatingsystem' },
        @{ Term = 'Free (GB) / OS Free Space (%)'; Meaning = 'System drive free space from Win32_LogicalDisk.'; Link = 'https://learn.microsoft.com/windows/win32/cimwin32prov/win32-logicaldisk' },
        @{ Term = 'Time Offset (s)'; Meaning = 'NTP delta via w32tm /stripchart; large drift breaks Kerberos.'; Link = 'https://learn.microsoft.com/windows-server/networking/windows-time-service/windows-time-service-tools-and-settings' },
        @{ Term = 'DNS Service'; Meaning = 'Microsoft DNS Server service status on the DC.'; Link = '' },
        @{ Term = 'NTDS Service (AD DS)'; Meaning = 'Active Directory Domain Services service status.'; Link = 'https://learn.microsoft.com/windows-server/identity/ad-ds/get-started/virtual-dc/active-directory-domain-services-overview' },
        @{ Term = 'NetLogon Service'; Meaning = 'DC locator/secure channel; required for logons.'; Link = 'https://learn.microsoft.com/windows-server/identity/ad-ds/manage/dc-locator' },

        # DCDIAG overview
        @{ Term = 'DCDIAG (all tests)'; Meaning = 'Built-in diagnostic suite for DC health.'; Link = 'https://learn.microsoft.com/windows-server/administration/windows-commands/dcdiag' },

        # DCDIAG one-liners
        @{ Term = 'Connectivity'; Meaning = 'Basic RPC/DNS/LDAP reachability to the DC.'; Link = 'https://learn.microsoft.com/windows-server/administration/windows-commands/dcdiag' },
        @{ Term = 'Advertising'; Meaning = 'DC advertises SRV records so clients can find it.'; Link = 'https://learn.microsoft.com/windows-server/identity/ad-ds/manage/dc-locator' },
        @{ Term = 'FrsEvent'; Meaning = 'Legacy FRS replication events (deprecated).'; Link = 'https://learn.microsoft.com/windows-server/storage/dfs-replication/migrate-sysvol-to-dfsr' },
        @{ Term = 'DFSREvent'; Meaning = 'DFS Replication events (SYSVOL/content replication).'; Link = 'https://learn.microsoft.com/windows-server/storage/dfs-replication/dfs-replication-overview' },
        @{ Term = 'SysVolCheck'; Meaning = 'SYSVOL is shared and healthy (GPOs/scripts available).'; Link = 'https://learn.microsoft.com/windows-server/storage/dfs-replication/migrate-sysvol-to-dfsr' },
        @{ Term = 'KccEvent'; Meaning = 'KCC topology generation health for replication.'; Link = 'https://learn.microsoft.com/openspecs/windows_protocols/ms-adts/f2e2f6c7-b232-406d-b48a-fc6ccf231202' },
        @{ Term = 'FSMO KnowsOfRoleHolders'; Meaning = 'Directory knows current FSMO owners.'; Link = 'https://learn.microsoft.com/troubleshoot/windows-server/active-directory/fsmo-roles' },
        @{ Term = 'MachineAccount'; Meaning = 'DC computer account integrity in AD.'; Link = 'https://learn.microsoft.com/windows-server/administration/windows-commands/dcdiag' },
        @{ Term = 'NCSecDesc'; Meaning = 'Security descriptors on directory partitions are valid.'; Link = 'https://learn.microsoft.com/windows-server/administration/windows-commands/dcdiag' },
        @{ Term = 'NetLogons'; Meaning = 'NetLogon functioning for logons and DC locator.'; Link = 'https://learn.microsoft.com/windows-server/identity/ad-ds/manage/dc-locator' },
        @{ Term = 'ObjectsReplicated / Replications'; Meaning = 'Inbound/outbound AD replication succeeding.'; Link = 'https://learn.microsoft.com/windows-server/identity/ad-ds/get-started/replication/active-directory-replication-concepts' },
        @{ Term = 'RidManager'; Meaning = 'RID pool allocation health (unique SIDs).'; Link = 'https://learn.microsoft.com/troubleshoot/windows-server/active-directory/fsmo-roles' },
        @{ Term = 'Services'; Meaning = 'Checks for DC-related services state via dcdiag.'; Link = 'https://learn.microsoft.com/windows-server/administration/windows-commands/dcdiag' },
        @{ Term = 'SystemLog'; Meaning = 'Scans the System event log for critical DC issues.'; Link = 'https://learn.microsoft.com/windows-server/administration/windows-commands/dcdiag' },
        @{ Term = 'VerifyReferences / CheckSDRefDom / CrossRefValidation'; Meaning = 'Directory reference and cross-ref integrity checks.'; Link = 'https://learn.microsoft.com/windows-server/administration/windows-commands/dcdiag' },
        @{ Term = 'LocatorCheck'; Meaning = 'Clients can locate this DC (DC Locator works).' ; Link = 'https://learn.microsoft.com/windows-server/identity/ad-ds/manage/dc-locator' },
        @{ Term = 'Intersite'; Meaning = 'Inter-site replication health.'; Link = 'https://learn.microsoft.com/windows-server/identity/ad-ds/get-started/replication/active-directory-replication-concepts' },
        @{ Term = 'FSMO Check'; Meaning = 'Can reach/validate FSMO role owners.'; Link = 'https://learn.microsoft.com/troubleshoot/windows-server/active-directory/fsmo-roles' }
    )

    # Render glossary
    [void]$sb.AppendLine('<tr><td style="padding-top:12px;">')
    [void]$sb.AppendLine('<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border:1px solid #e5e7eb;border-radius:10px;">')
    [void]$sb.AppendLine('<tr><td style="padding:12px;"><div style="font-weight:700;color:#0f172a;font-size:16px;margin-bottom:6px;">What each check means</div>')
    [void]$sb.AppendLine('<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">')
    foreach ($g in $glossary) {
        $term = HtmlEncode $g.Term
        $desc = HtmlEncode $g.Meaning
        $lnk  = [string]$g.Link
        $linkHtml = ''
        if ($lnk) {
            $lnkEnc = HtmlEncode $lnk
            $linkHtml = " <a href=""$lnkEnc"" style=""color:#0ea5e9;text-decoration:none;border-bottom:1px dotted #0ea5e9;"">Docs</a>"
        }
        [void]$sb.AppendLine("<tr><td style='vertical-align:top;padding:4px 8px 4px 0;font-weight:600;color:#0f172a;width:220px;'>$term</td><td style='vertical-align:top;padding:4px 0 4px 0;color:#334155;'>$desc$linkHtml</td></tr>")
    }
    [void]$sb.AppendLine('</table></td></tr></table>')
    [void]$sb.AppendLine('</td></tr>')

    # Footer + close
    [void]$sb.AppendLine(@"
      <tr><td style="padding-top:12px;font-size:12px;color:#64748b;">Generated for forest <b>$(HtmlEncode $Forest)</b> on $(HtmlEncode ($Generated.ToString('yyyy-MM-dd HH:mm:ss')))</td></tr>
    </table>
  </td></tr>
</table>
</body></html>
"@)

    $sb.ToString()
}

#endregion HTML

#region Main (fetch → build → save → email)
# ------------------------------------------

# Determine scope
if (-not $DomainName) {
    Write-Host "No domain specified: scanning all domains in forest" -ForegroundColor Yellow
    $domains = Get-AllDomains
} else {
    $domains = $DomainName
}

# Collect
foreach ($domain in $domains) {
    Write-Host "Scanning domain: $domain" -ForegroundColor Cyan
    $dcs = Get-AllDomainControllers -ComputerName $domain
    $total = $dcs.Count
    $i = 0

    foreach ($dc in $dcs) {
        $i++
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $name = $dc.HostName

        Write-Host ("  [{0}/{1}] {2}" -f $i, $total, $name) -ForegroundColor Gray

        $dcdiag = Get-DomainControllerDCDiagTestResults -ComputerName $name
        $svc    = Get-DomainControllerServices -ComputerName $name

        $row = [ordered]@{
            Server                      = ($name).ToLower()
            Domain                      = $domain                       # Required for grouping
            Site                        = $dc.Site
            'OS Version'                = $dc.OperatingSystem
            'IPv4 Address'              = $dc.IPv4Address
            'Operation Master Roles'    = $dc.OperationMasterRoles
            'DNS'                       = Get-DomainControllerNSLookup -ComputerName $name
            'Ping'                      = Get-DomainControllerPingStatus -ComputerName $name
            'Uptime (hours)'            = Get-DomainControllerUpTime -ComputerName $name
            'OS Free Space (%)'         = Get-DomainControllerOSDriveFreeSpace -ComputerName $name
            'OS Free Space (GB)'        = Get-DomainControllerOSDriveFreeSpaceGB -ComputerName $name
            'Time offset (seconds)'     = Get-TimeDifference -ComputerName $name
            'DNS Service'               = $svc.DNSService
            'NTDS Service'              = $svc.NTDSService
            'NetLogon Service'          = $svc.NETLOGONService
            'DCDIAG: Connectivity'      = $dcdiag.Connectivity
            'DCDIAG: Advertising'       = $dcdiag.Advertising
            'DCDIAG: FrsEvent'          = $dcdiag.FrsEvent
            'DCDIAG: DFSREvent'         = $dcdiag.DFSREvent
            'DCDIAG: SysVolCheck'       = $dcdiag.SysVolCheck
            'DCDIAG: KccEvent'          = $dcdiag.KccEvent
            'DCDIAG: FSMO KnowsOfRoleHolders' = $dcdiag.KnowsOfRoleHolders
            'DCDIAG: MachineAccount'    = $dcdiag.MachineAccount
            'DCDIAG: NCSecDesc'         = $dcdiag.NCSecDesc
            'DCDIAG: NetLogons'         = $dcdiag.NetLogons
            'DCDIAG: ObjectsReplicated' = $dcdiag.ObjectsReplicated
            'DCDIAG: Replications'      = $dcdiag.Replications
            'DCDIAG: RidManager'        = $dcdiag.RidManager
            'DCDIAG: Services'          = $dcdiag.Services
            'DCDIAG: SystemLog'         = $dcdiag.SystemLog
            'DCDIAG: VerifyReferences'  = $dcdiag.VerifyReferences
            'DCDIAG: CheckSDRefDom'     = $dcdiag.CheckSDRefDom
            'DCDIAG: CrossRefValidation'= $dcdiag.CrossRefValidation
            'DCDIAG: LocatorCheck'      = $dcdiag.LocatorCheck
            'DCDIAG: Intersite'         = $dcdiag.Intersite
            'DCDIAG: FSMO Check'        = $dcdiag.FSMOCheck
            'Processing Time (seconds)' = $sw.Elapsed.TotalSeconds.ToString('0')
        }

        $allTestedDomainControllers.Add($row)
    }
}

# Build HTML (use the exact email layout for the file too)
$forestName  = (Get-ADForest).Name
$reportTitle = "Domain Controller Health – $forestName"
$reportDesc  = "Email-style layout (inline CSS) identical to the email body."
$htmlreport  = Build-EmailSafeHtmlReport -Items $allTestedDomainControllers `
    -Forest $forestName -Generated $reportTime -Title $reportTitle -Description $reportDesc

# Save
if (-not (Test-Path $ReportPath)) { New-Item -Path $ReportPath -ItemType Directory -Force | Out-Null }
$reportFile = Join-Path $ReportPath ("ADHealthReport_{0}.html" -f $reportFileNameTime)
$htmlreport | Out-File -FilePath $reportFile -Encoding UTF8
Write-Host "✔ Report saved to: $reportFile" -ForegroundColor Green

# Email

$emailHtml = Build-EmailSafeHtmlReport -Items $allTestedDomainControllers `
    -Forest $forestName -Generated $reportTime `
    -Title "Domain Controller Health – $forestName" `
    -Description " Performs a comprehensive health check of all domain controllers in a specified domain
    or across the entire Active Directory forest. Executed via a scheduled task on DC03."


    if (-not $Credential) {
        Write-Error "SendEmail requested but no Credential provided (or -Password to build one)."
    } else {
        try {
            Send-MailMessage -From $UserFrom -To $UserTo -Subject $Subject `
              -Body $emailHtml -BodyAsHtml -SmtpServer $SmtpServer -Port $Port -UseSsl `
              -Credential $Credential -Encoding ([System.Text.Encoding]::UTF8)
            Write-Host "✔ Email sent." -ForegroundColor Green
        } catch {
            Write-Error "Failed to send email: $($_.Exception.Message)"
        }
    }


#endregion Main
