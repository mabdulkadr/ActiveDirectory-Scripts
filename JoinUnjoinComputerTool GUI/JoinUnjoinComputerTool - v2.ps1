<#
.SYNOPSIS
    Join/Unjoin Computer Tool – WPF GUI with Message Center & Credential/DC Status (PowerShell 5.1)

.DESCRIPTION
    Compact WPF tool to:
      • Set/validate AD credentials (button in Domain Configuration)
      • Show saved Username + Cred Status (Connected/Not Connected/Error)
      • Pick OU and join the computer to a domain
      • Disjoin to WORKGROUP
      • Delete computer from AD (only if disabled)
      • Show live PC info (Name, IP, Domain, DC Status, Entra ID, SCCM, Co-Management)
      • Message Center: color-coded log (INFO/SUCCESS/WARNING/ERROR), Copy All, Clear

.PARAMETERS
    -DefaultDomainController  FQDN of the Domain Controller
    -DefaultDomainName        FQDN of the Domain
    -DefaultSearchBase        Default OU DN for joins

.NOTES
    Author  : Mohammad Abdulkader Omar
    Website : https://momar.tech
    Date    : 2025-01-20
    Version : 2.5  (Fixed PS5.1 ternary usage + all braces; added credentials/DC status)
#>

[CmdletBinding()]
param(
    [string]$DefaultDomainController = "dc01.example.local",
    [string]$DefaultDomainName       = "example.local",
    [string]$DefaultSearchBase       = "OU=Domain Computers,DC=example,DC=local"
)

# -------------------- Assemblies & Globals --------------------
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms
Add-Type -AssemblyName System.DirectoryServices.AccountManagement

$script:ADCreds = $null
$script:LastCredError = $null
$script:PCInfoJob = $null
$script:PCInfoTimer = $null
$script:LastResult = $null
$script:LastCredValidation = $null

function New-Brush {
    param([string]$Hex = "#000000")
    $c = [Windows.Media.ColorConverter]::ConvertFromString($Hex)
    $b = New-Object Windows.Media.SolidColorBrush $c
    $b.Freeze()
    return $b
}

# Status pill brushes
$script:BrushOK = New-Brush "#28A745"
$script:BrushWarn = New-Brush "#FFA500"
$script:BrushBad = New-Brush "#DC3545"
$script:BrushWhite = [Windows.Media.Brushes]::White

# Message Center palette
$script:BrushInfoUI = New-Brush "#374151"  # dark gray
$script:BrushSuccessUI = New-Brush "#0A8A0A"  # green
$script:BrushWarnUI = New-Brush "#B58900"  # amber
$script:BrushErrorUI = New-Brush "#D13438"  # red

# -------------------- Dialogs & Helpers --------------------
function Show-WPFMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [string]$Title = "Message",
        [ValidateSet("Green","Orange","Red","Blue")] [string]$Color = "Blue",

        # New: sizing controls (defaults tuned for your UI)
        [int]$MinWidth   = 420,
        [int]$MaxWidth   = 820,
        [int]$MinHeight  = 160,
        [int]$MaxHeight = 600,
        [int]$MaxMessageHeight = 360   # scroll region height cap
    )

    switch ($Color) {
        "Green"  { $HeaderColor = "#28A745" }
        "Orange" { $HeaderColor = "#F59E0B" }
        "Red"    { $HeaderColor = "#DC3545" }
        "Blue"   { $HeaderColor = "#3B82F6" }
    }

    [xml]$xaml = @"
<Window xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
        xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
        ResizeMode='CanMinimize' WindowStartupLocation='CenterScreen'
        SizeToContent='WidthAndHeight' FontFamily='Segoe UI'>
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height='Auto'/>
      <RowDefinition Height='*'/>
      <RowDefinition Height='Auto'/>
    </Grid.RowDefinitions>

    <!-- Header uses requested color -->
    <Border Grid.Row='0' Padding='10' Background='$HeaderColor'>
      <TextBlock x:Name='hdr' Foreground='White' FontSize='16' FontWeight='Bold'
                 HorizontalAlignment='Center'/>
    </Border>

    <!-- Scrollable message area -->
    <ScrollViewer Grid.Row='1' x:Name='sv'
                  VerticalScrollBarVisibility='Auto'
                  HorizontalScrollBarVisibility='Auto'
                  Margin='14,12,14,6'>
      <TextBlock x:Name='txt'
                 TextWrapping='Wrap'
                 FontSize='13'
                 Foreground='#111827'/>
    </ScrollViewer>

    <StackPanel Grid.Row='2' Orientation='Horizontal' HorizontalAlignment='Center' Margin='0,8,0,12'>
      <Button x:Name='ok' Content='OK' Width='100' Height='30' Margin='6'
              Background='$HeaderColor' Foreground='White' BorderThickness='0'
              FontWeight='SemiBold' Cursor='Hand'/>
    </StackPanel>
  </Grid>
</Window>
"@

    $w = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))

    # Apply content + sizing safely at runtime
    $w.Title = $Title
    $w.MinWidth  = $MinWidth;  $w.MaxWidth  = $MaxWidth
    $w.MinHeight = $MinHeight; $w.MaxHeight = $MaxHeight
    $w.FindName('hdr').Text = $Title
    $w.FindName('txt').Text = $Message
    $w.FindName('sv').MaxHeight = $MaxMessageHeight

    $w.FindName('ok').Add_Click({ $w.Close() })
    [void]$w.ShowDialog()
}
function Show-WPFConfirmation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [string]$Title = "Confirm",
        [ValidateSet("Green","Orange","Red","Blue")] [string]$Color = "Orange",

        # New: sizing controls
        [int]$MinWidth = 420,
        [int]$MaxWidth = 820,
        [int]$MinHeight = 160,
        [int]$MaxHeight = 600,
        [int]$MaxMessageHeight = 360
    )

    switch ($Color) {
        "Green"  { $HeaderColor = "#28A745" }
        "Orange" { $HeaderColor = "#F59E0B" }
        "Red"    { $HeaderColor = "#DC3545" }
        "Blue"   { $HeaderColor = "#3B82F6" }
    }

    [xml]$xaml = @"
<Window xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
        xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
        ResizeMode='CanMinimize' WindowStartupLocation='CenterScreen'
        SizeToContent='WidthAndHeight' FontFamily='Segoe UI'>
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height='Auto'/>
      <RowDefinition Height='*'/>
      <RowDefinition Height='Auto'/>
    </Grid.RowDefinitions>

    <Border Grid.Row='0' Padding='10' Background='$HeaderColor'>
      <TextBlock x:Name='hdr' Foreground='White' FontSize='16' FontWeight='Bold'
                 HorizontalAlignment='Center'/>
    </Border>

    <!-- Scrollable message area -->
    <ScrollViewer Grid.Row='1' x:Name='sv'
                  VerticalScrollBarVisibility='Auto'
                  HorizontalScrollBarVisibility='Auto'
                  Margin='14,12,14,6'>
      <TextBlock x:Name='txt'
                 TextWrapping='Wrap'
                 FontSize='13'
                 Foreground='#111827'/>
    </ScrollViewer>

    <StackPanel Grid.Row='2' Orientation='Horizontal' HorizontalAlignment='Center' Margin='0,8,0,12'>
      <Button x:Name='yes' Content='Yes' Width='110' Height='30' Margin='6'
              Background='$HeaderColor' Foreground='White' BorderThickness='0' Cursor='Hand'/>
      <Button x:Name='no'  Content='No'  Width='110' Height='30' Margin='6'
              Background='#6B7280'  Foreground='White' BorderThickness='0' Cursor='Hand'/>
    </StackPanel>
  </Grid>
</Window>
"@

    $w = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))

    # Apply text + sizing after load (safe for quotes/special chars)
    $w.Title = $Title
    $w.MinWidth  = $MinWidth;  $w.MaxWidth  = $MaxWidth
    $w.MinHeight = $MinHeight; $w.MaxHeight = $MaxHeight
    $w.FindName('hdr').Text = $Title
    $w.FindName('txt').Text = $Message
    $w.FindName('sv').MaxHeight = $MaxMessageHeight

    # Correct return path using DialogResult
    $w.FindName('yes').Add_Click({ $w.DialogResult = $true  })
    $w.FindName('no'). Add_Click({ $w.DialogResult = $false })

    [void]$w.ShowDialog()
    return [bool]$w.DialogResult
}
function Test-Admin {
    $p = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Show-WPFMessage -Message "Run this script as Administrator." -Title "Insufficient Privileges" -Color Red
        exit
    }
}
function Convert-ADLargeInteger {
    param($Value)
    if ($null -eq $Value) { return $null }
    try {
        if ($Value -is [System.__ComObject]) {
            $high = $Value.HighPart
            $low  = $Value.LowPart
            $file = ([int64]$high -shl 32) -bor ($low -band 0xffffffff)
        } else {
            $file = [int64]$Value
        }
        if ($file -le 0) { return $null }
        return ([DateTime]::FromFileTimeUtc($file)).ToLocalTime()
    } catch { return $null }
}
function Get-UacEnabledString {
    param([int]$UserAccountControl)
    if ($UserAccountControl -band 2) { 'Disabled' } else { 'Enabled' }
}

# --- AD search helpers in script scope (visible to nested handlers) ---------
function Script:Escape-Ldap {
    param([string]$s)
    if ($null -eq $s) { return "" }
    $s = $s -replace '\\','\5c'
    $s = $s -replace '\*','\2a'
    $s = $s -replace '\(','\28'
    $s = $s -replace '\)','\29'
    $s = $s -replace '\x00','\00'
    return $s
}
function Script:Get-DSValue {
    param($Props, [string]$Name)
    if ($Props[$Name] -and $Props[$Name].Count -gt 0) { $Props[$Name][0] } else { $null }
}
function Script:Build-AdComputerFilter {
    param([string]$Term, [bool]$IncludeDisabled = $false)

    if ([string]::IsNullOrWhiteSpace($Term)) {
        $base = '(objectCategory=computer)'
    } else {
        $e = Script:Escape-Ldap $Term
        $base = "(&(objectCategory=computer)(|(name=*$e*)(dNSHostName=*$e*)(sAMAccountName=*$e*)))"
    }

    if ($IncludeDisabled) { 
        return $base
    } else {
        # exclude disabled (UAC bit 2)
        return "(&$base(!(userAccountControl:1.2.840.113556.1.4.803:=2)))"
    }
}

# -------------------- Message Center (UI log) --------------------
function Write-UiLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'SUCCESS', 'WARNING', 'ERROR')] [string]$Level = 'INFO'
    )
    if (-not $script:LogBox) { return }

    $brush = $script:BrushInfoUI
    if ($Level -eq 'SUCCESS') { $brush = $script:BrushSuccessUI }
    elseif ($Level -eq 'WARNING') { $brush = $script:BrushWarnUI }
    elseif ($Level -eq 'ERROR') { $brush = $script:BrushErrorUI }

    $script:LogBox.Dispatcher.Invoke([action] {
            $p = New-Object Windows.Documents.Paragraph
            $run = New-Object Windows.Documents.Run ("[{0}] {1}" -f $Level, $Message)
            $run.Foreground = $brush
            $p.Margin = New-Object System.Windows.Thickness 0, 0, 0, 2
            [void]$p.Inlines.Add($run)
            $script:LogBox.Document.Blocks.Add($p)
            $script:LogBox.ScrollToEnd()
        })
}

# -------------------- Credentials & Restart --------------------
function Test-AdCredentials {
    param(
        [Parameter(Mandatory)][pscredential]$Credential,
        [string]$DomainController,
        [string]$DomainName
    )
    try {
        # Choose the server to test against (prefer the DC textbox)
        $server = $null
        if ($DomainController) { $server = $DomainController.Trim() }
        elseif ($DomainName) { $server = $DomainName.Trim() }

        if (-not $server) {
            return @{ Ok = $false; Reason = "No DC/domain provided"; Server = "" }
        }

        # Reachability test (LDAP 389 / Kerberos 88)
        $reachable = $false
        foreach ($port in 389, 88) {
            try {
                $client = New-Object System.Net.Sockets.TcpClient
                $ar = $client.BeginConnect($server, $port, $null, $null)
                $ok = $ar.AsyncWaitHandle.WaitOne(1500, $false)
                if ($ok -and $client.Connected) {
                    $client.EndConnect($ar)
                    $reachable = $true
                }
                $client.Close()
            }
            catch { }
            if ($reachable) { break }
        }
        if (-not $reachable) {
            return @{ Ok = $false; Reason = "DC unreachable ($server)"; Server = $server }
        }

        # Validate credentials (Kerberos/NTLM)
        $plain = $Credential.GetNetworkCredential()
        $user = $plain.UserName
        $pass = $plain.Password

        $ctx = New-Object System.DirectoryServices.AccountManagement.PrincipalContext `
        ([System.DirectoryServices.AccountManagement.ContextType]::Domain, $server)

        $valid = $ctx.ValidateCredentials(
            $user, $pass,
            [System.DirectoryServices.AccountManagement.ContextOptions]::Negotiate
        )
        if (-not $valid) {
            return @{ Ok = $false; Reason = "Invalid username or password"; Server = $server }
        }

        # Optional LDAP bind to confirm directory access
        try {
            $ldap = "LDAP://$server"
            $de = New-Object DirectoryServices.DirectoryEntry($ldap, $Credential.UserName, $plain.Password)
            $null = $de.NativeObject
        }
        catch {
            return @{ Ok = $false; Reason = "Validated, but LDAP bind failed: $($_.Exception.Message)"; Server = $server }
        }

        return @{ Ok = $true; Reason = "Validated on $server"; Server = $server }
    }
    catch {
        # PS 5.1-safe fallback for Server value (no '??')
        $srv = ""
        if ($DomainController) { $srv = $DomainController }
        elseif ($DomainName) { $srv = $DomainName }
        return @{ Ok = $false; Reason = $_.Exception.Message; Server = $srv }
    }
}
function Prompt-Credentials {
    if ($script:ADCreds -and $script:ADCreds.UserName) { return $script:ADCreds }

    [xml]$xaml = @"
<Window xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
        xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
        ResizeMode='NoResize' WindowStartupLocation='CenterScreen'
        SizeToContent='WidthAndHeight' FontFamily='Segoe UI'>
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height='Auto'/>
      <RowDefinition Height='Auto'/>
      <RowDefinition Height='Auto'/>
    </Grid.RowDefinitions>

    <Border Grid.Row='0' Padding='10'>
      <Border.Background>
        <LinearGradientBrush StartPoint='0,0' EndPoint='1,0'>
          <GradientStop Color='#3B82F6' Offset='0'/>
          <GradientStop Color='#6366F1' Offset='1'/>
        </LinearGradientBrush>
      </Border.Background>
      <TextBlock Text='Enter AD Credentials' Foreground='White'
                 FontSize='16' FontWeight='Bold' HorizontalAlignment='Center'/>
    </Border>

    <StackPanel Grid.Row='1' Margin='18'>
      <StackPanel Orientation='Horizontal' Margin='0,6,0,6'>
        <Label Content='Username:' Width='110' VerticalAlignment='Center'/>
        <TextBox x:Name='u' Width='260'/>
      </StackPanel>
      <StackPanel Orientation='Horizontal' Margin='0,6,0,6'>
        <Label Content='Password:' Width='110' VerticalAlignment='Center'/>
        <PasswordBox x:Name='p' Width='260'/>
      </StackPanel>
    </StackPanel>

    <StackPanel Grid.Row='2' Orientation='Horizontal' HorizontalAlignment='Center' Margin='0,8,0,12'>
      <Button x:Name='ok'     Content='OK'     Width='110' Height='30' Margin='6'
              Background='#3B82F6' Foreground='White' BorderThickness='0' Cursor='Hand'/>
      <Button x:Name='cancel' Content='Cancel' Width='110' Height='30' Margin='6'
              Background='#6B7280' Foreground='White' BorderThickness='0' Cursor='Hand'/>
    </StackPanel>
  </Grid>
</Window>
"@

    $w = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
    $u = $w.FindName('u'); $p = $w.FindName('p')
    $ok = $w.FindName('ok'); $cancel = $w.FindName('cancel')

    $script:ADCreds = $null
    $script:LastCredError = $null
    $script:LastCredValidation = $null

    $ok.Add_Click({
            if (-not $u.Text -or -not $p.Password) {
                Show-WPFMessage -Message "Username and Password required." -Title "Info" -Color Blue
                return
            }
            try {
                $sec = ConvertTo-SecureString $p.Password -AsPlainText -Force
                $cred = New-Object System.Management.Automation.PSCredential ($u.Text, $sec)

                # DC & domain to test against
                $dcUI = $null; $domUI = $null
                try { if ($script:DomainControllerBox) { $dcUI = $script:DomainControllerBox.Text } } catch {}
                try { if ($script:DomainNameBox) { $domUI = $script:DomainNameBox.Text } } catch {}
                if (-not $dcUI) { $dcUI = $DefaultDomainController }
                if (-not $domUI) { $domUI = $DefaultDomainName }

                $res = Test-AdCredentials -Credential $cred -DomainController $dcUI -DomainName $domUI
                $script:LastCredValidation = $res

                if ($res.Ok) {
                    $script:ADCreds = $cred
                    $script:LastCredError = $null
                    Write-UiLog ("Credentials validated: {0}" -f $res.Reason) "SUCCESS"
                    $w.Close()
                }
                else {
                    $script:ADCreds = $null
                    $script:LastCredError = $res.Reason
                    Write-UiLog ("Credential validation failed: {0}" -f $res.Reason) "ERROR"
                    Show-WPFMessage -Message ("Credential validation failed.`n{0}" -f $res.Reason) -Title "Error" -Color Red
                }
            }
            catch {
                $script:ADCreds = $null
                $script:LastCredError = $_.Exception.Message
                $script:LastCredValidation = @{ Ok = $false; Reason = $script:LastCredError; Server = "" }
                Write-UiLog ("Credential validation failed: {0}" -f $script:LastCredError) "ERROR"
                Show-WPFMessage -Message ("Credential validation failed.`n{0}" -f $script:LastCredError) -Title "Error" -Color Red
            }
        })

    $cancel.Add_Click({ $script:ADCreds = $null; $w.Close() })

    [void]$w.ShowDialog()
    return $script:ADCreds
}
Function Prompt-Restart {
    [void][System.Reflection.Assembly]::LoadWithPartialName("presentationframework")
    [void][System.Reflection.Assembly]::LoadWithPartialName("windowsbase")

    [xml]$RestartXAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        ResizeMode="NoResize"
        WindowStartupLocation="CenterScreen"
        SizeToContent="WidthAndHeight">
  <Grid>
    <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header Section -->
        <Border Grid.Row='0' Background='#0078D7' Padding='5'>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                <TextBlock Text="Restart Confirmation" Foreground="White" FontSize="18" FontWeight="Bold" VerticalAlignment="Center"/>
            </StackPanel>
        </Border>

        <!-- Content Section -->
        <StackPanel Grid.Row="1" Margin="10" VerticalAlignment="Center" HorizontalAlignment="Center">
            <TextBlock Text="The computer needs to restart to complete the operation." FontSize="14" TextAlignment="Center" Margin="10"/>
            <TextBlock Text="Do you want to restart now?" FontSize="14" FontWeight="Bold" TextAlignment="Center" Margin="10"/>
        </StackPanel>

        <!-- Footer Section -->
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Center" Margin="10">
            <Button x:Name="YesButton" Content="Yes, Restart" Width="120" Height="30" Background="#32CD32" Foreground="White" FontWeight="Bold" Margin="5"/>
            <Button x:Name="NoButton" Content="No, Postpone" Width="120" Height="30" Background="#FF6347" Foreground="White" FontWeight="Bold" Margin="5"/>
        </StackPanel>
      </Grid>
  </Grid>
</Window>
"@

    try {
        $Reader = New-Object System.Xml.XmlNodeReader $RestartXAML
        $RestartWindow = [System.Windows.Markup.XamlReader]::Load($Reader)

        $YesButton = $RestartWindow.FindName('YesButton')
        $NoButton  = $RestartWindow.FindName('NoButton')

        $Global:RestartConfirmed = $false

        $YesButton.Add_Click({
            $Global:RestartConfirmed = $true
            $RestartWindow.Close()
        })
        $NoButton.Add_Click({
            $Global:RestartConfirmed = $false
            $RestartWindow.Close()
        })

        $RestartWindow.WindowStartupLocation = "CenterScreen"
        [void]$RestartWindow.ShowDialog()

        if ($Global:RestartConfirmed) {
            Show-WPFMessage -Message "Restarting the computer..." -Title "Warning" -Color Orange
            Restart-Computer -Force
        }
        else {
            Show-WPFMessage -Message "Restart postponed by the user." -Title "Info" -Color Blue
        }
    }
    catch {
        Show-WPFMessage -Message "An error occurred in the Restart Confirmation Window: $($_.Exception.Message)" -Title "Error" -Color Red
    }
}

# -------------------- PC Info (async) --------------------
function Update-PCInfo {
    if ($script:PCInfoJob -and ($script:PCInfoJob.State -in 'Completed', 'Failed', 'Stopped')) {
        Receive-Job -Job $script:PCInfoJob -ErrorAction SilentlyContinue | Out-Null
        Remove-Job  -Job $script:PCInfoJob -Force | Out-Null
        $script:PCInfoJob = $null
    }
    elseif ($script:PCInfoJob -and $script:PCInfoJob.State -eq 'Running') {
        Write-UiLog "A PC info update is already running." "INFO"
        return
    }

    Write-UiLog "Refreshing PC information..." "INFO"
    $dcName = $script:DomainControllerBox.Text.Trim()

    $script:PCInfoJob = Start-Job -ArgumentList $dcName -ScriptBlock {
        param($DCName)
        try {
            $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
            $ComputerName = $cs.Name
            $DomainStatus = if ($cs.PartOfDomain) { "Domain Joined" } else { "Workgroup" }

            $IPAddress = $null
            try {
                $nic = Get-NetAdapter -Physical | Where-Object Status -eq 'Up' | Sort-Object LinkSpeed -Descending | Select-Object -First 1
                if ($nic) {
                    $ip = Get-NetIPAddress -InterfaceIndex $nic.IfIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                    Where-Object { $_.IPAddress -notmatch '^127\.|^169\.254' } |
                    Select-Object -ExpandProperty IPAddress -First 1
                    if ($ip) { $IPAddress = $ip }
                }
                if (-not $IPAddress) {
                    $IPAddress = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                    Where-Object { $_.IPAddress -notmatch '^127\.|^169\.254' } |
                    Select-Object -ExpandProperty IPAddress -First 1
                }
            }
            catch { }

            # Entra status
            $ds = dsregcmd /status 2>$null | Out-String
            if ($ds -match "AzureAdJoined\s*:\s*YES" -and $ds -match "DomainJoined\s*:\s*YES") {
                $EntraIDStatus = "Hybrid AD Joined"
            }
            elseif ($ds -match "AzureAdJoined\s*:\s*YES") {
                $EntraIDStatus = "Azure AD Joined"
            }
            else {
                $EntraIDStatus = "Not Joined"
            }

            # SCCM
            $SCCMStatus = "Not Installed"; $SCCMColor = "#DC3545"
            $ccm = Get-Service ccmexec -ErrorAction SilentlyContinue
            if ($ccm) {
                if ($ccm.Status -eq 'Running') { $SCCMStatus = "SCCM Client (Running)"; $SCCMColor = "#28A745" }
                else { $SCCMStatus = "SCCM Client (Stopped)"; $SCCMColor = "#FFA500" }
            }

            # Co-management
            $CoStatus = "Not Detected"; $CoColor = "#DC3545"
            try {
                $cm = Get-WmiObject -Namespace 'root\ccm\CoManagementHandler' -Class 'CoManagement_Configuration' -ErrorAction Stop
                if ($cm -and $cm.Enable) { $CoStatus = "Co-Management Enabled"; $CoColor = "#28A745" }
                elseif ($cm) { $CoStatus = "SCCM Present - Co-Management Disabled"; $CoColor = "#DC3545" }
            }
            catch { }
            $reg = "HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\MDM"
            if (Test-Path $reg) {
                $v = Get-ItemProperty -Path $reg -Name "AutoEnrollMDM" -ErrorAction SilentlyContinue
                if ($v -and $v.AutoEnrollMDM -eq 1) { $CoStatus = "Intune Auto Enrollment Enabled"; $CoColor = "#28A745" }
            }

            # DC reachability (LDAP 389 or Kerberos 88)
            $DCStatus = "Unknown"; $DCColor = "#FFA500"
            try {
                $reachable = $false
                foreach ($port in 389, 88) {
                    try {
                        $client = New-Object System.Net.Sockets.TcpClient
                        $ar = $client.BeginConnect($DCName, $port, $null, $null)
                        $ok = $ar.AsyncWaitHandle.WaitOne(1500, $false)
                        if ($ok -and $client.Connected) { $client.EndConnect($ar); $client.Close(); $reachable = $true; break }
                        $client.Close()
                    }
                    catch { }
                }
                if ($reachable) { $DCStatus = "Online"; $DCColor = "#28A745" } else { $DCStatus = "Unreachable"; $DCColor = "#DC3545" }
            }
            catch { }

            [PSCustomObject]@{
                ComputerName       = $ComputerName
                DomainStatus       = $DomainStatus
                IPAddress          = $(if ($IPAddress) { $IPAddress } else { "N/A" })
                DCStatus           = $DCStatus
                DCColor            = $DCColor
                EntraIDStatus      = $EntraIDStatus
                SCCMStatus         = $SCCMStatus
                SCCMColor          = $SCCMColor
                CoManagementStatus = $CoStatus
                CoManagementColor  = $CoColor
            }
        }
        catch { throw $_ }
    }

    if ($script:PCInfoTimer) { $script:PCInfoTimer.Stop(); $script:PCInfoTimer = $null }

    $script:PCInfoTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:PCInfoTimer.Interval = [TimeSpan]::FromMilliseconds(450)

    $script:PCInfoTimer.Add_Tick({
            if (-not $script:PCInfoJob) { return }

            $data = Receive-Job -Job $script:PCInfoJob -Keep -ErrorAction SilentlyContinue
            if ($data) {
                $r = $data[-1]
                $script:LastResult = $r
                try {
                    $script:PcNameBlock.Dispatcher.Invoke([action] { $script:PcNameBlock.Text = $r.ComputerName })
                    $script:PcIPAddressBlock.Dispatcher.Invoke([action] { $script:PcIPAddressBlock.Text = $r.IPAddress })
                    $script:PcDomainStatusBlock.Dispatcher.Invoke([action] {
                            $script:PcDomainStatusBlock.Text = $r.DomainStatus
                            if ($r.DomainStatus -eq "Domain Joined") {
                                $script:PcDomainStatusBlock.Background = $script:BrushOK
                            }
                            else {
                                $script:PcDomainStatusBlock.Background = $script:BrushBad
                            }
                            $script:PcDomainStatusBlock.Foreground = $script:BrushWhite
                        })
                    $script:DcReachBlock.Dispatcher.Invoke([action] {
                            $script:DcReachBlock.Text = $r.DCStatus
                            $script:DcReachBlock.Background = New-Brush $r.DCColor
                            $script:DcReachBlock.Foreground = $script:BrushWhite
                        })
                    $script:PcEntraIDStatusBlock.Dispatcher.Invoke([action] {
                            $script:PcEntraIDStatusBlock.Text = $r.EntraIDStatus

                            if ($r.EntraIDStatus -eq "Not Joined") {
                                $script:PcEntraIDStatusBlock.Background = $script:BrushBad       # Red
                            }
                            elseif ($r.EntraIDStatus -like "*Joined*") {
                                $script:PcEntraIDStatusBlock.Background = $script:BrushOK        # Green (Hybrid/Azure AD Joined)
                            }
                            else {
                                $script:PcEntraIDStatusBlock.Background = $script:BrushWarn      # Orange (any unexpected state)
                            }
                            $script:PcEntraIDStatusBlock.Foreground = $script:BrushWhite
                        })

                    $script:SCCMStatusBlock.Dispatcher.Invoke([action] {
                            $script:SCCMStatusBlock.Text = $r.SCCMStatus
                            $script:SCCMStatusBlock.Background = New-Brush $r.SCCMColor
                            $script:SCCMStatusBlock.Foreground = $script:BrushWhite
                        })
                    $script:CoManagementBlock.Dispatcher.Invoke([action] {
                            $script:CoManagementBlock.Text = $r.CoManagementStatus
                            $script:CoManagementBlock.Background = New-Brush $r.CoManagementColor
                            $script:CoManagementBlock.Foreground = $script:BrushWhite
                        })
                }
                catch {
                    Write-UiLog ("Error updating UI: {0}" -f $_.Exception.Message) "ERROR"
                }
            }

            if ($script:PCInfoJob.State -in 'Completed', 'Failed', 'Stopped') {
                $script:PCInfoTimer.Stop()
                Receive-Job -Job $script:PCInfoJob -ErrorAction SilentlyContinue | Out-Null
                Remove-Job  -Job $script:PCInfoJob -Force | Out-Null
                $script:PCInfoJob = $null

                # PS 5.1-safe summary (no ternary)
                if ($script:LastResult) {
                    Write-UiLog -Message ("PC: {0} | IP: {1}" -f $script:LastResult.ComputerName, $script:LastResult.IPAddress) -Level INFO

                    if ($script:LastResult.DomainStatus -eq "Domain Joined") {
                        Write-UiLog -Message ("Domain: {0}" -f $script:LastResult.DomainStatus) -Level SUCCESS
                    }
                    else {
                        Write-UiLog -Message ("Domain: {0}" -f $script:LastResult.DomainStatus) -Level WARNING
                    }

                    if ($script:LastResult.DCStatus -eq "Online") {
                        Write-UiLog -Message ("DC: {0}" -f $script:LastResult.DCStatus) -Level SUCCESS
                    }
                    else {
                        Write-UiLog -Message ("DC: {0}" -f $script:LastResult.DCStatus) -Level ERROR
                    }

                    if ($script:LastResult.EntraIDStatus -eq "Not Joined") {
                        Write-UiLog -Message ("Entra ID: {0}" -f $script:LastResult.EntraIDStatus) -Level ERROR
                    }
                    elseif ($script:LastResult.EntraIDStatus -like "*Joined*") {
                        Write-UiLog -Message ("Entra ID: {0}" -f $script:LastResult.EntraIDStatus) -Level SUCCESS
                    }
                    else {
                        Write-UiLog -Message ("Entra ID: {0}" -f $script:LastResult.EntraIDStatus) -Level WARNING
                    }

                    if ($script:LastResult.SCCMStatus -like "*Running*") {
                        Write-UiLog -Message $script:LastResult.SCCMStatus -Level SUCCESS
                    }
                    elseif ($script:LastResult.SCCMStatus -like "*Not Installed*") {
                        Write-UiLog -Message $script:LastResult.SCCMStatus -Level ERROR
                    }
                    else {
                        Write-UiLog -Message $script:LastResult.SCCMStatus -Level WARNING
                    }

                    if ($script:LastResult.CoManagementStatus -match "Enabled|Auto Enrollment") {
                        Write-UiLog -Message $script:LastResult.CoManagementStatus -Level SUCCESS
                    }
                    else {
                        Write-UiLog -Message $script:LastResult.CoManagementStatus -Level INFO
                    }
                }
            }
        })  # <-- closes Add_Tick scriptblock

    $script:PCInfoTimer.Start()
} # <-- closes function Update-PCInfo

# -------------------- AD Operations --------------------
function Join-ComputerWithOU {
    param([string]$DomainName, [string]$OUPath)
    try {
        $curr = (Get-WmiObject Win32_ComputerSystem).Domain
        if ($curr -eq $DomainName) {
            Write-UiLog ("Already in domain '{0}'." -f $DomainName) "INFO"
            Show-WPFMessage -Message "This computer is already a member of '$DomainName'." -Title "Info" -Color Blue
            return
        }
        Write-UiLog ("Joining domain '{0}' with OU '{1}'..." -f $DomainName, $OUPath) "INFO"
        Add-Computer -DomainName $DomainName -OUPath $OUPath -Credential $script:ADCreds -Force
        Write-UiLog ("Successfully joined to domain '{0}' (OU='{1}'). Restart required." -f $DomainName, $OUPath) "SUCCESS"
        Show-WPFMessage -Message "Successfully joined to domain:`n  • Domain: $DomainName`n  • OU: $OUPath" -Title "Success" -Color Green
        Update-PCInfo
        Prompt-Restart
    }
    catch {
        if ($_.Exception.Message -match "server is not operational") {
            Write-UiLog "Cannot reach Active Directory. Check DC/Domain/OU/Creds." "ERROR"
            Show-WPFMessage -Message "Cannot reach Active Directory. Check DC/Domain/OU/Creds." -Title "Error" -Color Red
        }
        else {
            Write-UiLog ("Join failed: {0}" -f $_.Exception.Message) "ERROR"
            Show-WPFMessage -Message ("Failed to join domain:`n{0}" -f $_.Exception.Message) -Title "Error" -Color Red
        }
    }
}
function Disjoin-ComputerFromDomain {
    param([string]$ComputerName)
    $inDomain = (Get-WmiObject Win32_ComputerSystem).PartOfDomain
    if (-not $inDomain) {
        Write-UiLog ("'{0}' is not domain-joined." -f $ComputerName) "INFO"
        Show-WPFMessage -Message "Computer '$ComputerName' is not domain-joined." -Title "Info" -Color Blue
        return
    }
    $ok = Show-WPFConfirmation -Message "Disjoin '$ComputerName' from the domain?" -Title "Disjoin from Domain" -Color Orange
    if (-not $ok) { Write-UiLog "Disjoin canceled." "INFO"; return }
    try {
        Write-UiLog ("Disjoining '{0}' to WORKGROUP..." -f $ComputerName) "WARNING"
        Remove-Computer -WorkgroupName "WORKGROUP" -Force
        Write-UiLog ("Successfully disjoined '{0}' from domain. Restart required." -f $ComputerName) "SUCCESS"
        Show-WPFMessage -Message "Successfully disjoined '$ComputerName' from domain." -Title "Success" -Color Green
        Update-PCInfo
        Prompt-Restart
    }
    catch {
        if ($_.Exception.Message -match "server is not operational") {
            Write-UiLog "Cannot reach Active Directory. Check DC/Domain/OU." "ERROR"
            Show-WPFMessage -Message "Cannot reach Active Directory. Check DC/Domain/OU." -Title "Error" -Color Red
        }
        else {
            Write-UiLog ("Disjoin failed: {0}" -f $_.Exception.Message) "ERROR"
            Show-WPFMessage -Message ("Failed to disjoin:`n{0}" -f $_.Exception.Message) -Title "Error" -Color Red
        }
    }
}
function Delete-ComputerFromAD {
    param(
        [string]$ComputerName,
        [string]$DomainController,
        [string]$SearchBase
    )

    try {
        if (-not $script:ADCreds) {
            Write-UiLog "Credentials required. Prompting..." "INFO"
            $null = Prompt-Credentials
            if (-not $script:ADCreds) { return }
        }

        # Primary confirmation
        $ok = Show-WPFConfirmation -Message "Delete '$ComputerName' from Active Directory?" -Title "Delete from AD" -Color Orange
        if (-not $ok) { Write-UiLog "Deletion canceled." "INFO"; return }

        $ldap = "LDAP://$DomainController/$SearchBase"
        $de = New-Object System.DirectoryServices.DirectoryEntry(
            $ldap,
            $script:ADCreds.UserName,
            $script:ADCreds.GetNetworkCredential().Password
        )
        $ds = New-Object System.DirectoryServices.DirectorySearcher($de)
        $ds.Filter = "(sAMAccountName=$ComputerName`$)"
        $res = $ds.FindOne()

        if ($res) {
            $ce  = $res.GetDirectoryEntry()
            $uac = [int]$ce.Properties["userAccountControl"].Value
            $isDisabled = (($uac -band 2) -ne 0)

            # Extra warning if the account is enabled/active
            # Extra warning if the account is enabled/active
            if (-not $isDisabled) {
                $warn = "The computer account '$ComputerName' appears ENABLED (active)." + [Environment]::NewLine +
                        "Deleting it may break domain logon for the device if it’s still in use. Delete anyway?"
                $ok2 = Show-WPFConfirmation -Message $warn -Title "Delete enabled account?" -Color Red
                if (-not $ok2) {
                    Write-UiLog ("Delete aborted: '{0}' is enabled." -f $ComputerName) "INFO"
                    Show-WPFMessage -Message "Deletion aborted by user (account enabled)." -Title "Aborted" -Color Orange
                    return
                }
            }
            try {
                Write-UiLog ("Deleting '{0}' from AD..." -f $ComputerName) "WARNING"
                $ce.DeleteTree()
                $ce.CommitChanges()
                Write-UiLog ("Successfully deleted '{0}' from Active Directory." -f $ComputerName) "SUCCESS"
                Show-WPFMessage -Message "Successfully deleted '$ComputerName' from Active Directory." -Title "Success" -Color Green
            }
            catch {
                # Common cause: Protect object from accidental deletion (deny delete ACE) or insufficient rights
                if ($_.Exception.Message -match "Access is denied") {
                    Write-UiLog ("Access denied deleting '{0}'. Clear 'Protect object from accidental deletion' or adjust ACLs." -f $ComputerName) "ERROR"
                    Show-WPFMessage -Message "Access denied. The object may be protected from accidental deletion or you lack rights. Clear the protection (Object tab) or adjust ACLs, then retry." -Title "Access Denied" -Color Red
                }
                else {
                    Write-UiLog ("Failed to delete from AD: {0}" -f $_.Exception.Message) "ERROR"
                    Show-WPFMessage -Message ("Failed to delete from AD:`n{0}" -f $_.Exception.Message) -Title "Error" -Color Red
                }
            }
        }
        else {
            Write-UiLog ("Computer '{0}' not found in AD." -f $ComputerName) "INFO"
            Show-WPFMessage -Message "Computer not found in AD." -Title "Info" -Color Blue
        }
    }
    catch {
        if ($_.Exception.Message -match "server is not operational") {
            Write-UiLog "Cannot reach AD. Check DC/Domain/OU." "ERROR"
            Show-WPFMessage -Message "Cannot reach AD. Check DC/Domain/OU." -Title "Error" -Color Red
        }
        else {
            Write-UiLog ("Failed to delete from AD: {0}" -f $_.Exception.Message) "ERROR"
            Show-WPFMessage -Message ("Failed to delete from AD:`n{0}" -f $_.Exception.Message) -Title "Error" -Color Red
        }
    }
}

# -------------------- OU Picker --------------------
function Show-OUWindow {
    [xml]$xaml = @"
<Window xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
        xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
        Title='Select Organizational Unit'
        Width='720' Height='460' MinWidth='700' MinHeight='420'
        Background='#F8FAFC' WindowStartupLocation='CenterScreen'
        ResizeMode='CanMinimize' FontFamily='Segoe UI'>
  <Grid Margin='8'>
    <Grid.RowDefinitions><RowDefinition Height='Auto'/><RowDefinition Height='*'/><RowDefinition Height='Auto'/></Grid.RowDefinitions>

    <StackPanel Grid.Row='0' Orientation='Horizontal' Margin='4'>
      <Label Content='Search:' VerticalAlignment='Center'/>
      <TextBox x:Name='q' Width='320' Margin='6,0,0,0'/>
      <Button x:Name='find' Content='Search' Width='90' Margin='8,0,0,0' Background='#3B82F6' Foreground='White' BorderThickness='0' Cursor='Hand'/>
    </StackPanel>

    <DataGrid x:Name='grid' Grid.Row='1' Margin='4' AutoGenerateColumns='False' CanUserAddRows='False' IsReadOnly='True'>
      <DataGrid.Effect><DropShadowEffect Color='#000000' Opacity='0.08' ShadowDepth='1' BlurRadius='4'/></DataGrid.Effect>
      <DataGrid.Columns>
        <DataGridTextColumn Header='Name' Binding='{Binding Name}' Width='160'/>
        <DataGridTextColumn Header='Description' Binding='{Binding Description}' Width='240'/>
        <DataGridTextColumn Header='Distinguished Name' Binding='{Binding DistinguishedName}' Width='*'/>
      </DataGrid.Columns>
    </DataGrid>

    <StackPanel Grid.Row='2' Orientation='Horizontal' HorizontalAlignment='Right' Margin='4,6,4,2'>
      <Button x:Name='ok'     Content='Select' Width='110' Height='30' Margin='6' Background='#10B981' Foreground='White' BorderThickness='0' Cursor='Hand'/>
      <Button x:Name='cancel' Content='Cancel' Width='110' Height='30' Margin='6' Background='#6B7280'  Foreground='White' BorderThickness='0' Cursor='Hand'/>
    </StackPanel>
  </Grid>
</Window>
"@
    $w = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
    $grid = $w.FindName('grid'); $q = $w.FindName('q'); $find = $w.FindName('find')
    $ok = $w.FindName('ok'); $cancel = $w.FindName('cancel')

    $all = @()
    try {
        $dc = $script:DomainControllerBox.Text.Trim()
        $sb = $script:SearchBaseBox.Text.Trim()
        $ldap = "LDAP://$dc/$sb"

        $de = New-Object DirectoryServices.DirectoryEntry($ldap, $script:ADCreds.UserName, $script:ADCreds.GetNetworkCredential().Password)
        $ds = New-Object DirectoryServices.DirectorySearcher($de)
        $ds.Filter = "(objectClass=organizationalUnit)"
        [void]$ds.PropertiesToLoad.Add("name")
        [void]$ds.PropertiesToLoad.Add("description")
        [void]$ds.PropertiesToLoad.Add("distinguishedName")
        $rs = $ds.FindAll()

        foreach ($r in $rs) {
            $all += [PSCustomObject]@{
                Name              = $r.Properties["name"][0]
                Description       = (($r.Properties["description"] -join ", ") -replace "^$", "N/A")
                DistinguishedName = $r.Properties["distinguishedName"][0]
            }
        }
        $grid.ItemsSource = $all
    }
    catch {
        if ($_.Exception.Message -match "server is not operational") {
            Show-WPFMessage -Message "Failed to connect to Active Directory. Check DC/Domain/OU." -Title "Error" -Color Red
        }
        else {
            Show-WPFMessage -Message ("Failed to retrieve OUs:`n{0}" -f $_.Exception.Message) -Title "Error" -Color Red
        }
    }

    $find.Add_Click({
            $t = $q.Text.Trim()
            if ($t -eq "") { $grid.ItemsSource = $all }
            else { $grid.ItemsSource = @($all | Where-Object { $_.Name -match $t -or $_.Description -match $t -or $_.DistinguishedName -match $t }) }
        })

    $ok.Add_Click({
            $sel = $grid.SelectedItem
            if ($sel -ne $null) {
                $script:SelectedOUBox.Text = $sel.DistinguishedName
                Write-UiLog ("Selected OU: {0}" -f $sel.DistinguishedName) "INFO"
                $w.Close()
            }
            else {
                Show-WPFMessage -Message "Select an OU first." -Title "Error" -Color Red
            }
        })
    $cancel.Add_Click({ $w.Close() })

    [void]$w.ShowDialog()
}
function Show-FindInADWindow {

    if (-not $script:ADCreds) { $null = Prompt-Credentials }
    if (-not $script:ADCreds) {
        Show-WPFMessage -Message "Credentials are required to search Active Directory." -Title "Find in AD" -Color Red
        return
    }

    # ---------- helpers ----------
    if (-not (Get-Command 'Script:Escape-Ldap' -ErrorAction SilentlyContinue)) {
        function Script:Escape-Ldap {
            param([string]$s)
            if ($null -eq $s) { return "" }
            $s = $s -replace '\\','\5c'
            $s = $s -replace '\*','\2a'
            $s = $s -replace '\(','\28'
            $s = $s -replace '\)','\29'
            $s = $s -replace '\x00','\00'
            return $s
        }
    }
    if (-not (Get-Command 'Script:Get-DSValue' -ErrorAction SilentlyContinue)) {
        function Script:Get-DSValue { param($Props,[string]$Name)
            if ($Props[$Name] -and $Props[$Name].Count -gt 0) { $Props[$Name][0] } else { $null }
        }
    }
    if (-not (Get-Command 'Script:Build-AdComputerFilter' -ErrorAction SilentlyContinue)) {
        function Script:Build-AdComputerFilter {
            param([string]$Term,[bool]$IncludeDisabled=$false)
            if ([string]::IsNullOrWhiteSpace($Term)) {
                $base='(objectCategory=computer)'
            } else {
                $e = Script:Escape-Ldap $Term
                $base = "(&(objectCategory=computer)(|(name=*$e*)(dNSHostName=*$e*)(sAMAccountName=*$e*)))"
            }
            if ($IncludeDisabled) { $base } else { "(&$base(!(userAccountControl:1.2.840.113556.1.4.803:=2)))" }
        }
    }
    if (-not (Get-Command 'Convert-ADLargeInteger' -ErrorAction SilentlyContinue)) {
        function Convert-ADLargeInteger {
            param($Value)
            if ($null -eq $Value) { return $null }
            try {
                if ($Value -is [System.__ComObject]) {
                    $high=$Value.HighPart; $low=$Value.LowPart
                    $file=([int64]$high -shl 32) -bor ($low -band 0xffffffff)
                } else { $file=[int64]$Value }
                if ($file -le 0) { return $null }
                ([DateTime]::FromFileTimeUtc($file)).ToLocalTime()
            } catch { $null }
        }
    }
    if (-not (Get-Command 'Get-UacEnabledString' -ErrorAction SilentlyContinue)) {
        function Get-UacEnabledString { param([int]$UserAccountControl)
            if ($UserAccountControl -band 2) { 'Disabled' } else { 'Enabled' }
        }
    }
    if (-not (Get-Command 'Script:Get-OUFromDN' -ErrorAction SilentlyContinue)) {
        function Script:Get-OUFromDN {
            param([string]$DN)
            if ([string]::IsNullOrWhiteSpace($DN)) { return $null }
            ($DN -split ',' | Where-Object { $_ -like 'OU=*' }) -join ','
        }
    }
    # Quick, non-blocking-ish logged-on user resolver (2s WSMan timeout)
    if (-not (Get-Command 'Script:Get-LoggedOnUserQuick' -ErrorAction SilentlyContinue)) {
        function Script:Get-LoggedOnUserQuick {
            param([string]$Computer,[pscredential]$Credential)
            try {
                # Prefer FQDN for sessions
                $target = $Computer
                $opt = New-CimSessionOption -Protocol WSMan -OperationTimeoutSec 2
                if ($Credential) {
                    $sess = New-CimSession -ComputerName $target -Credential $Credential -SessionOption $opt -ErrorAction Stop
                } else {
                    $sess = New-CimSession -ComputerName $target -SessionOption $opt -ErrorAction Stop
                }
                $cs = Get-CimInstance -CimSession $sess -ClassName Win32_ComputerSystem -ErrorAction Stop
                $sess | Remove-CimSession -ErrorAction SilentlyContinue
                return $cs.UserName
            } catch { return $null }
        }
    }

    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue

    # ----------------------------- XAML --------------------------------------
    [xml]$xaml = @"
<Window xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
        xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
        Title='Find Computers in Active Directory'
        Width='950' Height='560' MinWidth='900' MinHeight='520'
        WindowStartupLocation='CenterScreen' FontFamily='Segoe UI' Background='#F8FAFC'>
  <Grid Margin='8'>
    <Grid.RowDefinitions>
      <RowDefinition Height='Auto'/>
      <RowDefinition Height='*'/>
      <RowDefinition Height='Auto'/>
      <RowDefinition Height='Auto'/>
    </Grid.RowDefinitions>

    <!-- Search bar -->
    <Border Grid.Row='0' Padding='10' CornerRadius='10' Background='White' BorderBrush='#E5E7EB' BorderThickness='1'>
      <DockPanel LastChildFill='True'>
        <StackPanel Orientation='Horizontal' DockPanel.Dock='Left'>
          <Label Content='Search:' VerticalAlignment='Center'/>
          <TextBox x:Name='q' Width='360' Margin='6,0,0,0' ToolTip='Name, DNSHostName, or sAMAccountName'/>
          <CheckBox x:Name='chkDisabled' Content='Include disabled' Margin='12,0,0,0' VerticalAlignment='Center' IsChecked='True'/>
          <Button x:Name='btnSearch'  Content='Search'  Width='100' Height='30' Margin='12,0,0,0'/>
          <Button x:Name='btnClear'   Content='Clear'   Width='90'  Height='30' Margin='6,0,0,0'/>
        </StackPanel>
      </DockPanel>
    </Border>

    <!-- Results -->
    <Border Grid.Row='1' Margin='0,8,0,8' Padding='6' CornerRadius='10' Background='White' BorderBrush='#E5E7EB' BorderThickness='1'>
      <DataGrid x:Name='grid' AutoGenerateColumns='False' CanUserAddRows='False' IsReadOnly='True'
                AlternationCount='2' SelectionMode='Single' SelectionUnit='FullRow'
                VerticalScrollBarVisibility='Auto' HorizontalScrollBarVisibility='Auto'>
        <DataGrid.Columns>
          <DataGridTextColumn Header='Name'            Binding='{Binding Name}'             Width='160'/>
          <DataGridTextColumn Header='DNS Hostname'    Binding='{Binding DNSHostName}'      Width='200'/>
          <DataGridTextColumn Header='Enabled'         Binding='{Binding Enabled}'          Width='80'/>
          <DataGridTextColumn Header='OS'              Binding='{Binding OperatingSystem}'  Width='180'/>
          <DataGridTextColumn Header='Last Logon'      Binding='{Binding LastLogon}'        Width='160'/>
          <DataGridTextColumn Header='Pwd Last Set'    Binding='{Binding PwdLastSet}'       Width='150'/>
          <DataGridTextColumn Header='When Created'    Binding='{Binding WhenCreated}'      Width='150'/>
          <DataGridTextColumn Header='Distinguished Name' Binding='{Binding DistinguishedName}' Width='*'/>
        </DataGrid.Columns>
      </DataGrid>
    </Border>

    <!-- Details: TWO COLUMNS -->
    <Border Grid.Row='2' Padding='8' CornerRadius='10' Background='White' BorderBrush='#E5E7EB' BorderThickness='1'>
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width='*'/>
          <ColumnDefinition Width='*'/>
        </Grid.ColumnDefinitions>

        <!-- Left column -->
        <TextBox x:Name='txtLeft' Grid.Column='0' IsReadOnly='True' TextWrapping='Wrap' FontFamily='Consolas' FontSize='12'
                 Height='120' VerticalScrollBarVisibility='Auto' Background='#F9FAFB' Margin='0,0,6,0'/>

        <!-- Right column -->
        <TextBox x:Name='txtRight' Grid.Column='1' IsReadOnly='True' TextWrapping='Wrap' FontFamily='Consolas' FontSize='12'
                 Height='120' VerticalScrollBarVisibility='Auto' Background='#F9FAFB' Margin='6,0,0,0'/>
      </Grid>
    </Border>

    <!-- Footer / status -->
    <DockPanel Grid.Row='3' Margin='0,6,0,0'>
      <TextBlock x:Name='lblStatus' Text='Ready' Foreground='#374151' />
    </DockPanel>
  </Grid>
</Window>
"@

    # ------------------------- window & controls -----------------------------
    $w          = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
    $q          = $w.FindName('q')
    $chkDisabled= $w.FindName('chkDisabled')
    $btnSearch  = $w.FindName('btnSearch')
    $btnClear   = $w.FindName('btnClear')
    $grid       = $w.FindName('grid')
    $lblStatus  = $w.FindName('lblStatus')
    $txtLeft    = $w.FindName('txtLeft')
    $txtRight   = $w.FindName('txtRight')

    $results = New-Object System.Collections.ObjectModel.ObservableCollection[object]
    $grid.ItemsSource = $results

    # ---------------------------- search -------------------------------------
    $searchAction = {
        try {
            $results.Clear(); $txtLeft.Clear(); $txtRight.Clear()
            $term   = $q.Text.Trim()
            $incDis = [bool]($chkDisabled.IsChecked -eq $true)
            $filter = Script:Build-AdComputerFilter -Term $term -IncludeDisabled:$incDis

            $dc = $null; $sb = $null
            try { if ($script:DomainControllerBox) { $dc = $script:DomainControllerBox.Text.Trim() } } catch {}
            try { if ($script:SearchBaseBox)       { $sb = $script:SearchBaseBox.Text.Trim() } } catch {}
            if (-not $dc) { $dc = $DefaultDomainController }
            if (-not $sb) { $sb = $DefaultSearchBase }

            $ldap = "LDAP://$dc/$sb"
            $de = New-Object DirectoryServices.DirectoryEntry(
                $ldap,
                $script:ADCreds.UserName,
                $script:ADCreds.GetNetworkCredential().Password
            )
            $ds = New-Object DirectoryServices.DirectorySearcher($de)
            $ds.Filter   = $filter
            $ds.PageSize = 500
            foreach ($p in @(
                'name','dNSHostName','distinguishedName','operatingSystem',
                'userAccountControl','whenCreated','lastLogonTimestamp','pwdLastSet','description'
            )) { [void]$ds.PropertiesToLoad.Add($p) }

            $sw = [Diagnostics.Stopwatch]::StartNew()
            $rs = $ds.FindAll()

            # add items
            foreach ($r in $rs) {
                $props = $r.Properties
                $uac   = [int](Script:Get-DSValue $props 'useraccountcontrol')
                $llt   = Convert-ADLargeInteger (Script:Get-DSValue $props 'lastlogontimestamp')
                $pwdls = Convert-ADLargeInteger (Script:Get-DSValue $props 'pwdlastset')
                $whenC = Script:Get-DSValue $props 'whencreated'
                if ($whenC -is [datetime]) { $whenC = $whenC.ToString('yyyy-MM-dd HH:mm') }

                $dn    = Script:Get-DSValue $props 'distinguishedname'
                $ou    = Script:Get-OUFromDN $dn

                $obj = [PSCustomObject]@{
                    Name              = (Script:Get-DSValue $props 'name')
                    DNSHostName       = (Script:Get-DSValue $props 'dnshostname')
                    Enabled           = (Get-UacEnabledString -UserAccountControl $uac)
                    OperatingSystem   = (Script:Get-DSValue $props 'operatingsystem')
                    LastLogon         = if ($llt)   { $llt.ToString('yyyy-MM-dd HH:mm') }   else { $null }
                    PwdLastSet        = if ($pwdls) { $pwdls.ToString('yyyy-MM-dd HH:mm') } else { $null }
                    WhenCreated       = $whenC
                    DistinguishedName = $dn
                    OU                = $ou
                    Description       = (Script:Get-DSValue $props 'description')
                    LoggedOnUser      = $null   # filled on selection
                }
                $results.Add($obj) | Out-Null
            }
            $sw.Stop()

            # default sort by Name (ascending)
            $grid.Items.SortDescriptions.Clear()
            $sort = New-Object System.ComponentModel.SortDescription('Name',[System.ComponentModel.ListSortDirection]::Ascending)
            $grid.Items.SortDescriptions.Add($sort)
            $grid.Items.Refresh()

            $lblStatus.Text = "Found $($results.Count) computer object(s) in $([int]$sw.Elapsed.TotalMilliseconds) ms."
        }
        catch { $lblStatus.Text = "Search failed: $($_.Exception.Message)" }
    }

    # ----------------------------- events ------------------------------------
    $btnSearch.Add_Click($searchAction)
    $q.Add_KeyDown({ if ($_.Key -eq 'Return') { & $searchAction } })
    $btnClear.Add_Click({ $q.Clear(); $results.Clear(); $txtLeft.Clear(); $txtRight.Clear(); $lblStatus.Text='Ready' })

    $grid.Add_SelectionChanged({
        $sel = $grid.SelectedItem
        if ($sel) {
            # resolve logged-on user quickly (2s timeout)
            $target = if ($sel.DNSHostName) { $sel.DNSHostName } else { $sel.Name }
            $who = Script:Get-LoggedOnUserQuick -Computer $target -Credential $script:ADCreds
            if (-not $who) { $who = "N/A" }

            $txtLeft.Text = @"
Name:              $($sel.Name)
DNS Hostname:      $($sel.DNSHostName)
Enabled:           $($sel.Enabled)
Operating System:  $($sel.OperatingSystem)
Logged-on User:    $who
Last Logon:        $($sel.LastLogon)
Password Last Set: $($sel.PwdLastSet)
When Created:      $($sel.WhenCreated)
"@.Trim()

            $txtRight.Text = @"
Distinguished Name:
$($sel.DistinguishedName)

OU:
$($sel.OU)

Description:
$($sel.Description)
"@.Trim()
        } else { $txtLeft.Clear(); $txtRight.Clear() }
    })

    $w.Add_ContentRendered({ $q.Focus() })
    [void]$w.ShowDialog()
}

# -------------------- Main GUI --------------------
function Show-MainGUI {
    [xml]$xaml = @"
<Window xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
        xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
        Title='Join / Unjoin Computer Tool'
        Width="950" SizeToContent="Height"
        WindowStartupLocation="CenterScreen"
        Background="#FFF5F7FB"
        FontFamily='Segoe UI'
        FontSize='14'>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height='Auto'/>
      <RowDefinition Height='*'/>
      <RowDefinition Height='Auto'/>
    </Grid.RowDefinitions>

   <!-- Header -->
<Border Grid.Row="0" Padding="16">
  <Border.Background>
    <LinearGradientBrush StartPoint="0,0" EndPoint="1,0">
      <GradientStop Color="#2563EB" Offset="0.0"/>   <!-- blue -->
      <GradientStop Color="#7C3AED" Offset="0.55"/>  <!-- purple -->
      <GradientStop Color="#06B6D4" Offset="1.0"/>   <!-- cyan -->
    </LinearGradientBrush>
  </Border.Background>

  <Grid>
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="Auto"/>
      <ColumnDefinition Width="*"/>
      <ColumnDefinition Width="Auto"/>
    </Grid.ColumnDefinitions>

    
    <StackPanel Grid.Column="1">
      <TextBlock Text="Join / Unjoin Computer Tool"
                 Foreground="White" FontSize="22" FontWeight="Bold"/>
      <TextBlock x:Name="HeaderDesc"
                 Text="Manage AD join/disjoin, credential validation, DC reachability, and device info."
                 Foreground="#E0E7FF" FontSize="12" Margin="0,4,0,0"
                 TextWrapping="Wrap"/>
    </StackPanel>

    
  </Grid>
</Border>

    <!-- Body -->
    <Grid Grid.Row='1' Margin='12'>
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width='2.3*'/>
        <ColumnDefinition Width='2*'/>
      </Grid.ColumnDefinitions>

      <!-- Left column -->
      <StackPanel Grid.Column='0' Margin='6'>

        <!-- Domain Configuration -->
        <Border Padding='12' Margin='0,0,0,10' CornerRadius='12' Background='#FFFFFF' BorderBrush='#E5E7EB' BorderThickness='1'>
          <Border.Effect><DropShadowEffect Color='#000000' Opacity='0.06' ShadowDepth='1' BlurRadius='6'/></Border.Effect>
          <StackPanel>
            <TextBlock Text='Domain Configuration' FontSize='16' FontWeight='Bold' Margin='0,0,0,8'/>
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height='Auto'/><RowDefinition Height='Auto'/><RowDefinition Height='Auto'/><RowDefinition Height='Auto'/>
                <RowDefinition Height='Auto'/><RowDefinition Height='Auto'/><RowDefinition Height='Auto'/>
              </Grid.RowDefinitions>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width='140'/>
                <ColumnDefinition Width='*'/>
              </Grid.ColumnDefinitions>

              <TextBlock Grid.Row='0' Grid.Column='0' Text='Domain Controller:' VerticalAlignment='Center' />
              <TextBox   Grid.Row='0' Grid.Column='1' x:Name='DomainControllerBox' Text='$DefaultDomainController' Margin='2,4,0,4' Background='#F3F4F6' />

              <TextBlock Grid.Row='1' Grid.Column='0' Text='Domain Name:' VerticalAlignment='Center'/>
              <TextBox   Grid.Row='1' Grid.Column='1' x:Name='DomainNameBox' Text='$DefaultDomainName' Margin='2,4,0,4' Background='#F3F4F6' />

              <TextBlock Grid.Row='2' Grid.Column='0' Text='Search Base (OU):' VerticalAlignment='Center'/>
              <TextBox   Grid.Row='2' Grid.Column='1' x:Name='SearchBaseBox' Text='$DefaultSearchBase' Margin='2,4,0,4' Background='#F3F4F6' />

              <TextBlock Grid.Row='3' Grid.Column='0' Text='Selected OU:' VerticalAlignment='Center' Visibility="Collapsed"/>
              <TextBox   Grid.Row='3' Grid.Column='1' x:Name='SelectedOUBox' Margin='2,4,0,4' Background='#FFFFFF' BorderBrush='#93C5FD' Visibility="Collapsed"/>

              <TextBlock Grid.Row='4' Grid.Column='0' Text='Domain Credentials:' VerticalAlignment='Center'/>
              <Button    Grid.Row='4' Grid.Column='1' x:Name='CredBtn' Content='User / Password' Height='30' Margin='2,4,0,4'
                         Background='#3B82F6' Foreground='White' BorderThickness='0' Cursor='Hand'/>

              <TextBlock Grid.Row='5' Grid.Column='0' Text='Username:' VerticalAlignment='Center'/>
              <TextBlock Grid.Row='5' Grid.Column='1' x:Name='CredUserBlock' Text='Not Set' Margin='0,4,0,4'/>

              <TextBlock Grid.Row='6' Grid.Column='0' Text='Cred Status:' VerticalAlignment='Center'/>
              <TextBlock Grid.Row='6' Grid.Column='1' x:Name='CredStatusBlock' Text='Not Connected' Margin='2,4,0,0' Padding='6,2'
                         Background='#DC3545' Foreground='White'/>
            </Grid>
          </StackPanel>
        </Border>

        <!-- PC Information -->
        <Border Padding='12' CornerRadius='12' Background='#FFFFFF' BorderBrush='#E5E7EB' BorderThickness='1'>
          <Border.Effect><DropShadowEffect Color='#000000' Opacity='0.06' ShadowDepth='1' BlurRadius='6'/></Border.Effect>
          <StackPanel>
            <TextBlock Text='PC Information' FontSize='16' FontWeight='Bold' Margin='0,0,0,8'/>
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height='Auto'/><RowDefinition Height='Auto'/><RowDefinition Height='Auto'/>
                <RowDefinition Height='Auto'/><RowDefinition Height='Auto'/><RowDefinition Height='Auto'/><RowDefinition Height='Auto'/>
              </Grid.RowDefinitions>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width='130'/>
                <ColumnDefinition Width='*'/>
              </Grid.ColumnDefinitions>

              <TextBlock Grid.Row='0' Grid.Column='0' Text='Computer Name:'/>
              <TextBlock Grid.Row='0' Grid.Column='1' x:Name='PcNameBlock' Text='Loading...' Margin='2,2,0,2'/>

              <TextBlock Grid.Row='1' Grid.Column='0' Text='IP Address:'/>
              <TextBlock Grid.Row='1' Grid.Column='1' x:Name='PcIPAddressBlock' Text='Loading...' Margin='2,2,0,2'/>

              <TextBlock Grid.Row='2' Grid.Column='0' Text='Domain Status:'/>
              <TextBlock Grid.Row='2' Grid.Column='1' x:Name='PcDomainStatusBlock' Text='Loading...' Margin='2,2,0,2' Padding='6,2' />

              <TextBlock Grid.Row='3' Grid.Column='0' Text='DC Status:'/>
              <TextBlock Grid.Row='3' Grid.Column='1' x:Name='DcReachBlock' Text='Checking...' Margin='2,2,0,2' Padding='6,2' />

              <TextBlock Grid.Row='4' Grid.Column='0' Text='Entra ID Status:'/>
              <TextBlock Grid.Row='4' Grid.Column='1' x:Name='PcEntraIDStatusBlock' Text='Loading...' Margin='2,2,0,2' Padding='6,2'/>

              <TextBlock Grid.Row='5' Grid.Column='0' Text='SCCM Agent:'/>
              <TextBlock Grid.Row='5' Grid.Column='1' x:Name='SCCMStatusBlock' Text='Loading...' Margin='2,2,0,2' Padding='6,2'/>

              <TextBlock Grid.Row='6' Grid.Column='0' Text='Co-Management:'/>
              <TextBlock Grid.Row='6' Grid.Column='1' x:Name='CoManagementBlock' Text='Loading...' Margin='2,2,0,2' Padding='6,2'/>
            </Grid>

            <Button x:Name='RefreshBtn' Content='Refresh Info' Height='34' Margin='0,10,0,0'
                    Background='#3B82F6' Foreground='White' BorderThickness='0' Cursor='Hand'/>
          </StackPanel>
        </Border>

      </StackPanel>

      <!-- Right column -->
      <StackPanel Grid.Column='1' Margin='6'>

        <!-- Active Directory Actions -->
<Border Padding="12" Margin="0,0,0,10" CornerRadius="12"
        Background="#FFFFFF" BorderBrush="#E5E7EB" BorderThickness="1">
  <Border.Effect>
    <DropShadowEffect Color="#000000" Opacity="0.06" ShadowDepth="1" BlurRadius="6"/>
  </Border.Effect>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>  <!-- Title -->
      <RowDefinition Height="Auto"/>  <!-- 2 buttons -->
      <RowDefinition Height="Auto"/>  <!-- Delete full width -->
    </Grid.RowDefinitions>
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="*"/>
      <ColumnDefinition Width="*"/>
    </Grid.ColumnDefinitions>

    <TextBlock Grid.Row="0" Grid.ColumnSpan="2"
               Text="Active Directory Actions"
               FontSize="16" FontWeight="Bold" Margin="0,0,0,8"/>

    <!-- Full-width across each half -->
    <Button x:Name="JoinButton"
            Grid.Row="1" Grid.Column="0"
            Content="Join Domain"
            Height="38" Margin="5" Padding="16,0"
            HorizontalAlignment="Stretch"
            Background="#10B981" Foreground="White" BorderThickness="0" Cursor="Hand"/>

    <Button x:Name="DisjoinButton"
            Grid.Row="1" Grid.Column="1"
            Content="Disjoin from Domain"
            Height="38" Margin="5" Padding="16,0"
            HorizontalAlignment="Stretch"
            Background="#F59E0B" Foreground="White" BorderThickness="0" Cursor="Hand"/>

            <!-- Full-width delete button -->
    <Button x:Name="FindinAD"
            Grid.Row="2" Grid.Column="0"
            Content="Find in AD"
            Height="38" Margin="5" Padding="16,0"
            HorizontalAlignment="Stretch"
            Background="#3B82F6" Foreground="White" BorderThickness="0" Cursor="Hand"/>

    <!-- Full-width delete button -->
    <Button x:Name="DeleteButton"
            Grid.Row="2" Grid.Column="1"
            Content="Delete from Domain"
            Height="38" Margin="5" Padding="16,0"
            HorizontalAlignment="Stretch"
            Background="#EF4444" Foreground="White" BorderThickness="0" Cursor="Hand"/>
  </Grid>
</Border>


        <!-- Message Center -->
        <Border Padding='12' CornerRadius='12' Background='#FFFFFF' BorderBrush='#E5E7EB' BorderThickness='1'>
          <Border.Effect><DropShadowEffect Color='#000000' Opacity='0.06' ShadowDepth='1' BlurRadius='6'/></Border.Effect>
          <StackPanel>
            <TextBlock Text='Message Center' FontSize='16' FontWeight='Bold' Margin='0,0,0,8'/>
            <RichTextBox x:Name='LogBox'
                         Height='305'
                         IsReadOnly='True'
                         VerticalScrollBarVisibility='Auto'
                         Background='#0F172A'
                         Foreground='#E5E7EB'
                         BorderBrush='#1F2937'
                         BorderThickness='1'
                         FontFamily='Consolas'
                         FontSize='12'
                         Padding='8'/>
          </StackPanel>
        </Border>

      </StackPanel>
    </Grid>

   <!-- Footer (gradient, 3 zones) -->
<Border Grid.Row="2" Padding="8">
  <Border.Background>
    <LinearGradientBrush StartPoint="0,0" EndPoint="1,0">
      <GradientStop Color="#2563EB" Offset="0.0"/>  <!-- sky -->
      <GradientStop Color="#06B6D4" Offset="1.0"/>  <!-- green -->
    </LinearGradientBrush>
  </Border.Background>

  <Grid>
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="*"/>
      <ColumnDefinition Width="Auto"/>
      <ColumnDefinition Width="*"/>
    </Grid.ColumnDefinitions>

    <TextBlock x:Name="FooterLeft"
               Grid.Column="0" Text="IT Operations"
               Foreground="White" FontSize="12"
               HorizontalAlignment="Left" Margin="6,0"/>

    <TextBlock x:Name="FooterCenter"
               Grid.Column="1"
              Text=""
               Foreground="White" FontSize="12"
               HorizontalAlignment="Center"/>

    <TextBlock x:Name="FooterRight"
               Grid.Column="2" 
                Text="© 2025 M.omar (momar.tech) — All Rights Reserved"
               Foreground="White" FontSize="12"
               HorizontalAlignment="Right" Margin="0,0,6,0"/>
  </Grid>
</Border>


  </Grid>
</Window>
"@

    $w = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))

    # Expose controls
    $script:DomainControllerBox = $w.FindName('DomainControllerBox')
    $script:DomainNameBox = $w.FindName('DomainNameBox')
    $script:SearchBaseBox = $w.FindName('SearchBaseBox')
    $script:SelectedOUBox = $w.FindName('SelectedOUBox')

    $script:CredBtn = $w.FindName('CredBtn')
    $script:CredUserBlock = $w.FindName('CredUserBlock')
    $script:CredStatusBlock = $w.FindName('CredStatusBlock')

    $script:PcNameBlock = $w.FindName('PcNameBlock')
    $script:PcIPAddressBlock = $w.FindName('PcIPAddressBlock')
    $script:PcDomainStatusBlock = $w.FindName('PcDomainStatusBlock')
    $script:DcReachBlock = $w.FindName('DcReachBlock')
    $script:PcEntraIDStatusBlock = $w.FindName('PcEntraIDStatusBlock')
    $script:SCCMStatusBlock = $w.FindName('SCCMStatusBlock')
    $script:CoManagementBlock = $w.FindName('CoManagementBlock')

    $JoinButton = $w.FindName('JoinButton')
    $DisjoinButton = $w.FindName('DisjoinButton')
    $DeleteButton = $w.FindName('DeleteButton')
    $RefreshBtn = $w.FindName('RefreshBtn')
    $FindinAD = $w.FindName('FindinAD')
    $script:LogBox = $w.FindName('LogBox')

    # Handlers
    $RefreshBtn.Add_Click({ Update-PCInfo })

    $JoinButton.Add_Click({
            if (-not $script:ADCreds) { $null = Prompt-Credentials }
            if (-not $script:ADCreds) {
                $script:CredUserBlock.Text = "Not Set"
                $script:CredStatusBlock.Text = "Not Connected"
                $script:CredStatusBlock.Background = $script:BrushBad
                $script:CredStatusBlock.Foreground = $script:BrushWhite
                return
            }
            Show-OUWindow
            $sel = $script:SelectedOUBox.Text.Trim()
            if ($sel -ne "") {
                Join-ComputerWithOU -DomainName $script:DomainNameBox.Text.Trim() -OUPath $sel
            }
            else {
                Write-UiLog "No OU selected. Click Select OU first." "ERROR"
                Show-WPFMessage -Message "No OU selected. Click Select OU first." -Title "Error" -Color Red
            }
        })

    $DisjoinButton.Add_Click({
            if (-not $script:ADCreds) { $null = Prompt-Credentials }
            if (-not $script:ADCreds) { return }
            Disjoin-ComputerFromDomain -ComputerName $env:COMPUTERNAME
        })

    $DeleteButton.Add_Click({
            if (-not $script:ADCreds) { $null = Prompt-Credentials }
            if (-not $script:ADCreds) { return }
            Delete-ComputerFromAD -ComputerName $env:COMPUTERNAME `
                -DomainController $script:DomainControllerBox.Text.Trim() `
                -SearchBase $script:SearchBaseBox.Text.Trim()
        })

    $FindinAD.Add_Click({
        if (-not $script:ADCreds) { $null = Prompt-Credentials }
        if ($script:ADCreds) { Show-FindInADWindow }
    })
    # Credentials button
    $script:CredBtn.Add_Click({
            $null = Prompt-Credentials
            if ($script:ADCreds) {
                $script:CredUserBlock.Text = $script:ADCreds.UserName
                $script:CredStatusBlock.Text = "Connected"
                $script:CredStatusBlock.Background = $script:BrushOK
                $script:CredStatusBlock.Foreground = $script:BrushWhite
            }
            else {
                $script:CredUserBlock.Text = "Not Set"
                if ($script:LastCredError) {
                    $script:CredStatusBlock.Text = "Error: $script:LastCredError"
                }
                else {
                    $script:CredStatusBlock.Text = "Not Connected"
                }
                $script:CredStatusBlock.Background = $script:BrushBad
                $script:CredStatusBlock.Foreground = $script:BrushWhite
            }
        })

    # Initial state
    $script:CredUserBlock.Text = "Not Set"
    $script:CredStatusBlock.Text = "Not Connected"
    $script:CredStatusBlock.Background = $script:BrushBad
    $script:CredStatusBlock.Foreground = $script:BrushWhite

    Write-UiLog "Ready. Set credentials, then use AD actions as needed." "INFO"
    Update-PCInfo

    [void]$w.ShowDialog()
}

# -------------------- Main --------------------
try {
    Test-Admin
    Show-MainGUI
}
catch {
    Show-WPFMessage -Message ("Unexpected error:`n{0}" -f $_.Exception.Message) -Title "Error" -Color Red
}
