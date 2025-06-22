$DaysInactive = 180
$InactiveComputers = Search-ADAccount -AccountInactive -ComputersOnly -TimeSpan ([TimeSpan]::FromDays($DaysInactive))

foreach ($Computer in $InactiveComputers) {
    $ComputerName = $Computer.Name
    $LastLogonDate = $Computer.LastLogonDate
    Write-Host "Computer Name: $ComputerName"
    Write-Host "Last Logon Date: $LastLogonDate"
    Write-Host "---------------------------"
}