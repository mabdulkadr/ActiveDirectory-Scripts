$DaysInactive = 180
$SearchBaseOU = "OU=Domain Computers,DC=QassimU,DC=local"  # Replace with your specific OU

# Search for inactive computers in the specified OU
$InactiveComputers = Search-ADAccount -AccountInactive -ComputersOnly -TimeSpan ([TimeSpan]::FromDays($DaysInactive)) -SearchBase $SearchBaseOU

# Prepare a collection to store the results
$Results = @()

# Loop through each inactive computer
foreach ($Computer in $InactiveComputers) {
    $ComputerName = $Computer.Name
    $LastLogonDate = $Computer.LastLogonDate

    Write-Host "Disabling computer $ComputerName"
    Disable-ADAccount -Identity $Computer.DistinguishedName
    Write-Host "Computer disabled successfully."
    Write-Host "---------------------------"

    # Add the result to the collection
    $Results += [PSCustomObject]@{
        ComputerName = $ComputerName
        LastLogonDate = $LastLogonDate
    }
}

# Export the results to a CSV file
$Results | Export-Csv -Path "C:\path\to\output\InactiveComputers.csv" -NoTypeInformation

Write-Host "Results exported to CSV file successfully."
