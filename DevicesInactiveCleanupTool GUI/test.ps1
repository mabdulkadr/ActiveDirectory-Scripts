try {
    # Ensure the Active Directory module is loaded
    Import-Module ActiveDirectory -ErrorAction Stop
} catch {
    Write-Error "Active Directory module is not installed or available."
    exit
}

# Retrieve all computer objects from the domain
$computers = Get-ADComputer -Filter * -Properties Name

# Display the total number of devices (computers)
Write-Output "Total devices in the domain: $($computers.Count)"