<#
.SYNOPSIS
    Automates Windows activation using a KMS server and monitors activation counts.

.DESCRIPTION
    This script performs the following actions:
    1. Checks connectivity to the specified KMS server on port 1688.
    2. Configures Windows to use the specified KMS server for activation.
    3. Attempts to activate Windows via the KMS server.
    4. Retrieves and displays detailed activation information.

.NOTES
    - Run this script with administrative privileges.
    - Ensure that the system can communicate with the KMS server.
#>

# Function to Check for Administrative Privileges
function Test-Administrator {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Ensure the script is running as Administrator
if (-not (Test-Administrator)) {
    Write-Warning "This script must be run as an Administrator. Please rerun the script with elevated privileges."
    exit 1
}

# Variables
$KMS_Server = "KMS16.QassimU.local"
$KMS_Port = 1688

try {
    Write-Output "=== Step 1: Checking Connectivity to KMS Server ===`n"

    # Test connectivity to the KMS server on port 1688
    $connectionTest = Test-NetConnection -ComputerName $KMS_Server -Port $KMS_Port

    if ($connectionTest.TcpTestSucceeded) {
        Write-Output "Success: Able to connect to KMS server '$KMS_Server' on port $KMS_Port.`n"
    }
    else {
        throw "Error: Unable to connect to KMS server '$KMS_Server' on port $KMS_Port. Please ensure the server is reachable and the port is open."
    }

    Write-Output "=== Step 2: Configuring Windows to Use the KMS Server ===`n"

    # Set the KMS server for Windows activation
    $setKMSServerCommand = "slmgr.vbs /skms $KMS_Server : $KMS_Port"
    Write-Output "Executing: $setKMSServerCommand"
    cscript.exe //Nologo slmgr.vbs /skms $KMS_Server : $KMS_Port

    Write-Output "KMS server set to '$KMS_Server' on port $KMS_Port.`n"

    Write-Output "=== Step 3: Activating Windows via KMS Server ===`n"

    # Attempt to activate Windows
    $activateCommand = "slmgr.vbs /ato"
    Write-Output "Executing: $activateCommand"
    cscript.exe //Nologo slmgr.vbs /ato

    Write-Output "Activation command executed.`n"

    # Wait briefly to allow activation to process
    Start-Sleep -Seconds 5

    Write-Output "=== Step 4: Retrieving Detailed Activation Information ===`n"

    # Retrieve detailed license information
    $detailedInfo = cscript.exe //Nologo slmgr.vbs /dlv
    Write-Output "Detailed Activation Information:"
    Write-Output $detailedInfo

    Write-Output "`n=== Activation Process Completed Successfully ==="

} catch {
    Write-Error "An error occurred: $_"
    exit 1
}
