param (
    [string]$DomainController
)

function Reset-SecureChannel {
    param (
        [string]$DC
    )
    Write-Host "Resetting secure channel with domain controller: $DC" -ForegroundColor Cyan
    try {
        $Credential = Get-Credential -Message "Enter domain admin credentials"
        Reset-ComputerMachinePassword -Server $DC -Credential $Credential
        Write-Host "Secure channel reset successfully." -ForegroundColor Green
    } catch {
        Write-Host "Error resetting secure channel: $_" -ForegroundColor Red
    }
}

if (-not $DomainController) {
    Clear-Host
    Write-Host "======================================" -ForegroundColor Cyan
    Write-Host "           Domain Trust Fixer          " -ForegroundColor Cyan
    Write-Host "======================================" -ForegroundColor Cyan
    Write-Host ""
    $DomainController = "DC01"
}

Reset-SecureChannel -DC $DomainController
