$SubscriptionIdOrName = "azr-spe-m365-prd"
$ExchangeOrganization = "spe.onmicrosoft.com"
$ManagedIdentityAccountId = "45965a03-86a9-4c53-982b-0df66561dc85"

Write-Output "RUNBOOK STARTED"

try {
    Write-Output "Attempting Azure login..."
    Connect-AzAccount -Identity -AccountId $ManagedIdentityAccountId | Out-Null
    Write-Output "Azure login SUCCESS"
}
catch {
    Write-Output "MI login FAILED"
    Write-Output $_
}

try {
    Write-Output "Attempting Graph login..."
    Connect-MgGraph -Identity -ClientId $ManagedIdentityAccountId -NoWelcome | Out-Null
    Write-Output "Graph login SUCCESS"
}
catch {
    Write-Output "Graph login FAILED"
    Write-Output $_
}

try {
    Write-Output "Attempting EXO login..."
    Connect-ExchangeOnline -ManagedIdentity -Organization $ExchangeOrganization -ManagedIdentityAccountId $ManagedIdentityAccountId -ShowBanner:$false | Out-Null
    Write-Output "EXO login SUCCESS"
}
catch {
    Write-Output "EXO login FAILED"
    Write-Output $_
}
