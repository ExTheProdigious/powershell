$SubscriptionIdOrName = "azr-spe-m365-prd"
$ExchangeOrganization = "spe.onmicrosoft.com"
$ManagedIdentityAccountId = "45965a03-86a9-4c53-982b-0df66561dc85"

Write-Output "RUNBOOK STARTED"

Connect-ExchangeOnline -ManagedIdentity -Organization "spe.onmicrosoft.com" -ManagedIdentityAccountId "45965a03-86a9-4c53-982b-0df66561dc85"