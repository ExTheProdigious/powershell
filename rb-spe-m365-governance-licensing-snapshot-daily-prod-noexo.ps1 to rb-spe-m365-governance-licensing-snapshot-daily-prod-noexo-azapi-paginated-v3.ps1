$DcrImmutableIdUsers = "dcr-ed4492d3cd6a46debdf6c61c0f2780e0"
$DcrImmutableIdSku = "dcr-ed4492d3cd6a46debdf6c61c0f2780e0"
$LogsIngestionEndpoint = "https://dce-spe-m365-prod-n9ow.westus2-1.ingest.monitor.azure.com"
$UsersStreamName = "Custom-SPE_M365_LicensedUsers_CL"
$SkuStreamName = "Custom-SPE_M365_LicenseSkuMetrics_CL"
$SubscriptionIdOrName = "azr-spe-m365-prd"
$ExchangeOrganization = "spe.onmicrosoft.com"
$ManagedIdentityAccountId = "1a14f2ab-7c6a-4c07-8233-afd7bcc03592"

$ErrorActionPreference = 'Stop'

# Helpers

function Convert-SecureStringToPlainText {
    param(
        [Parameter(Mandatory)]
        [System.Security.SecureString]$SecureString
    )

    $ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    try {
        return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
    }
    finally {
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
    }
}

function Get-MonitorBearerToken {
    $secureToken = (Get-AzAccessToken -ResourceUrl "https://monitor.azure.com/").Token
    return (Convert-SecureStringToPlainText -SecureString $secureToken)
}

function Get-SafeString {
    param([object]$Value)
    if ($null -eq $Value) { return "" }
    return [string]$Value
}

function Join-UniqueValues {
    param(
        [AllowEmptyCollection()]
        [object[]]$Values
    )

    if (-not $Values -or $Values.Count -eq 0) {
        return ""
    }

    return (
        $Values |
        Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
        ForEach-Object { [string]$_ } |
        Sort-Object -Unique
    ) -join ';'
}

function Get-LicenseAssignmentSummary {
    param(
        [Parameter(Mandatory)]
        [object[]]$LicenseAssignmentStates
    )

    $groupIds = @()
    $hasDirect = $false
    $hasGroup = $false

    foreach ($state in $LicenseAssignmentStates) {
        $assignedByGroup = $state.AssignedByGroup
        if ([string]::IsNullOrWhiteSpace([string]$assignedByGroup)) {
            $hasDirect = $true
        }
        else {
            $hasGroup = $true
            $groupIds += [string]$assignedByGroup
        }
    }

    $method =
        if ($hasDirect -and $hasGroup) { 'Mixed' }
        elseif ($hasGroup) { 'GroupOnly' }
        elseif ($hasDirect) { 'DirectOnly' }
        else { 'Unknown' }

    [pscustomobject]@{
        HasDirectAssignedLicense_b = [bool]$hasDirect
        HasGroupAssignedLicense_b  = [bool]$hasGroup
        LicenseAssignmentMethod_s  = $method
        AssignedByGroupIds_s       = (Join-UniqueValues -Values $groupIds)
    }
}

function Send-LogIngestionBatch {
    param(
        [Parameter(Mandatory)] [array]$Rows,
        [Parameter(Mandatory)] [string]$DcrImmutableId,
        [Parameter(Mandatory)] [string]$StreamName,
        [Parameter(Mandatory)] [string]$LogsIngestionEndpoint,
        [Parameter(Mandatory)] [string]$MonitorToken
    )

    if (-not $Rows -or $Rows.Count -eq 0) {
        Write-Output "No rows to send for stream [$StreamName]."
        return
    }

    $uri = '{0}/dataCollectionRules/{1}/streams/{2}?api-version=2023-01-01' -f `
        $LogsIngestionEndpoint.TrimEnd('/'),
        $DcrImmutableId,
        $StreamName

    $headers = @{
        "Authorization" = "Bearer $monitorToken"
        "Content-Type" = "application/json"
    }

    $batchSize = 500
    for ($i = 0; $i -lt $Rows.Count; $i += $batchSize) {
        $upper = [Math]::Min($i + $batchSize - 1, $Rows.Count - 1)
        $batch = @($Rows[$i..$upper])
        $body = $batch | ConvertTo-Json -Depth 10

        Write-Output "Posting $($batch.Count) rows to [$StreamName]"
        Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -Body $body | Out-Null
    }
}


# Auth

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

$monitorToken = Get-MonitorBearerToken

# Snapshot metadata

$snapshotTime = (Get-Date).ToUniversalTime()
$runId = [guid]::NewGuid().Guid

# Preload EXO hold data

Write-Output "Getting SKUs."

$skus = Get-MgSubscribedSku -All

# skuId -> skuPartNumber
$skuPartNumberMap = @{}

# skuPartNumber -> friendly display name
$skuDisplayNameMap = @{
    "AAD_PREMIUM_P2"                    = "Microsoft Entra ID P2"
    "ATA"                               = "Microsoft Defender for Identity"
    "DEFENDER_ENDPOINT_P1"              = "Microsoft Defender for Endpoint P1"
    "EMS"                               = "Enterprise Mobility + Security E3"
    "ENTERPRISEPACK"                    = "Office 365 E3"
    "FLOW_FREE"                         = "Microsoft Power Automate Free"
    "Microsoft_365_E3_Extra_Features"   = "Microsoft 365 E3 Extra Features"
    "Microsoft_Teams_Rooms_Basic"       = "Microsoft Teams Rooms Basic"
    "POWERAPPS_DEV"                     = "Microsoft Power Apps for Developer"
    "STANDARDPACK"                      = "Office 365 E1"
    "CCIBOTS_PRIVPREV_VIRAL"            = "Microsoft Copilot Studio Viral Trial"
    "DYN365_FINANCIALS_ACCOUNTANT_SKU"  = "Dynamics 365 Business Central External Accountant"
    "EXCHANGEENTERPRISE"                = "Exchange Online (Plan 2)"
    "MCOCAP"                            = "Microsoft Teams Shared Devices"
    "MCOEV"                             = "Microsoft Teams Phone Standard"
    "Microsoft_Teams_Rooms_Pro"         = "Microsoft Teams Rooms Pro"
    "PBI_PREMIUM_PER_USER"              = "Power BI Premium Per User"
    "PBI_PREMIUM_PER_USER_ADDON"        = "Power BI Premium Per User Add-On"
    "PHONESYSTEM_VIRTUALUSER"           = "Microsoft Teams Phone Resource Account"
    "POWER_BI_PRO"                      = "Power BI Pro"
    "POWER_BI_STANDARD"                 = "Microsoft Fabric (Free)"
    "POWERAPPS_PER_USER"                = "Power Apps Premium"
    "POWERAPPS_VIRAL"                   = "Microsoft Power Apps Plan 2 Trial"
    "POWERAUTOMATE_ATTENDED_RPA"        = "Power Automate Premium"
    "PROJECTPROFESSIONAL"               = "Planner and Project Plan 3"
    "VISIOCLIENT"                       = "Visio Plan 2"
    "VISIO_PLAN2_DEPT"                  = "Visio Plan 2"
}

Write-Output "Evaluating SKUs."

$skuRows = foreach ($sku in $skus) {
    $skuId = [string]$sku.SkuId
    $skuPartNumber = [string]$sku.SkuPartNumber
    $skuPartNumberMap[$skuId] = $skuPartNumber

    $enabledUnits   = [double]($sku.PrepaidUnits.Enabled)
    $warningUnits   = [double]($sku.PrepaidUnits.Warning)
    $suspendedUnits = [double]($sku.PrepaidUnits.Suspended)
    $lockedOutUnits = if ($sku.PrepaidUnits.PSObject.Properties.Name -contains 'LockedOut') { [double]$sku.PrepaidUnits.LockedOut } else { 0 }
    $consumedUnits  = [double]$sku.ConsumedUnits
    $availableUnits = $enabledUnits - $consumedUnits
    $utilizationPct = if ($enabledUnits -gt 0) { [math]::Round(($consumedUnits / $enabledUnits) * 100, 2) } else { 0 }

    $servicePlans = @()
    foreach ($sp in $sku.ServicePlans) {
        if ($sp.ServicePlanName) {
            $servicePlans += [string]$sp.ServicePlanName
        }
    }

    $skuDisplayName = if ($skuDisplayNameMap.ContainsKey($skuPartNumber)) {
        $skuDisplayNameMap[$skuPartNumber]
    }
    else {
        $skuPartNumber
    }

    [pscustomobject]@{
        SnapshotTime_t             = $snapshotTime
        RunId_g                    = $runId
        SkuId_g                    = $skuId
        SkuPartNumber_s            = $skuPartNumber
        SkuDisplayName_s           = $skuDisplayName
        ConsumedUnits_d            = $consumedUnits
        EnabledUnits_d             = $enabledUnits
        WarningUnits_d             = $warningUnits
        SuspendedUnits_d           = $suspendedUnits
        LockedOutUnits_d           = $lockedOutUnits
        AvailableUnits_d           = $availableUnits
        UtilizationPercent_d       = $utilizationPct
        CapabilityStatus_s         = Get-SafeString $sku.CapabilityStatus
        LowAvailabilityThreshold_b = [bool]($availableUnits -le 25)
        OverprovisionedThreshold_b = [bool]($enabledUnits -gt 0 -and $utilizationPct -lt 60)
        RelevantServicePlans_s     = (Join-UniqueValues -Values $servicePlans)
    }
}

# -----------------------------
# Users (RAW PAGINATION)
# -----------------------------

Write-Output "Starting Raw Paginated User Retrieval."

$UserCount = 0
# Initial Graph URL
$uri = "https://graph.microsoft.com/v1.0/users?`$top=999&`$select=id,displayName,userPrincipalName,assignedLicenses,licenseAssignmentStates,usageLocation,accountEnabled,userType,department,companyName,onPremisesSyncEnabled,signInActivity"

do {
    Write-Output "Fetching page from: $uri"
    $response = Invoke-MgGraphRequest -Method Get -Uri $uri
    $users = $response.value
    $userRows = @()

    foreach ($user in $users) {
        # Raw JSON uses camelCase: assignedLicenses
        if (-not $user.assignedLicenses -or $user.assignedLicenses.Count -eq 0) { continue }

        $assignedSkuIds = @($user.assignedLicenses | ForEach-Object { [string]$_.skuId })
        
        $assignedSkuPartNumbers = foreach ($skuId in $assignedSkuIds) {
            if ($skuPartNumberMap.ContainsKey($skuId)) { $skuPartNumberMap[$skuId] } else { "UNKNOWN_SKU" }
        }

        $assignedDisplayNames = foreach ($skuPartNumber in $assignedSkuPartNumbers) {
            if ($skuDisplayNameMap.ContainsKey($skuPartNumber)) { $skuDisplayNameMap[$skuPartNumber] } else { $skuPartNumber }
        }

        # Raw JSON property: licenseAssignmentStates
        $assignmentSummary = Get-LicenseAssignmentSummary -LicenseAssignmentStates @($user.licenseAssignmentStates)
        
        $overlappingSkuFlag = [bool](
            $assignedSkuIds.Count -ne (@($assignedSkuIds | Sort-Object -Unique).Count) -or
            $assignedSkuPartNumbers.Count -gt 1
        )

        $userRows += [pscustomobject]@{
            SnapshotTime_t                 = $snapshotTime
            RunId_g                        = $runId
            UserId_g                       = [string]$user.id
            UserPrincipalName_s            = Get-SafeString $user.userPrincipalName
            DisplayName_s                  = Get-SafeString $user.displayName
            AccountEnabled_b               = [bool]$user.accountEnabled
            UserType_s                     = Get-SafeString $user.userType
            UsageLocation_s                = Get-SafeString $user.usageLocation
            Department_s                   = Get-SafeString $user.department
            CompanyName_s                  = Get-SafeString $user.companyName
            OnPremisesSyncEnabled_b        = [bool]$user.onPremisesSyncEnabled
            # Nested JSON objects also use camelCase
            LastSuccessfulSignInDateTime_t = if ($user.signInActivity.lastSuccessfulSignInDateTime) { [datetime]$user.signInActivity.lastSuccessfulSignInDateTime } else { $null }

            AssignedLicenseCount_d         = [double]$assignedSkuIds.Count
            AssignedSkuIds_s               = (Join-UniqueValues -Values $assignedSkuIds)
            AssignedSkuPartNumbers_s       = (Join-UniqueValues -Values $assignedSkuPartNumbers)
            AssignedLicenseDisplayNames_s  = (Join-UniqueValues -Values $assignedDisplayNames)

            HasGroupAssignedLicense_b      = [bool]$assignmentSummary.HasGroupAssignedLicense_b
            HasDirectAssignedLicense_b     = [bool]$assignmentSummary.HasDirectAssignedLicense_b
            LicenseAssignmentMethod_s      = $assignmentSummary.LicenseAssignmentMethod_s
            AssignedByGroupIds_s           = $assignmentSummary.AssignedByGroupIds_s
            OverlappingSkuFlag_b           = $overlappingSkuFlag
        }
    }

    if ($userRows.Count -gt 0) {
        Send-LogIngestionBatch -Rows $userRows -DcrImmutableId $DcrImmutableIdUsers -StreamName $UsersStreamName -LogsIngestionEndpoint $LogsIngestionEndpoint -MonitorToken $monitorToken
        $UserCount += $userRows.Count
        Write-Output "Processed $UserCount licensed users..."
    }

    # Grab the next link
    $uri = $response.'@odata.nextLink'

} while ($null -ne $uri)

# -----------------------------
# Finalize SKU Metrics
# -----------------------------

Write-Output "Posting final SKU metrics."
Send-LogIngestionBatch `
    -Rows $skuRows `
    -DcrImmutableId $DcrImmutableIdSku `
    -StreamName $SkuStreamName `
    -LogsIngestionEndpoint $LogsIngestionEndpoint `
    -MonitorToken $monitorToken

Write-Output "RUNBOOK COMPLETE. Total Licensed Users: $UserCount. Total SKUs: $($skuRows.Count)."