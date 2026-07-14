<# License Analysis and Usage Reporting Script
Created By - Tata Consultancy Services Limited
Purpose - Fetch the license information of all users who own atleast a single license with details of Last Login and Usage across M365 Apps and other modules

Permissions
1. Connect using Azure AD login with requisite permissions or preferably connect with MFA credentials using Delegated Application Privileges
2. Uses Microsoft Graph API (MgBetaUser) to fetch the requisite User details
3. Requires User.Read.All, Organization.Read.All permissions
4. For usage reports, Reports.Read.All permissions need to be granted

Pre-requisites
1. MS Graph API needs to be installed and imported prior to running this script
2. Application credentials with required privileges
3. Output is in the form of Excel CSV reports exported to local drive

 Usage Flow
 1. Connect to M365 Tenant
 2. Get Service Plans - Used to fetch all M365 Product Licensing plans and subscriptions and related users
 3. Get Licensed Users - Fetch user details including User details, license allocated, last signin activity
 4. Get GroupLicenses in any
 5. Get User details by specific license family for quick check (E3, E5, F1, Power BI Premium etc.)
 6. Get Unlicensed Users
#>
 
#$tenantId = Get-AutomationVariable -Name 'TenantID' #"f93f5d4d-6bff-47c3-bb89-a8beeabf78f8"
$ManagedIdentityClientId  = "24ff3033-1beb-4758-a22f-bfb5a73b6f7c"
$ManagedIdentityObjectId  = "45965a03-86a9-4c53-982b-0df66561dc85" # PnP user-assigned MI support commonly uses object/principal ID
$spSiteUrl = "https://spe.sharepoint.com"
$spSiteName = "SPE-IT-M365Automations"
$spLibraryName = "LicenseReports"
$accessToken =$null
$DestinationURL ="$spSiteUrl/sites/$spSiteName/$spLibraryName"


<# 1. Connect to Tenant - Requires Azure App or any other secure method #>
Function ConnectToGraph()
{
   try
    {
          #Connect-MgGraph -ClientId $clientId -TenantID $TenantId -CertificateThumbprint $certThumbprint  -Nowelcome
          Connect-MgGraph -Identity -ClientId $ManagedIdentityClientId -Nowelcome
          Write-Output "Connected to Graph"
       <# $graphtokenBody = @{
            Grant_Type    = "client_credentials"
            Client_Id     = $clientId
            Client_Secret = $clientSecret
            Scope         = "https://graph.microsoft.com/.default"
        }

        $jsonBody = $graphtokenBody #| ConvertTo-Json
        $oauth = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body $jsonBody 
        $global:authHeader = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}
        $global:authHeader["ConsistencyLevel"] = "eventual"     #>      
      
   }
  
    catch {
            Write-Error $Error[0]
            Write-Output "Error in connecting to Microsoft Graph"
            $Error=$null
          }
}

<# 2. Connect to SharePoint - Requires a SP Site to upload the generated reports #>
function ConnectToSharePoint
{
    $SiteUrl = "$spSiteUrl/sites/$spSiteName"
    try 
    {
        Write-Output "Trying to connect to $SiteUrl"
         $SiteUrl =$spSiteUrl + "/sites/" +$spSiteName
        #Connect-PnPOnline -Url $SiteUrl -ClientId $clientId -Thumbprint $certThumbprint -Tenant $tenantId  -WarningAction Ignore
        Connect-PnPOnline -Url $Url -ManagedIdentity -UserAssignedManagedIdentityObjectId $ManagedIdentityObjectId
        Write-Output "Connected to $SiteUrl"
    }
    catch{ 
            Write-Error $Error[0]
             $Error=$null
            Write-Output "Error in connecting to SharePoint Online"
           
            }
}

<# 3. Fetch all Service Plan details from Get-MgSubscribedSku API #>
Function GetServicePlans
{
    param ( $CSVFileName)
    $filePath=$null
  
    $serviceplans=$null
    try
    { 
        $Uri = "https://graph.microsoft.com/v1.0/subscribedSkus"
        
        [array]$SkuData = Invoke-RestMethod -Uri $Uri -Method Get -Headers $global:authHeader
        $ServicePlanReport = [System.Collections.Generic.List[Object]]::new()
        $subscribedUsers = [System.Collections.Generic.List[Object]]::new()  
        $rownum=0     
        Foreach( $Sku in $SkuData.Value)
        { 
            
            $obj = [pscustomObject][ordered] @{
                SkuId= $Sku.SkuId
                Id= $Sku.id
                SkuPartNumber=$Sku.skuPartNumber
                Enabled= $Sku.prepaidUnits.enabled
                Consumed =$Sku.consumedUnits
                LockedOut= $Sku.prepaidUnits.lockedOut
                Suspended= $Sku.prepaidUnits.suspended
                Warning= $Sku.prepaidUnits.warning
                Status =$Sku.capabilityStatus
                AccountName = $Sku.accountName
                Accoundid =$Sku.accountId
                plans=$Sku.servicePlans.servicePlanName -join ","  
                }
            $ServicePlanReport.Add($obj)
               
           } #end for reach SKU 
   
     $filePath =$env:Temp
    
     try {
   		    $ServicePlanReport | Export-Csv -Path  $filePath\$CSVFileName -NoTypeInformation 
     		 $Values = @{"Title" = $CSVFileName}
    		# Add the file to the Reports folder
      	 	 WritetoSharePoint($CSVFileName)
    	    
    	}#try	
	catch {  Write-Error $Error[0]
            Write-Output "Error in generating report"
            $Error=$null
            } 
    } # end try
    catch
    {
        Write-Output "Error in getting service plans"
        Write-Output $Error[0]
    }
} #function

<#3. Fetch all users who have been assigned at least one license with their Last activity details using Get-MgBetaUser #>
Function GetLicensedUsers
{
    param ( $CSVFileName)
    try
    {
        $Uri="https://graph.microsoft.com/beta/users?`$filter=assignedLicenses/`$count ne 0&`$count=true&`$select=id,displayName,createdDateTime,userPrincipalName,jobTitle,department,assignedLicenses,assignedPlans,accountEnabled,companyName,employeeType,officeLocation,signInActivity,usageLocation,refreshTokensValidFromDateTime,licenseAssignmentStates&`$top=999" 
        $report = [System.Collections.Generic.List[Object]]::new()
        $rownum =0
        Do
        {
            $users =  Invoke-RestMethod -Uri $Uri -Method Get -Headers $global:authHeader
            $Uri = $users.'@odata.nextLink'

            foreach ($user in $users.Value)
            {
                $UPN=$user.userPrincipalName
                $AccountEnabled =$user.accountEnabled
                $rownum++
                $logindate =$user.signInActivity.lastSignInDateTime
                    if ($logindate -eq $null) 
                {
                    $LastInteractiveSignIn = $null
                    $InactiveDays_InteractiveSignIn =$null
                }
                else
                {
                
                $InactiveDays_InteractiveSignIn = (New-TimeSpan -Start $logindate).Days
                $LastInteractiveSignIn =$logindate.ToString("dd-MM-yyyy hh:mm:ss")
                }
                  $logindate =$user.signInActivity.lastNonInteractiveSignInDateTime
                if ($logindate -eq $null)
                    {
                            $LastNonInteractiveSignIn = $null
                            $InactiveDays_NonInteractiveSignIn = $null
                    }
                else
                {
                            $LastNonInteractiveSignIn =$logindate.ToString("dd-MM-yyyy hh:mm:ss")
                            $InactiveDays_NonInteractiveSignIn = (New-TimeSpan -Start $logindate).Days
                }
                if ($UPN -notmatch '#EXT#')
                {
                $Type = "Member"
                }
                else
                {
                $Type="External User"
                }
                
                if($AccountEnabled -eq $true)
                {
                    $AccountStatus='Enabled'
                }
                else
                {
                    $AccountStatus='Disabled'
                }
             $obj = [pscustomObject][ordered] @{
                No=$rownum
                Id =$user.id
                DisplayName = $user.displayName
                UserPrincipalName =$user.userPrincipalName
                UserType=$Type
                AccountEnabled =$AccountStatus
                UsageLocation=$user.usageLocation
                Dept=$user.department
                JobTitle =$user.jobTitle
                CompanyName=$user.companyName
                CreatedDate=$user.createdDateTime.ToString("dd-MM-yyyy hh:mm:ss")
                Skus=$user.assignedLicenses.skuId -join ","
                Services=$user.assignedPlans.service -join ","
                LicensePlan=$user.assignedPlans.servicePlanId -join ","
                LastInteractiveSignIn=$LastInteractiveSignIn  
                LastNonInteractiveSignIn= $LastNonInteractiveSignIn
                InactiveDays=$InactiveDays_InteractiveSignIn
                InactiveDays_NoInteractiveSignIn =$InactiveDays_NonInteractiveSignIn   
                RefreshTokenValidFrom=$user.refreshTokensValidFromDateTime
                AssignedByGroup=$user.licenseAssignmentStates.assignedByGroup -join ","
         
                }   
                $report.Add($obj) 
            } #end for
            $Uri = $users.'@odata.nextLink'         
        } until (-not $Uri)
        if ($report -eq $null)
        {
            Write-Output "No licensed users found in Tenant"
                $obj = [pscustomObject][ordered] @{
                No=$null
                Id =$null
                DisplayName = $null
                UserPrincipalName =$null
                UserType=$null
                AccountEnabled =$null
                UsageLocation=$null
                Dept=$null
                JobTitle =$null
                CompanyName=$null
                CreatedDate=$null
                Skus=$null
                Services=$null
                LicensePlan=$null
                LastInteractiveSignIn=$null
                LastNonInteractiveSignIn= $null
                InactiveDays=$null
                InactiveDays_NoInteractiveSignIn =$null
                RefreshTokenValidFrom =$null
                AssignedByGroup=$null
                }
            $report.Add($obj) 
            
        }
        $filePath = $env:Temp
        Write-Output "Writing file:" $CSVFileName
        $report | Export-Csv -Path $filePath\$CSVFileName -NoTypeInformation
        WritetoSharePoint($CSVFileName)
    }
    catch
    {
        Write-Output $Error[0]
    }
} #function

<#5. Fetch activity reports#>
function ExecuteReportAPI
{ 
    param([String[]] $Param )
    $CSVFileName = $Param[0]
    $graphApiUri = $Param[1]
    $filePath = $env:Temp
try{
                
        Write-Output $CSVFileName 
         $global:authHeader["ContentType"] = "text/csv"        
        $Reports = Invoke-RestMethod -Method Get -Uri $graphApiUri -Headers $global:authHeader | ConvertFrom-Csv
        $Reports | Export-Csv $filePath\$CSVFileName -NoTypeInformation
        $Values = @{"Title" = $CSVFileName}
        # Add the file to the Reports folder
        Write-Output "$CSVFileName"
        Write-Output $Reports
               WritetoSharePoint($CSVFileName)
       
    }
    catch 
    { Write-Output "Exception while executing API for report"
      Write-Error $Error[0]
             $Error=$null}

}

<#6. Get Unlicensed Users #>
Function GetUnlicensedUsers 
{
param ( $CSVFileName)
$users =$null
try
{
    
        $Uri="https://graph.microsoft.com/beta/users?`$filter=accountEnabled eq true and assignedLicenses/`$count ne 0 &`$count=true&`$select=id,displayName,createdDateTime,userPrincipalName,jobTitle,department,assignedLicenses,assignedPlans,accountEnabled,companyName,employeeType,officeLocation,signInActivity,usageLocation&`$top=999" 
        $report = [System.Collections.Generic.List[Object]]::new()
        $rownum =0
        Do
        {
            $users =  Invoke-RestMethod -Uri $Uri -Method Get -Headers $global:authHeader
            $Uri = $users.'@odata.nextLink'
            foreach ($user in $users.Value)
            {
                $UPN=$user.userPrincipalName
                $AccountEnabled =$user.accountEnabled
                $rownum++
                $logindate =$user.signInActivity.lastSignInDateTime
                    if ($logindate -eq $null) 
                {
                    $LastInteractiveSignIn = $null
                    $InactiveDays_InteractiveSignIn =$null
                }
                else
                {
                
                $InactiveDays_InteractiveSignIn = (New-TimeSpan -Start $logindate).Days
                $LastInteractiveSignIn =$logindate.ToString("dd-MM-yyyy hh:mm:ss")
                }
                  $logindate =$user.signInActivity.lastNonInteractiveSignInDateTime
                if ($logindate -eq $null)
                    {
                            $LastNonInteractiveSignIn = $null
                            $InactiveDays_NonInteractiveSignIn = $null
                    }
                else
                {
                            $LastNonInteractiveSignIn =$logindate.ToString("dd-MM-yyyy hh:mm:ss")
                            $InactiveDays_NonInteractiveSignIn = (New-TimeSpan -Start $logindate).Days
                }
                if ($UPN -notmatch '#EXT#')
                {
                $Type = "Member"
                }
                else
                {
                $Type="External User"
                }
                
                if($AccountEnabled -eq $true)
                {
                    $AccountStatus='Enabled'
                }
                else
                {
                    $AccountStatus='Disabled'
                }
             $obj = [pscustomObject][ordered] @{
                No=$rownum
                Id =$user.id
                DisplayName = $user.displayName
                UserPrincipalName =$user.userPrincipalName
                UserType=$Type
                AccountEnabled =$AccountStatus
                UsageLocation=$user.usageLocation
                Dept=$user.department
                JobTitle =$user.jobTitle
                CreatedDate=$user.createdDateTime.ToString("dd-MM-yyyy hh:mm:ss")
                LastInteractiveSignIn=$LastInteractiveSignIn  
                LastNonInteractiveSignIn= $LastNonInteractiveSignIn
                InactiveDays=$InactiveDays_InteractiveSignIn
                InactiveDays_NoInteractiveSignIn =$InactiveDays_NonInteractiveSignIn            
                }   
                $report.Add($obj) 
            } #end for             
        } until (-not $Uri)
        if ($report -eq $null)
        {
            Write-Output "No unlicensed users found in Tenant"
                $obj = [pscustomObject][ordered] @{
                No=$null
                Id =$null
                DisplayName = $null
                UserPrincipalName =$null
                UserType=$null
                AccountEnabled =$null
                UsageLocation=$null
                Dept=$null
                JobTitle =$null
                CreatedDate=$null
                LastInteractiveSignIn=$null
                LastNonInteractiveSignIn= $null
                InactiveDays=$null
                InactiveDays_NoInteractiveSignIn =$null
                }
            $report.Add($obj) 
            
        }
        $filePath = $env:Temp
        Write-Output "Writing file:" $CSVFileName
        $report | Export-Csv -Path $filePath\$CSVFileName -NoTypeInformation
        WritetoSharePoint($CSVFileName)

 } #end try
 catch {
         Write-Error $Error[0]
         $Error=$null
         Write-Output "Error creating report"
        }
}


#7. Output results to SharePoint folder
function WritetoSharePoint{
   param( $CSVFileName )
     $filePath = $env:Temp
  
    # Upload to SharePoint
    $FolderObject = Get-PnPFolder -Url $DestinationURL
    Write-Output $FolderObject
    $Upload= Add-PnPFile -Path $filePath\$CSVFileName  -Folder $FolderObject -Verbose
    If ($Upload -ne $null)
    {
         Write-Output $CSVFileName " Report sucessfully uploaded"
    }
    # Clean up local file
    Remove-Item -Path $filePath\$CSVFileName -Force
}

#8.Fetch subscriptions
function GetSubscriptions
{
    param($CSVFileName)
    Write-Output "Get subscriptions"
    $filePath =$env:Temp
    $Uri = "https://graph.microsoft.com/v1.0/subscribedSkus"
   [array]$SkuData = Invoke-RestMethod -Uri $Uri -Method Get -Headers $global:authHeader 
    $report = [System.Collections.Generic.List[Object]]::new()
   
    Foreach( $Sku in $SkuData.Value)
    {
         $SkuRenewalDate = $Sku.nextLifecycleDateTime
         $obj = [pscustomObject][ordered] @{
            SkuId= $Sku.SkuId
            Id= $Sku.id
            Trial =$Sku.isTrial
            NextLifecycleDateTime = $SkuRenewalDate
            SkuPartNumber=$Sku.skuPartNumber
            TotalLicenses= $Sku.totalLicenses 
            ServicePlanId=$Sku.serviceStatus.servicePlanId   -join ","  
            ServicePlans=$Sku.serviceStatus.servicePlanName  -join ","
            OwnerId = $Sku.ownerId
         }
      
         $report.add($obj)
    }
        $report | Export-Csv $filePath\$CSVFileName -NoTypeInformation
        $Values = @{"Title" = $CSVFileName}
        # Add the file to the Reports folder
        WritetoSharePoint($CSVFileName)
   
}

 #Execute the functions as needed.
   Function main()
   {
    ConnectToGraph
    ConnectToSharePoint
    GetServicePlans("1.ServicePlans.csv")
    #Uses graph API to fetch the License Expiry details
    GetSubscriptions("1.Subscriptions.csv")
    GetLicensedUsers("2.LicensedUsers.csv")
    GetUnlicensedUsers("3.UnlicensedUsers.csv")
    ExecuteReportAPI("4.ActiveUserDetail_30.csv","https://graph.microsoft.com/v1.0/reports/getOffice365ActiveUserDetail(period='D30')")
    #Optional
    ExecuteReportAPI("4.ActiveUserCounts_30.csv","https://graph.microsoft.com/v1.0/reports/getOffice365ActiveUserCounts(period='D30')")
    ExecuteReportAPI("4.M365AppUsageDetail_30.csv","https://graph.microsoft.com/v1.0/reports/getM365AppUserDetail(period='D30')")
    #Optional - not used in report
    ExecuteReportAPI("4.O365ActivationUserDetail.csv","https://graph.microsoft.com/v1.0/reports/getOffice365ActivationsUserDetail")
    #Optional - not used in report
    ExecuteReportAPI("4.O365ActivationCounts.csv","https://graph.microsoft.com/v1.0/reports/getOffice365ActivationCounts")
}
main