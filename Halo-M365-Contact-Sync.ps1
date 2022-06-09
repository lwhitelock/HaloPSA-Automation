#### Halo Settings ####
$VaultName = "Your Key Vault"
$HaloClientID = Get-AzKeyVaultSecret -VaultName $VaultName -Name "HaloClientID" -AsPlainText
$HaloClientSecret = Get-AzKeyVaultSecret -VaultName $VaultName -Name "HaloClientSecret" -AsPlainText
$HaloURL = Get-AzKeyVaultSecret -VaultName $VaultName -Name "HaloURL" -AsPlainText


#### M365 Settings ####
#Microsoft Secure Application Model Info
$customerExclude = (Get-AzKeyVaultSecret -vaultName $VaultName -name "customerExclude" -AsPlainText) -split ',' 
$script:ApplicationId = Get-AzKeyVaultSecret -vaultName $VaultName -name "ApplicationID" -AsPlainText
$script:ApplicationSecret = Get-AzKeyVaultSecret -vaultName $VaultName -name "ApplicationSecret" -AsPlainText
$script:TenantID = Get-AzKeyVaultSecret -vaultName $VaultName -name "TenantID" -AsPlainText
$script:RefreshToken = Get-AzKeyVaultSecret -vaultName $VaultName -name "RefreshToken"-AsPlainText
$script:ExchangeRefreshToken = Get-AzKeyVaultSecret -vaultName $VaultName -name "ExchangeRefreshToken"-AsPlainText

########################## End Secrets Management ##########################
#$VerbosePreference = "continue"
#$DebugPreference = "continue"

#### Script Settings ####

# This is the value for the Yes option of the Keep Active custom field CFM365SyncKeepActive.
# THIS CUSTOM FIELD MUST BE ADDED TO USERS. PLEASE SEE BLOG POST.
# If you run the script once in $ListContactChangesOnly = $true and then look at the 'Halo M3565 Contact PowerShell Script Report'
# and find a contact that has the custom field set to keep them active you can see what this should be.
$HaloCustomFieldKeepActiveValue = 2

# Recommended to set this to true on the first run so that you can make sure companies are being mapped correctly and fix any issues.
$CheckMatchesOnly = $false

# Recommended to set this on first run. It will only tell you what the script would have done and not make any changes
$ListContactChangesOnly = $false

# This will enable the generation of a csv report on which items would have been set to inactive.
$GenerateInactiveReport = $true
$InactiveReportName = "C:\Temp\InactiveUsersReport.csv"

# Import only licensed users
$licensedUsersOnly = $true

# Create Users missing in Halo
$CreateUsers = $true

# Set unlicensed users as inactive in Halo. (This can be overriden by setting the M365SyncKeepActive UDF on a contact to Y)
$InactivateUsers = $true

# Set the AzureTenantID in Azure on a Successful match using any other method
$SetHuduAzureID = $true

##########################          Script         ############################

if (Get-Module -ListAvailable -Name HaloAPI) {
	Import-Module HaloAPI 
} else {
	Install-Module HaloAPI -Force
	Import-Module HaloAPI
}


function New-GraphGetRequest ($uri, $tenantid, $scope, $AsApp, $noPagination) {

	if ($scope -eq "ExchangeOnline") { 
		$Headers = $Script:ExchangeAuthheaders
	} else {
		$headers = $Script:Authheaders
	}
	Write-Verbose "Using $($uri) as url"
	$nextURL = $uri
	$ReturnedData = do {
		try {
			$Data = (Invoke-RestMethod -Uri $nextURL -Method GET -Headers $headers -ContentType "application/json; charset=utf-8")
			if ($data.value) { $data.value } else { ($Data) }
			if ($noPagination) { $nextURL = $null } else { $nextURL = $data.'@odata.nextLink' }                
		} catch {
			$Message = ($_.ErrorDetails.Message | ConvertFrom-Json).error.message
			if ($null -eq $Message) { $Message = $($_.Exception.Message) }
			throw $Message
		}
	} until ($null -eq $NextURL)
   
	return $ReturnedData   

}

function Get-GraphToken($tenantid, $scope, $AsApp, $AppID, $refreshToken, $ReturnRefresh) {
	if (!$scope) { $scope = 'https://graph.microsoft.com/.default' }

	$AuthBody = @{
		client_id     = $script:ApplicationId
		client_secret = $script:ApplicationSecret
		scope         = $Scope
		refresh_token = $script:RefreshToken
		grant_type    = "refresh_token"
                    
	}
	if ($asApp -eq $true) {
		$AuthBody = @{
			client_id     = $script:ApplicationId
			client_secret = $script:ApplicationSecret
			scope         = $Scope
			grant_type    = "client_credentials"
		}
	}

	if ($null -ne $AppID -and $null -ne $refreshToken) {
		$AuthBody = @{
			client_id     = $appid
			refresh_token = $RefreshToken
			scope         = $Scope
			grant_type    = "refresh_token"
		}
	}

	if (!$tenantid) { $tenantid = $script:tenantid }
	$AccessToken = (Invoke-RestMethod -Method post -Uri "https://login.microsoftonline.com/$($tenantid)/oauth2/v2.0/token" -Body $Authbody -ErrorAction Stop)
	if ($ReturnRefresh) { $header = $AccessToken } else { $header = @{ Authorization = "Bearer $($AccessToken.access_token)" } }

	return $header
}

function Invoke-HaloReport($Report) {
	# This will check for a Halo report. Create it if it doesn't exist and return the results if it does
	$HaloReportBase = Get-HaloReport -Search $report.name
	$FoundReportCount = ($HaloReportBase | Measure-Object).Count

	if ($FoundReportCount -eq 0) {
		$HaloReportBase = New-HaloReport -Report $report
	} elseif ($FoundReportCount -gt 1) {
		throw "Found more than one report with the name '$($HaloContactReportBase.name)'. Please delete all but one and try again."
	}

	$HaloResults = (Get-HaloReport -ReportID $HaloReportBase.id -LoadReport).report.rows

	return $HaloResults
}

# Connect to Halo
try {

	Connect-HaloAPI -URL $HaloURL -ClientId $HaloClientID -ClientSecret $HaloClientSecret -Scopes "all"

	# Fetch Existing Halo Azure Tenant Mappings
	$HaloAzureConnections = Get-HaloAzureADConnection -Type 2
	$HaloM365CompanyMapping = $HaloAzureConnections | ForEach-Object {
		$AZConnectionID = $_.id
    (Get-HaloAzureADConnection -Type 2 -AzureConnectionID $AZConnectionID -IncludeDetails).mappings_client | Select-Object azure_tenant_id, azure_tenant_name, client_id, client_name
	}

	# Create / Retrieve the Customer Mapping Report
	$HaloCustomerReport = @{
		name                    = "Halo M3565 Customer PowerShell Script Report"
		sql                     = "SELECT DISTINCT Aarea as 'id', aareadesc as 'name', AWebsite as 'website', ssitenum as 'main_site_id' from AREA left join SITE on aarea = sarea where SIsInvoiceSite = 1"
		description             = "This report is used to quickly obtain company information for use with the M365 Mapping Script"
		type                    = 0
		datasource_id           = 0
		canbeaccessedbyallusers = $false
	}

	$HaloCompanies = Invoke-HaloReport -Report $HaloCustomerReport

	# Create / Retrieve the Contact Mapping Report
	$HaloContactReport = @{
		name                    = "Halo M3565 Contact PowerShell Script Report"
		sql                     = "SELECT DISTINCT UAzureOID as 'azureoid',uemail as 'emailAddress',Usite as 'site_id',uusername as 'displayName',uid as 'id',uextn as 'phone',umobile as 'mobilePhone',uinactive as 'inactive',aarea as 'client_id',CFM365SyncKeepActive FROM USERS LEFT JOIN site ON usite = ssitenum LEFT JOIN area ON sarea = aarea"
		description             = "This report is used to quickly obtain Contact / User IDs for use with the M365 Mapping Script"
		type                    = 0
		datasource_id           = 0
		canbeaccessedbyallusers = $false
	}

	$HaloContacts = Invoke-HaloReport -Report $HaloContactReport

	# Create / Retrieve the Site Mapping Report
	$HaloSiteReport = @{
		name                    = "Halo M3565 Sites PowerShell Script Report"
		sql                     = "SELECT DISTINCT ASSiteID as 'id', ASLine1 as 'line1', sarea as 'client_id' FROM ADDRESSSTORE LEFT JOIN SITE on ASSiteID = SSiteNum Where sarea <> ''"
		description             = "This report is used to quickly obtain site information for use with the M365 Mapping Script"
		type                    = 0
		datasource_id           = 0
		canbeaccessedbyallusers = $false
	}

	$HaloSites = Invoke-HaloReport -Report $HaloSiteReport



	# Prepare webAddresses for lookup
	$CompanyWebDomains = foreach ($HaloCompany in $HaloCompanies) {
		if ($null -ne $HaloCompany.website) {
			$website = $HaloCompany.website
			$website = $website -replace 'https://'
			$website = $website -replace 'http://'
			$website = ($website -split '/')[0]
			$website = $website -replace 'www.'
			[PSCustomObject]@{
				companyID = $HaloCompany.id
				website   = $website
			}
		}
	}

	# Prepare contact domains for matching
	$DomainCounts = $HaloContacts | Where-Object { ($_.emailAddress).length -gt 1 } | Select-Object client_id, @{N = 'email'; E = { $($_.emailAddress -split "@")[1] } } | group-object email, client_id | sort-object count -descending

	#Connect to your Azure AD Account.
	$Script:Authheaders = Get-GraphToken -tenantid $script:Tenantid
	# Get Customers
	[System.Collections.Generic.List[PSCustomObject]]$M365Customers = (New-GraphGetRequest -uri "https://graph.microsoft.com/beta/contracts?`$top=999" -tenantid $script:Tenantid) | Select-Object CustomerID, DefaultdomainName, DisplayName | Where-Object -Property DisplayName -NotIn $customerExclude

	$GlobalContactsToRemove = [System.Collections.ArrayList]@()

} catch {
	Write-Error "An error occured during initial data gathering: $_"
	exit 1
}

foreach ($customer in $M365Customers) {	
	write-host "Connecting to $($customer.Displayname)" -foregroundColor green
	try {
		$TenantFilter = $Customer.CustomerId
		$Script:ExchangeAuthHeaders = Get-GraphToken -AppID 'a0c73c16-a7e3-4564-9a95-2bdf47383716' -RefreshToken $script:ExchangeRefreshToken -Scope 'https://outlook.office365.com/.default' -Tenantid $TenantFilter
		$Script:Authheaders = Get-GraphToken -tenantid $TenantFilter
	} catch {
		Write-Error "Failed to Connect to M365"
		continue
	}
		
	$defaultdomain = $customer.defaultDomainName

	$customerDomains = (New-GraphGetRequest -uri "https://graph.microsoft.com/beta/domains" -tenantid $TenantFilter).id

	# Let try to match to an Halo company
	# First lets check default domain against azuretenantid in Halo
	$MappedCompany = $HaloM365CompanyMapping | Where-Object { $_.azure_tenant_id -eq $customer.customerId }
	$matchedCompany = $HaloCompanies | Where-Object { $_.id -eq $MappedCompany.client_id }

	if (($matchedCompany | measure-object).count -ne 1) {
		# Now lets try to match tenant names to company names
		$matchedCompany = $HaloCompanies | Where-Object { $_.name -eq $Customer.DisplayName }
		if (($matchedCompany | measure-object).count -ne 1) {
			# Now lets try to match to the web address set on the company in Halo to default domain
			$matchedWebsite = $CompanyWebDomains | Where-Object { $_.website -eq $defaultdomain }
			if (($matchedWebsite | measure-object).count -eq 1) {
				#Populate matched company
				$matchedCompany = $HaloCompanies | Where-Object { $_.id -eq $matchedWebsite.companyID }
				Write-Host "Matched Default Domain to Website" -ForegroundColor green
			} else {
				# Now to try matching any verified domain to a website
				$matchedWebsite = $CompanyWebDomains | Where-Object { $_.website -in $customerDomains }
				if (($matchedWebsite | measure-object).count -eq 1) {
					$matchedCompany = $HaloCompanies | Where-Object { $_.id -eq $matchedWebsite.companyID }
					Write-Host "Matched a verified domain to website" -ForegroundColor green
				} else {
					# Now try to match on contact domains
					$matchedContactDomains = $DomainCounts | where-object { (($_.name) -split ',')[0] -in $customerDomains }
					$matchedIDs = ($matchedContactDomains.name -split ', ')[1] | Select-Object -unique
					if (($matchedIDs | measure-object).count -eq 1) {
						$matchedCompany = $HaloCompanies | Where-Object { $_.id -eq $matchedIDs }
						Write-Host "Matched a verified domain to contacts domain" -ForegroundColor green
					} else {
						Write-Host "$($Customer.DisplayName) Could not be matched please set the Azure Tenant Id in the Halo company to $($customer.CustomerContextId)" -ForegroundColor red
						continue
					}

				}


			}
				
		} else {
			Write-Host "Matched on Tenant and Customer Name" -ForegroundColor green
		}
			
		if ($SetHuduAzureID -eq $true) {
			$ClientUpdate = @{
				id            = $matchedCompany.id
				azure_tenants = @(@{
						azure_tenant_id   = $customer.CustomerContextId
						azure_tenant_name = $customer.DisplayName
						details_name      = 'Default'
					})
			}
			
			Write-Host "Setting $($M365Asset.company_name) - HaloID: $HaloID - TenantID $TenantID"
			$Null = Set-HaloClient -Client $ClientUpdate
		}
					
	} else {
		Write-Host "Matched on azuretenantid in Halo" -ForegroundColor green
	}
	

	Write-Host "M365 Company: $($Customer.DisplayName) Matched to Halo Company: $($matchedCompany.name)"
		
		
	if ($CheckMatchesOnly -eq $false) {
		try {
			$UsersRaw = New-GraphGetRequest -uri "https://graph.microsoft.com/beta/users" -tenantid $TenantFilter
		} catch {
			Write-Error "Failed to download users"
			continue
		}

		#Grab licensed users		
		if ($licensedUsersOnly -eq $true) {
			$M365Users = $UsersRaw | where-object { $null -ne $_.AssignedLicenses.SkuId } | Sort-Object UserPrincipalName
		} else {
			$M365Users = $UsersRaw 
		}

		$HaloCompanyContacts = $HaloContacts | Where-Object { $_.client_id -eq $matchedCompany.id }
		$ContactsToCreate = $M365Users | Where-Object { $_.id -notin $HaloCompanyContacts.azureoid -and $_.UserPrincipalName -notmatch "admin" }
		$existingContacts = $M365Users | Where-Object { $_.id -in $HaloCompanyContacts.azureoid }
		$contactsToInactiveRaw = $HaloCompanyContacts | Where-Object { $_.azureoid -notin $M365Users.id -and $_.emailAddress -notin $ContactsToCreate.UserPrincipalName -and (($($_.emailAddress -split "@")[1]) -in $customerDomains) -or ($_.emailAddress -eq "" -and $_.mobilePhone -eq "" -and $_.phone -eq "") }
		$contactsToInactive = $contactsToInactiveRaw | where-object {$_.CFM365SyncKeepActive -ne $HaloCustomFieldKeepActiveValue -and $_.inactive -eq $False -and $_.displayName -ne 'General User'}
			
		Write-Host "Existing Contacts"
		Write-Host "$($existingContacts | Select-Object DisplayName, UserPrincipalName | Out-String)"
		Write-Host "Contacts to be Created"
		Write-Host "$($ContactsToCreate | where-object {$_.UserPrincipalName -notin $HaloCompanyContacts.emailAddress} | Select-Object DisplayName, UserPrincipalName | Out-String)" -ForegroundColor Green
		Write-Host "Contacts to be Paired"
		Write-Host "$($ContactsToCreate | where-object {$_.UserPrincipalName -in $HaloCompanyContacts.emailAddress} | Select-Object DisplayName, UserPrincipalName | Out-String)" -ForegroundColor Yellow
		Write-Host "Contacts to be set inactive"
		Write-Host "$($contactsToInactive | Select-Object displayName, emailAddress | Format-Table | out-string)" -ForegroundColor Red

			
		if ($GenerateInactiveReport) {
			foreach ($inactiveContact in $contactsToInactive) {
				$ReturnContact = [PSCustomObject]@{
					'Company'    = $customer.DisplayName
					'Name'       = $inactiveContact.displayName
					'Email'      = $inactiveContact.emailAddress
					'Mobile'     = $inactiveContact.mobilenumber
					'Phone'      = $inactiveContact.phone
				}
				$null = $GlobalContactsToRemove.add($ReturnContact)
			}
		}
			
		# If not in list only mode carry out changes
		if ($ListContactChangesOnly -eq $false) {
			# Inactivate Users
			if ($InactivateUsers -eq $true) {
				foreach ($deactivateUser in $contactsToInactive) {
					$DeactivateBody = @{
						id       = $deactivateUser.id
						inactive = $true
					}
						
					try {
						$Result = Set-HaloUser -User $DeactivateBody
						Write-Verbose "User Set Inactive: $($deactivateUser.firstName) $($deactivateUser.surname)"
						Write-Debug $Result
					} catch {
						Write-Error "Error Inactivating:  $($deactivateUser.firstName) $($deactivateUser.surname)"
						Write-Error "$($DeactivateBody | convertto-json | out-string)"
						continue
					}
						
						
				}
			}

			# Create Users
			if ($CreateUsers -eq $true) {
				foreach ($createUser in $ContactsToCreate) {
					# Find the site for the contact
					$ContactSite = ($HaloSites | Where-Object { $_.line1 -eq $createUser.StreetAddress -and $_.client_id -eq $matchedCompany.id } | Sort-Object id -Descending | Select-Object -first 1).id
					if (!$ContactSite) {
						$ContactSite = $MatchedCompany.main_site_id
					}


					# Check if there is a user who just needs azureoid to be set
					$MatchedUnpairedUser = $HaloCompanyContacts | Where-Object { $_.emailAddress -eq $createUser.UserPrincipalName }
					if (($MatchedUnpairedUser | measure-object).count -eq 1) {
						$UpdateBody = @{
							id       = $MatchedUnpairedUser.id
							azureoid = $createUser.id
							inactive = $false
						}
						try {
							$Result = Set-HaloUser -User $UpdateBody
							Write-Verbose "User Paired to existing user $($createUser.DisplayName)"
							Write-Debug $Result
							Continue
						} catch {
							Write-Error "Error Pairing:  $($createUser.DisplayName)"
							Write-Error "$($UpdateBody | convertto-json | out-string)"
							continue
						}
							
					}

								
						
					# Get Email Addresses
					$Email2 = ""
					$Email3 = ""
					$aliases = (($createUser.ProxyAddresses | Where-Object { $_ -cnotmatch "SMTP" -and $_ -notmatch ".onmicrosoft.com" }) -replace "SMTP:", " ")
					$AliasCount = ($aliases | measure-object).count
					if ($AliasCount -eq 1) {
						$Email2 = $aliases.trim()
					} elseif ($AliasCount -gt 1) {
						$Email2 = $aliases[0].trim()
						$Email3 = $aliases[1].trim()
					}

					# Build the body of the user
					$CreateBody = @{
						companyID     = $matchedCompany.id
						name          = $createUser.DisplayName
						firstName     = $createUser.GivenName
						lastName      = $createUser.Surname

						title         = $createUser.JobTitle
							
						phonenumber   = $createUser.businessPhones[0]
						mobilenumber2 = $createUser.mobilePhone
						fax           = $createUser.faxNumber
							
						site_id       = $ContactSite

						emailAddress  = $createUser.UserPrincipalName
						email2        = $Email2
						email3        = $Email3
						azureoid      = $createUser.id
							
					}

					# Create the user
					try {
						$Result = New-HaloUser -User $CreateBody
						Write-Verbose "User Created: $($createUser.DisplayName)"
						Write-Debug $Result

					} catch {
						Write-Error "Error Creating:  $($createUser.DisplayName)"
						Write-Error "$($CreateBody | convertto-json | out-string)"
						continue
					}
						
						

				}
			}
	
		}

	}
}		


if ($GenerateInactiveReport) {
	$GlobalContactsToRemove | Export-Csv $InactiveReportName
	Write-Host "Report Written to $InactiveReportName"
}
