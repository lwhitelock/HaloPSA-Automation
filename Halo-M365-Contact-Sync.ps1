#### Halo Settings ####
$VaultName = "Your Key Vault"
$HaloClientID = Get-AzKeyVaultSecret -VaultName $VaultName -Name "HaloClientID" -AsPlainText
$HaloClientSecret = Get-AzKeyVaultSecret -VaultName $VaultName -Name "HaloClientSecret" -AsPlainText
$HaloURL = Get-AzKeyVaultSecret -VaultName $VaultName -Name "HaloURL" -AsPlainText


#### M365 Settings ####
#Microsoft Secure Application Model Info
$customerExclude = (Get-AzKeyVaultSecret -vaultName $VaultName -name "customerExclude" -AsPlainText) -split ',' 
$ApplicationId = Get-AzKeyVaultSecret -vaultName $VaultName -name "ApplicationID" -AsPlainText
$ApplicationSecret = (Get-AzKeyVaultSecret -vaultName $VaultName -name "ApplicationSecret").SecretValue
$TenantID = Get-AzKeyVaultSecret -vaultName $VaultName -name "TenantID" -AsPlainText
$RefreshToken = Get-AzKeyVaultSecret -vaultName $VaultName -name "RefreshToken"-AsPlainText
$UPN = Get-AzKeyVaultSecret -vaultName $VaultName -name "UPN" -AsPlainText

########################## End Secrets Management ##########################
#$VerbosePreference = "continue"
#$DebugPreference = "continue"

#### Script Settings ####

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

# Get Dependencies
if (Get-Module -ListAvailable -Name AzureAD.Standard.Preview) {
	Import-Module AzureAD.Standard.Preview 
} else {
	Install-Module AzureAD.Standard.Preview -Force
	Import-Module AzureAD.Standard.Preview
}


if (Get-Module -ListAvailable -Name HaloAPI) {
	Import-Module HaloAPI 
} else {
	Install-Module HaloAPI -Force
	Import-Module HaloAPI
}

if (Get-Module -ListAvailable -Name PartnerCenterLW) {
	Import-Module PartnerCenterLW 
} else {
	Install-Module PartnerCenterLW -Force
	Import-Module PartnerCenterLW
}

# Connect to Halo
Connect-HaloAPI -URL $HaloURL -ClientId $HaloClientID -ClientSecret $HaloClientSecret -Scopes "all"

$HaloCompaniesRaw = Get-HaloClient
$HaloCompanies = ForEach ($Client in $HaloCompaniesRaw) {
    Get-HaloClient -ClientID $Client.id -IncludeDetails
}
$HaloContacts = Get-HaloUser -FullObjects

$RawHaloSites = Get-HaloSite
$HaloSites = foreach ($Site in $RawHaloSites) {
	Get-HaloSite -SiteID $Site.id
}


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
Write-Host "Connecting to Partner Azure AD"
$credential = New-Object System.Management.Automation.PSCredential($ApplicationId, $ApplicationSecret)
$aadGraphToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes 'https://graph.windows.net/.default' -ServicePrincipal -Tenant $tenantID 
$graphToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes 'https://graph.microsoft.com/.default' -ServicePrincipal -Tenant $tenantID 
Connect-AzureAD -AadAccessToken $aadGraphToken.AccessToken -AccountId $UPN -MsAccessToken $graphToken.AccessToken -TenantId $tenantID | Out-Null
$M365Customers = Get-AzureADContract -All:$true
Disconnect-AzureAD

$GlobalContactsToRemove = [System.Collections.ArrayList]@()

foreach ($customer in $M365Customers) {	
	#Check if customer should be excluded
	if (-Not ($customerExclude -contains $customer.DisplayName)) {
		write-host "Connecting to $($customer.Displayname)" -foregroundColor green
		try {
			$CustAadGraphToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes "https://graph.windows.net/.default" -ServicePrincipal -Tenant $customer.CustomerContextId
			$CustGraphToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes "https://graph.microsoft.com/.default" -ServicePrincipal -Tenant $customer.CustomerContextId
			Connect-AzureAD -AadAccessToken $CustAadGraphToken.AccessToken -AccountId $upn -MsAccessToken $CustGraphToken.AccessToken -TenantId $customer.CustomerContextId | out-null
		} catch {
			Write-Error "Failed to get Azure AD Tokens"
			continue
		}
		
		$defaultdomain = $customer.DefaultDomainName
		$customerDomains = (Get-AzureADDomain | Where-Object { $_.IsVerified -eq $True }).Name

		# Let try to match to an Halo company
		# First lets check default domain against azuretenantid in Halo
		$matchedCompany = $HaloCompanies | Where-Object { $_.azure_tenants.azure_tenant_id -eq $customer.CustomerContextId }
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
							Disconnect-AzureAD
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
						azure_tenant_id = $customer.CustomerContextId
						azure_tenant_name = $customer.DisplayName
						details_name = 'Default'
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
				$UsersRaw = Get-AzureADUser -All:$true
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
			$ContactsToCreate = $M365Users | Where-Object { $_.ObjectId -notin $HaloCompanyContacts.azureoid -and $_.UserPrincipalName -notmatch "admin" }
			$existingContacts = $M365Users | Where-Object { $_.ObjectId -in $HaloCompanyContacts.azureoid }
			$contactsToInactiveRaw = $HaloCompanyContacts | Where-Object { $_.azureoid -notin $M365Users.ObjectId -and $_.emailAddress -notin $ContactsToCreate.UserPrincipalName -and (($($_.emailAddress -split "@")[1]) -in $customerDomains) -or ($_.emailAddress -eq "" -and $_.mobilePhone -eq "" -and $_.phone -eq "") }
			$contactsToInactive = foreach ($inactiveContact in $contactsToInactiveRaw) {
				$inactiveContactUDF = $inactiveContact.customfields | Where-Object { $_.name -eq "CFM365SyncKeepActive" }
				if ($inactiveContactUDF.display -ne 'Y') {
					$inactiveContact
				}
			}
			
			Write-Host "Existing Contacts"
			Write-Host "$($existingContacts | Select-Object DisplayName, UserPrincipalName | Out-String)"
			Write-Host "Contacts to be Created"
			Write-Host "$($ContactsToCreate | where-object {$_.UserPrincipalName -notin $HaloCompanyContacts.emailAddress} | Select-Object DisplayName, UserPrincipalName | Out-String)" -ForegroundColor Green
			Write-Host "Contacts to be Paired"
			Write-Host "$($ContactsToCreate | where-object {$_.UserPrincipalName -in $HaloCompanyContacts.emailAddress} | Select-Object DisplayName, UserPrincipalName | Out-String)" -ForegroundColor Yellow
			Write-Host "Contacts to be set inactive"
			Write-Host "$($contactsToInactive | Select-Object firstName, lastName, emailAddress, mobilePhone, phone | Format-Table | out-string)" -ForegroundColor Red

			
			if ($GenerateInactiveReport) {
				foreach ($inactiveContact in $contactsToInactive) {
					$ReturnContact = [PSCustomObject]@{
						'Company'    = $customer.DisplayName
						'First Name' = $inactiveContact.firstname
						'Last Name'  = $inactiveContact.surname
						'Email'      = $inactiveContact.emailAddress
						'Mobile'     = $inactiveContact.mobilenumber2
						'Phone'      = $inactiveContact.phonenumber
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
					# Get the inactive contacts for the company
					$HaloCompanyInactiveContacts = Get-HaloUser -IncludeInactive -FullObjects -client_id $MatchedCompany.id | where-object { $_.inactive -eq $true }

					foreach ($createUser in $ContactsToCreate) {
						# Find the site for the contact
						$ContactSite = $HaloSites | Where-Object { $_.delivery_address.line1 -eq $createUser.StreetAddress -and $_.client_id -eq $matchedCompany.id }
						if (!$ContactSite) {
							$ContactSite = $HaloSites | where-object { $_.id -eq $MatchedCompany.main_site_id }
						}

						# Check if there is a user who just needs azureoid to be set
						$MatchedUnpairedUser = $HaloCompanyContacts | Where-Object { $_.emailAddress -eq $createUser.UserPrincipalName }
						if (($MatchedUnpairedUser | measure-object).count -eq 1) {
							$UpdateBody = @{
								id       = $MatchedUnpairedUser.id
								azureoid = $createUser.ObjectId
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

						# Check if there is an inactive matching user
						$MatchedInactiveUser = $HaloCompanyInactiveContacts | Where-Object { $_.emailAddress -eq $createUser.UserPrincipalName }
						if (($MatchedInactiveUser | Measure-Object).count -eq 1) {
							$ActivateBody = @{
								id       = $MatchedInactiveUser.id
								inactive = $false
								azureoid = $createUser.ObjectId
							}
							try {
								$Result = Set-HaloUser -User $ActivateBody
								Write-Verbose "User Set Active $($createUser.DisplayName))"
								Write-Debug $Result
								Continue
							} catch {
								Write-Error "Error Activating:  $($createUser.DisplayName)"
								Write-Error "$($ActivateBody | convertto-json | out-string)"
								continue
							}
							
							
						}
								
						
						# Get Email Addresses
						$Email2 = ""
						$Email3 = ""
						$aliases = (($createUser.ProxyAddresses | Where-Object { $_ -cnotmatch "SMTP" -and $_ -notmatch ".onmicrosoft.com" }) -replace "SMTP:", " ")
						$AliasCount = ($aliases | measure-object).count
						if ($AliasCount -eq 1) {
							$Email2 = $aliases
						} elseif ($AliasCount -gt 1) {
							$Email2 = $aliases[0]
							$Email3 = $aliases[1]
						}

						# Build the body of the user
						$CreateBody = @{
							companyID     = $matchedCompany.id
							name          = $createUser.DisplayName
							firstName     = $createUser.GivenName
							lastName      = $createUser.Surname

							title         = $createUser.JobTitle
							
							phonenumber   = $createUser.TelephoneNumber
							mobilenumber2 = $createUser.Mobile
							fax           = $createUser.FacsimileTelephoneNumber
							
							site_id       = $ContactSite.id

							emailAddress  = $createUser.UserPrincipalName
							email2        = $Email2
							email3        = $Email3
							azureoid      = $createUser.ObjectId
							
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

		Disconnect-AzureAD

	}		
}


if ($GenerateInactiveReport) {
	$GlobalContactsToRemove | Export-Csv $InactiveReportName
	Write-Host "Report Written to $InactiveReportName"
}
