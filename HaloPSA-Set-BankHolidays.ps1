# Set the Halo connection details
$VaultName = "Your Azure Keyvault Name"
$HaloClientID = Get-AzKeyVaultSecret -VaultName $VaultName -Name "HaloClientID" -AsPlainText
$HaloClientSecret = Get-AzKeyVaultSecret -VaultName $VaultName -Name "HaloClientSecret" -AsPlainText
$HaloURL = Get-AzKeyVaultSecret -VaultName $VaultName -Name "HaloURL" -AsPlainText

# Connect to Halo
Connect-HaloAPI -URL $HaloURL -ClientId $HaloClientID -ClientSecret $HaloClientSecret -Scopes "all"

#Get all UK Bank holidays
$UKBankHolidays = Invoke-RestMethod -method get -uri "https://www.gov.uk/bank-holidays.json"

# Get just the england / wales bank holidays
$EnglandBankHolidays = $UKBankHolidays.'england-and-wales'.events

# Object array should have title and date properties for the bank holidays
$BankHolidays = $EnglandBankHolidays | Where-object {(get-date($_.date)) -ge (Get-Date)}

$Workdays = Get-HaloWorkday

# Loop all Halo Workdays
foreach ($Day in $Workdays){
    # Get the full object
    $Workday = Get-HaloWorkday -WorkdayID $Day.id -IncludeDetails

    # Confirm if the holidays should be added to the workday
    Write-Host "Would you like to add bank holidays to $($Workday.name)"
    $Answer = Read-Host "Enter Y or N"
    if ($Answer -eq "Y"){
        # Parse existing bank holidays from Hudu
        $ExistingDays = $Workday.holidays | foreach-object {get-date($_.date) -format 'yyyy-MM-dd'}

        # Create an array of existing holidays to add to
        [System.Collections.Generic.List[PSCustomObject]]$Holidays = $Workday.holidays

        # Loop through all bank holidays from gov.uk
        foreach ($BankHoliday in $BankHolidays){
            # Check if it is in Hudu
            if ($ExistingDays -notcontains $BankHoliday.date){
                Write-Host "Adding $($BankHoliday.title) - $($BankHoliday.date)"
                # Add the holiday to the array
                $Holidays.add([pscustomobject]@{
                    name = $BankHoliday.title
                    date = $BankHoliday.date
                })
                
            }
        }
        # Create the update object to add the holidays
        $UpdateWorkday = @{
            id = $Workday.id
            holidays = $Holidays
        }

        # Perform the update
        $Null = Set-HaloWorkday -Workday $UpdateWorkday
    }
}
