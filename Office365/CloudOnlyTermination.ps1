#cloudonly account termination
Import-Module ExchangeOnlineManagement
Import-Module MSOnline
Import-Module AzureAD
Connect-ExchangeOnline
Connect-AzureAD
Connect-MsolService

# Connected to MsolService run the command Get-MsolAccountSku
# This will return all the Current LicenseSKUS
# Output will look like this
#AccountSkuId                                 ActiveUnits WarningUnits ConsumedUnits
#------------                                 ----------- ------------ -------------
#contosocom:WINDOWS_STORE               1000000     0            0
#contosocom:ENTERPRISEPACK              1           0            0
#contosocom:FLOW_FREE                   10000       0            16
#contosocom:SPB                         28          0            28
#contosocom:EXCHANGESTANDARD            3           0            3
#contosocom:O365_BUSINESS_PREMIUM       0           0            0
#contosocom:POWER_BI_STANDARD           1000000     0            1
#contosocom:MCOPSTNC                    10000000    0            0
#contosocom:TEAMS_EXPLORATORY           0           100          2
#contosocom:AAD_PREMIUM                 12          0            12
#contosocom:O365_BUSINESS_ESSENTIALS    9           0            9
#contosocom:ATP_ENTERPRISE              1           0            1

$licensesToRemove = "contosocom:SPB,contosocom:EXCHANGESTANDARD,contosocom:O365_BUSINESS_ESSENTIALS,contosocom:AAD_PREMIUM"



Get-Content "$HOME\termination.txt" | ForEach-Object {
    #Disables Account
    Set-AzureADUser -ObjectID $_ -AccountEnabled $false
    #Sets Mailbox to Shared
    Set-Mailbox $_ -Type Shared
    #Removes Licenses
    Set-MsolUserLicense -UserPrincipalName $_ -RemoveLicenses $licensesToRemove
}
#Clean UP Session
Get-PSSession | Remove-PSSession