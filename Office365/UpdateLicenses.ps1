# Login to MSonline services
#$username = Read-Host "Enter your email address"
#$securePwd = Read-Host -assecurestring "Please enter your password"
#Create credential object
#$credObject = New-Object System.Management.Automation.PSCredential -ArgumentList $username, $securePwd
#Connect to Office 365
#Connect-MsolService  -Credential $credObject
#Connect-MsolService

# Log File
$date=Get-Date -Format "MM-dd-yyyy.HH.mm"
$companyDomainArray = Get-MsolDomain
$companyDomain = $companyDomainArray[0].Name
$logPath="c:\"
$logFile="UpdateLicense.log.$companyDomain.$date.txt"
$logCombined="$logPath\$logFile"
# Create License Array
$companyLicenses = Get-MsolAccountSku

# Create Log File
New-Item -Path $logPath -Name $logFile -ItemType "file"
Start-Transcript -Append $logCombined
Write-Host "`n-----------------------------------------------------------------------------------------`n"
Write-Host "                   Office 365 License Removal Script"
Write-Host "`n-----------------------------------------------------------------------------------------`n"
 pause
 Write-Host " "

## User Interaction
# User enters in location of User Text File
$userFile = Read-Host -Prompt "Enter location of user text file. i.e. ($HOME\users-upn.txt)"
#$userFile = "C:\UpdateLicense.log.tiderockholdings.com.07-26-2022.17.16.txt"
# User chooses Licenses to Add
$gridAddLicenses = $companyLicenses.AccountSkuId | Out-GridView -Title "Choose licenes to Add" -PassThru
# User chooses licenses to remove
$gridRemoveLicenses = $companyLicenses.AccountSkuId | Out-GridView -Title "Choose licenes to Remove" -PassThru
## End

# Creates a user list array
$usersList = Get-Content $userFile

# Converts to correct Format
$gridAddJoined = ($gridAddLicenses -join ", ")
$gridRemovedJoined = ($gridRemoveLicenses -join ", ")

# Confirms with User if they want to proceed.
# Confirms Information
Write-Warning "The following settings have been choosen:"
Write-Output "Licenses to Add:    $gridAddJoined"
Write-Output "Licenses to Remove: $gridRemovedJoined"
Write-Output "For the users below."
Write-Output "--------------------"
Write-Output $usersList
Write-Host " "
# Confirm Yes or No
Write-Host -nonewline "Do you want to proceed? (Y/N): "
$Response = Read-Host
    IF ($Response -ne "Y")
    {
        Write-Warning "Script Ending"
        Stop-Transcript
        Get-PSSession | Remove-PSSession
    } else {
        Write-Warning "Starting Updates"
        $usersList | ForEach-Object {
            Set-MsolUserLicense -UserPrincipalName $_ -AddLicenses $gridAddLicenses -RemoveLicenses $gridRemoveLicenses
            Write-Output "$_ had $gridAddJoined added and had $gridRemoveJoined removed "
        }
        Stop-Transcript
    }
#Clean up session
Get-PSSession | Remove-PSSession