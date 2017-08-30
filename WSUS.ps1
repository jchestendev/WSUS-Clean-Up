#Region VARIABLES

# WSUS Connection Parameters:
## Change settings below to your situat,ion. ##
# FQDN of the WSUS server
[String]$parentServer = "WSUS Server Address"
# Use secure connection $True or $False
[Boolean]$useSecureConnection = $False
[Int32]$portNumber = 8530
# From address of email notification
[String]$emailFromAddress = "user@company.com"
# To address of email notification
[String]$emailToAddress = "user@company.com"
# Subject of email notification
[String]$emailSubject = "WSUS Cleanup Results"
# Exchange server
[String]$emailMailserver = "EmailServer"


# Cleanup Parameters:
# Delete computers that have not contacted the server in 30 days or more.
#EndRegion VARIABLES

#Region SCRIPT

# Load .NET assembly
[reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | Out-Null

# Connect to WSUS Server
$wsusParent = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($parentServer,$useSecureConnection,$portNumber);

# Log the date first
$DateNow = Get-Date
$Days = "30"
$LSR = $DateNow.AddDays(-$Days)

Get-WsusComputer -ToLastSyncTime $LSR| Select FullDomainName,Make,Model,OSFamily | Export-Csv -Path "C:/Users\Username\Desktop\MachinesRemoved.csv"

# Perform Cleanup
$Body += "$parentServer ($DateNow ) :" | Out-String 
$CleanupManager = $wsusParent.GetCleanupManager();
$CleanupScope = New-Object Microsoft.UpdateServices.Administration.CleanupScope;
$CleanupScope.CleanupObsoleteComputers = $true
$Body += $CleanupManager.PerformCleanup($CleanupScope) | Out-String 
$file = "C:/Users\Username\Desktop\MachinesRemoved.csv"



# Send the results in an email
Send-MailMessage -From $emailFromAddress -To $emailToAddress -Subject $emailSubject -Body $Body -Attachments $file -SmtpServer $emailMailserver -Priority High -dno OnSuccess, OnFailure



#EndRegion SCRIPT
