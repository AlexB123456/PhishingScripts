# variables
$PhishingInbox = $null

### connect to O365
#
# Install EWS Client from https://www.microsoft.com/en-us/download/details.aspx?id=42951
# 
$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"  
#
# Call the EWS dll
[void][Reflection.Assembly]::LoadFile($dllpath)  
#
#
$service = new-object Microsoft.Exchange.WebServices.Data.ExchangeService 
#
#
$service.Url = new-object Uri("https://outlook.office365.com/EWS/Exchange.asmx"); 
#
# Get your login
Get-Credential | Export-Clixml "c:\tools\Cred.xml" 
#
# Pass your login and convert to a PSCredential
$preCredential = Import-Clixml -Path "C:\tools\Cred.xml" 
$cred = New-Object System.Management.Automation.PSCredential -ArgumentList $preCredential.username, $preCredential.Password 
#
$Service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials($cred) 
#
###

# create Property Set to include body and header of email
$PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)

# Set how many emails we want to read at a time
$pagesize = 100

# Index to keep track of where we are up to. Set to 0 initially. 
$offset = 0 

# Set what we want to retrieve from the folder. This will grab the $pagesize number of emails, starting at $offset. If we have already read 100 emails, offset will indicate that we need to read from 101 to 200.
$view = New-Object Microsoft.Exchange.WebServices.Data.ItemView($pagesize,$offset) 

# Indicate which folder and mailbox we want to read
$folderid = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
 
# Place email identifiers into array so that we can ignore new arrivals while processing through the batch
$PhishingInbox += $service.FindItems($folderid,$view) 

$MailMessage = $PhishingInbox.Items | Where-Object subject -eq "[EXTERNAL] Fwd: Bsides Test"

$MailMessage.Load($PropertySet)


# Mess around with $MailMessage in ISE