# variables
$DateFrom = "6/23/2020 1:00:00 AM"
$Mailbox =  "phishing@xyz.com"

### connect to O365
#
$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"  
#
[void][Reflection.Assembly]::LoadFile($dllpath)  
#
$service = new-object Microsoft.Exchange.WebServices.Data.ExchangeService 
#
$service.Url = new-object Uri("https://outlook.office365.com/EWS/Exchange.asmx"); 
#
#Get-Credential | Export-Clixml "c:\tools\Cred.xml" 
#
$preCredential = Import-Clixml -Path "C:\tools\Cred.xml" 
#
$cred = New-Object System.Management.Automation.PSCredential -ArgumentList $preCredential.username, $preCredential.Password 
#
$Service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials($cred) 
#
###


### Hook inbox + subfolders




# Start counting Emails #
""
""
# Total count
$TotalCount = 0
$PhishingInbox = $null

# Connect to Phishing, view up to 30 subfolders from Inbox
$folderid = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $Mailbox)
$view = New-Object Microsoft.Exchange.WebServices.Data.FolderView(30)
$PhishingInbox = $service.FindFolders($folderid,$view) 


$PhishingInbox  | ForEach-Object -Process {  
    
    #Clear variable for loop
    $SubTotal = $null
    
    # Find items needs to be < Int32 max
    $SubTotal = $_.FindItems(543210) | Where-Object DateTimeReceived -GT $DateFrom
    $_.Displayname + " - " + $SubTotal.count
    $TotalCount += $SubTotal.count  
    }

""
"Total emails received since $DateFrom : $TotalCount"