
#Variables
$MailboxAddress = "phishing@xyz.com"

#Service-now email variables, match them to your Incident fields
$SNMailbox = "xyz@service-now.com"
$SMTPServer = "yourinternalsmtprelay.domain.com"
$SNConfigurationItem = "Phishing"
$SNUserId = "x"
$SNPriority = 5 # Low
$SNCategory = "x"
$SNSubcategory = "x"
$SNAssignmentGroup = "x"



##
# Main
##


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
# Get login for O365
# Get-credential gets an encrypted version of the password and only the original userid has the key to convert it back to plain text
# https://social.technet.microsoft.com/wiki/contents/articles/4546.working-with-passwords-secure-strings-and-credentials-in-windows-powershell.aspx
# $cred = Get-Credential -UserName "xyz@xyz.com" -Message "Please enter your own email and password for auth to shared mailbox. This prompt creates an encrypted PSCredential for O365."

Get-Credential | Export-Clixml "c:\tools\Cred.xml" 
#
$preCredential = Import-Clixml -Path "C:\tools\Cred.xml" 
#
$cred = New-Object System.Management.Automation.PSCredential -ArgumentList $preCredential.username, $preCredential.Password 
#
#
###

# Alternative auth with Cyberark AIM
# Import Cyberark Get-Password function
#. D:\Automation\Scripts\PhishingMailbox\Get-Password.ps1
#Get password 
#$preCredential = Import-Clixml -Path "D:\Automation\Scripts\PhishingMailbox\FilterSpam.xml"
#$CyberarkPassword = Get-Password -AppID "App_Security_Phishing" -Safe "Incident Response" -Account "svc_phishingtriage" 
#$CyberarkPassword = $CyberarkPassword | ConvertTo-SecureString -AsPlainText -Force
#
##$cred = New-Object System.Management.Automation.PSCredential -ArgumentList $preCredential.username, $preCredential.Password
#$cred = New-Object System.Management.Automation.PSCredential -ArgumentList "svc_phishingtriage@williams.com", $CyberarkPassword

#These are your O365 credentials
$Service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials($cred)

# create Property Set to include body and header of email
$PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)

# Set how many emails we want to read at a time
$pagesize = 100

# Index to keep track of where we are up to. Set to 0 initially. 
$offset = 0 

# Initialize variables
$PhishingInbox = $null

# Connect to Phishing, view up to 30 subfolders from Inbox
$folderid = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $MailboxAddress)
$view = New-Object Microsoft.Exchange.WebServices.Data.FolderView(30)
$PhishingInbox = $service.FindFolders($folderid,$view) 

# Connect to Spam and Triage folders
$SpamFolder = $PhishingInbox | Where-Object DisplayName -eq "2-Spam"
$TriageFolder = $service.FindFolders($folderid,$view) | Where-Object DisplayName -eq "AutoTriage"


$a = $null
$b = $null

# Find items in spam until $UniqueSubjects > 400
do 
{ 
    # Set what we want to retrieve from the folder. This will grab the $pagesize number of emails, starting at $offset. If we have already read 100 emails, offset will indicate that we need to read from 101 to 200.
    $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView($pagesize,$offset) 
    
    # Place email identifiers into array so that we can ignore new arrivals while processing through the batch
    $a = $SpamFolder.FindItems($view) 
    $b += $a | Sort-Object -Property subject | Select-Object subject
    $UniqueSubjects = $b | Get-Unique -AsString

    ## Increment $offset to next block of emails
    $offset += $pagesize
} while ($UniqueSubjects.count -lt 400) # Do/While there are more emails to retrieve


$offset = 0
$view = $null

# Connect Inbox and retrieve list of all emails
do 
{ 
    # Set what we want to retrieve from the folder. This will grab the $pagesize number of emails, starting at $offset. If we have already read 100 emails, offset will indicate that we need to read from 101 to 200.
    $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView($pagesize,$offset) 
    # Indicate which folder and mailbox we want to read
    $folderid = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $MailboxAddress)
     
    # Place email identifiers into array so that we can ignore new arrivals while processing through the batch
    $PhishingInbox += $service.FindItems($folderid,$view) 

    ## Increment $offset to next block of emails
    $offset += $pagesize
} while ($service.MoreAvailable) # Do/While there are more emails to retrieve

# Filter through array and find email IDs with subjects in list

$PhishingInbox | ForEach-Object -Process {
    
    # Pass message to a variable
    $InboxEmail = $null
    $InboxEmail = $_

    # See if message subject is in list
    If ($UniqueSubjects | Where-Object subject -Contains $InboxEmail.Subject){
        
        # Found an email for triage. Document and move
                
        # Ticket
#        Send-MailMessage -Body "CI: $SNConfigurationItem
#User_id: $SNUserId
#Priority: 5
#Category: $SNCategory
#Subcategory: $SNSubcategory
#Group: $SNAssignmentGroup
#State: 6
#Assign: Me
#Close_notes: Reported email matches spam template, moving to Triage folder.
#
#Subject - $InboxEmail.Subject
#Sender - $InboxEmail.From.Name
#Time Received - $InboxEmail.DateTimeReceived" -Subject "*inc: Phishing Mailbox AutoTriage" -From "xyz@xyz.com" -SmtpServer $SMTPServer -To $SNMailbox            
#
#        
        # Move email
        $InboxEmail.Move($TriageFolder.Id)
        }
    


    }



Get-Variable | Remove-Variable -EA 0



