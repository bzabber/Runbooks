####################################################################################################
#- Name: O36MailFlowGraphApiTest.ps1                                                              -#
#- Date: March 4, 2021                                                                            -#
#- Description: This script will leverage Graph API to send an email then validate the email      -# 
#-              was recieved and feed results into Log Analytics                                  -#
#- Dependencies:                                                                                  -#
#- 	- MSAL.PS PowerShell Module Library (Available from PowerShell Gallery)                       -#
#-  - OMSIngestionAPI Module Library (Available from PowerShell Gallery)                          -#  
#- 	- User credentials for Sender and Reciever                                                    -#
#- 	- Log analytics workspace and key                                                             -#
#-  - AAD App Registration with the following API Permissions                                     -#
#- 	  -  Mail.Read                                                                                -#
#- 	  -  Mail.ReadBasic                                                                           -#
#- 	  -  Mail.ReadWrite                                                                           -#
#- 	  -  Mail.Send                                                                                -#
#- 	  -  Mail.Send.Shared                                                                         -#
#- 	                                                                                              -#
####################################################################################################

Import-Module MSAL.PS
Import-Module OMSIngestionAPI

#--- Get AD Application info from variables ---#
$clientId = Get-AutomationVariable -Name 'clientId'
$tenantId = Get-AutomationVariable -Name 'tenantId'
$redirectUri = Get-AutomationVariable -Name 'redirectUri'

#--- Get Log Analytics authentication info from variables ---#
$workspaceId = Get-AutomationVariable -Name 'OMSWorkSpaceID'
$workspaceKey = Get-AutomationVariable -Name 'OMSPrimaryKey'

$LogType = "O365SyntheticGraphAPI"
$TestID = (get-date -format MMddyyyyhhmmss) + "-" + (get-random)

$SenderCredential = Get-AutomationPSCredential -Name 'ExoMailFlowSender'
$ReceiverCredential = Get-AutomationPSCredential -Name 'ExoMailFlowReceiver'
$RecieverEmail = $ReceiverCredential.UserName


# Send email to reciever
$token = Get-MsalToken -ClientId $clientId -TenantId $tenantId -RedirectUri $redirectUri -UserCredential $SenderCredential
$accessToken = $token.AccessToken

$header = @{"Authorization" = "Bearer $accessToken"; "Content-Type" = "application/json" };
$sendMailMessageUrl = "https://graph.microsoft.com/v1.0/me/sendMail"

# Create Message body
$JSON = @"
{
    "message": {
      "subject": "Test Email $TestID",
      "body": {
        "contentType": "Text",
        "content": "Testing sending of Emails."
      },
      "toRecipients": [
        {
          "emailAddress": {
            "address": "$RecieverEmail"
          }
        }
      ]
    },
    "saveToSentItems": "true"
  }
"@

# Capture Call Metrics
$sw = New-Object Diagnostics.Stopwatch

$sw.Start()

Try {
  Invoke-RestMethod -Method POST -Headers $header -Uri $sendMailMessageUrl -Body $JSON
  $SendTime = $sw.ElapsedMilliseconds
  $SendStatus = "success"
  $TransMsg = ""
} 
Catch {
  $sw.Stop()
  $SendTime = $sw.ElapsedMilliseconds
  $SendStatus = "failure"
  $TransMsg = "$_"
}

$omsjson = @"
[{  "Computer": "$ENV:COMPUTERNAME",
    "TestType": "mailflow",
    "TestID": "$TestID",
    "TransactionType": "sendmessage",
    "TransactionResult": "$SendStatus",
    "TransactionTime": $SendTime,
    "TransactionMessage": "$TransMsg"
}]
"@
write-output $omsjson
# Send Monitoring Data for email
#Send-OMSAPIIngestionFile -customerId $workspaceId -sharedKey $workspaceKey -body $omsjson -logType $logType

if ($SendStatus -eq "failure") {
  Exit
}

# reply to email 
$replyToken = Get-MsalToken -ClientId $clientId -TenantId $tenantId -RedirectUri $redirectUri -UserCredential $ReceiverCredential
$replyAccessToken = $replyToken.AccessToken

# Obtain latest message from Graph API
$getMessageHeader = @{"Authorization" = "Bearer $replyAccessToken"; "Content-Type" = "application/json" };
$getMessageMessageUrl = "https://graph.microsoft.com/v1.0/me/messages?`$search=`"Subject:Test Email $TestID`""

write-output $getMessageMessageUrl

$sw = New-Object Diagnostics.Stopwatch
$sw.Start()

$getMessageReply = Invoke-RestMethod -Method GET -Headers $getMessageHeader -Uri $getMessageMessageUrl
$sw.Stop()

$ReceiveTime = $sw.ElapsedMilliseconds

$messageCount = $getMessageReply.value.Count

if ($messageCount -gt 0) {
  $ReceiveStatus = "success"
  $TransMsg = ""
}
else {
  $ReceiveStatus = "failure"
  $TransMsg = "Message with Subject 'Test Email $TestID' not found"
}

$omsjson = @"
[{   "Computer": "$ENV:COMPUTERNAME",
    "TestType": "mailflow",
    "TestID": "$TestID",
    "TransactionType": "searchformessage",
    "TransactionResult": "$ReceiveStatus",
    "TransactionTime": $ReceiveTime,
    "TransactionMessage": "$TransMsg"
}]
"@
write-output $omsjson
#Send-OMSAPIIngestionFile -customerId $workspaceId -sharedKey $workspaceKey -body $omsjson -logType $logType