Import-Module ..\EXOFunctions\exofunctions.psm1
Import-Module ..\SPFunctions\spofunctions.psm1

.\variables.ps1

$token = Get-GraphToken -appID $appID -clientSecret $clientSecret -tenantID $tenantID

New-SPListFromObject -token $token -siteName "Team Site" -listName "Patty's Emails" #-colunmns @("From","To","DateReceived","Subject","Body","isRead") 
Update-SPListColumnName -token $token -siteName "Team Site" -listName "PattysEmails" -OldColumnName "Title" -NewColumnName "Subject"
Add-SPListColumn -token $token -siteName "Team Site" -listName "PattysEmails" -ColumnName "Sender" -ColumnType "Text"
Add-SPListColumn -token $token -siteName "Team Site" -listName "PattysEmails" -ColumnName "To" -ColumnType "Text"
Add-SPListColumn -token $token -siteName "Team Site" -listName "PattysEmails" -ColumnName "Body" -ColumnType "multilineText"
Add-SPListColumn  -token $token  -siteName "Team Site"  -listName "PattysEmails"  -ColumnName "DateReceived"  -ColumnType "dateTime"
Add-SPListColumn  -token $token  -siteName "Team Site"  -listName "PattysEmails"  -ColumnName "messageid"  -ColumnType "Text"

$list = Get-SPLists  -token $token  -siteName "Team Site"  -listName "PattysEmails"

$folderId = (Get-MailFolder -accessToken $token -emailAddress "PattiF@zpzbx.onmicrosoft.com" -folderName "Inbox").id
$emailsAll = Get-MailMessages -accessToken $token -emailAddress "PattiF@zpzbx.onmicrosoft.com"  -folderId $folderId -limit 10
$emailsProcessed = @()
foreach ($email in $emailsAll) {
  $emailsProcessed += [PSCustomObject]@{
    messageid = $email.id
    Title = ConvertFrom-Html -inputString $email.subject
    Sender = $email.sender.emailAddress.address
    To = $email.toRecipients.emailAddress.address
    DateReceived = $email.receivedDateTime
    Body = ConvertFrom-Html -inputString $email.body.content
  }
  


}

Add-ObjectToSPList -token $token -siteName "Team Site" -listName "PattysEmails" -objects $emailsProcessed
foreach ($email in $emailsProcessed) {
  $email.messageId = Move-MailMessage -accessToken $token -emailAddress "PattiF@zpzbx.onmicrosoft.com" -folderName "Archive" -messageId $email.messageid
  $email.messageId = Set-MailMessageAsRead -accessToken $token -emailAddress "PattiF@zpzbx.onmicrosoft.com" -messageId $email.messageid
  $email.messageId
}
<#
$folderId = (Get-MailFolder -accessToken $token -emailAddress "PattiF@zpzbx.onmicrosoft.com" -folderName "Archive").id
$emailsAll = Get-MailMessages -accessToken $token -emailAddress "PattiF@zpzbx.onmicrosoft.com"  -folderId $folderId -limit 0
foreach ($email in $emailsAll) {
  $email.Id = Move-MailMessage -accessToken $token -emailAddress "PattiF@zpzbx.onmicrosoft.com" -folderName "Inbox" -messageId $email.id
  #$email.messageId = Set-MailMessageAsRead -accessToken $token -emailAddress "PattiF@zpzbx.onmicrosoft.com" -messageId $email.messageid
  $email.Id
}
#>