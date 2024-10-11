Import-Module ..\EXOFunctions\exofunctions.psm1
Import-Module ..\SPFunctions\spofunctions.psm1

.\variables.ps1

$token = Get-GraphToken -appID $appID -clientSecret $clientSecret -tenantID $tenantID
<#
New-SPListFromObject -token $token -siteName "Team Site" -listName "Patty's Emails" #-colunmns @("From","To","DateReceived","Subject","Body","isRead") 
Update-SPListColumnName -token $token -siteName "Team Site" -listName "PattysEmails" -OldColumnName "Title" -NewColumnName "MessageId"
Add-SPListColumn  -token $token  -siteName "Team Site"  -listName "PattysEmails"  -ColumnName "Subject"  -ColumnType "Text"
Add-SPListColumn -token $token -siteName "Team Site" -listName "PattysEmails" -ColumnName "Sender" -ColumnType "Text"
Add-SPListColumn -token $token -siteName "Team Site" -listName "PattysEmails" -ColumnName "To" -ColumnType "Text"
Add-SPListColumn -token $token -siteName "Team Site" -listName "PattysEmails" -ColumnName "Body" -ColumnType "multilineText"
Add-SPListColumn  -token $token  -siteName "Team Site"  -listName "PattysEmails"  -ColumnName "DateReceived"  -ColumnType "dateTime"
Add-SPListColumn -token $token  -siteName "Team Site"  -listName "PattysEmails"  -ColumnName "Status"  -ColumnType "Text"
#>
$list = Get-SPLists  -token $token  -siteName "Team Site"  -listName "PattysEmails"
#$items = Get-SPListItems -accessToken $token -siteName "Team Site" -listName "PattysEmails" -ColumnName "Status" -filter "In Progress"
$folderId = (Get-MailFolder -accessToken $token -emailAddress "PattiF@zpzbx.onmicrosoft.com" -folderName "Inbox").id
$emailsAll = Get-MailMessages -accessToken $token -emailAddress "PattiF@zpzbx.onmicrosoft.com"  -folderId $folderId -limit 10
$emailsProcessed = @()
foreach ($email in $emailsAll) {
  $emailsProcessed += [PSCustomObject]@{
    Subject = ConvertFrom-Html $email.subject
    Title = $email.id
    Sender = $email.sender.emailAddress.address
    To = $email.toRecipients.emailAddress.address
    DateReceived = $email.receivedDateTime
    Body = ConvertFrom-Html -inputString $email.body.content
    Status = "Staged"
  }
}

Add-ObjectToSPList -token $token -siteName "Team Site" -listName "PattysEmails" -objects $emailsProcessed

$emailsStaged = Get-SPListItems -accessToken $token   -siteName "Team Site"  -listName "PattysEmails"   -ColumnName "Status"  -filter "Staged"

foreach ($email in $emailsStaged) {
  $email = Get-SPListItem -accessToken $token -siteName "Team Site" -listName "PattysEmails" -itemId $email.ID
  $newMessageId = Move-MailMessage -accessToken $token -emailAddress "PattiF@zpzbx.onmicrosoft.com" -folderName "In Progress" -messageId $email.fields.Title
  If ($newMessageId)
  {
    $update = Update-SPListItems -accessToken $token -siteName "Team Site" -listName "PattysEmails" -ColumnName "Title" -filter $email.fields.Title -newValue $newMessageId
  }
  $update
}

<#
$folderId = (Get-MailFolder -accessToken $token -emailAddress "PattiF@zpzbx.onmicrosoft.com" -folderName "Archive").id
$emailsAll = Get-MailMessages -accessToken $token -emailAddress "PattiF@zpzbx.onmicrosoft.com"  -folderId $folderId -limit 0 -isRead $true
foreach ($email in $emailsAll) {
  $email.Id = Move-MailMessage -accessToken $token -emailAddress "PattiF@zpzbx.onmicrosoft.com" -folderName "Inbox" -messageId $email.id
  $email.id = Set-MailMessageAsRead -accessToken $token -emailAddress "PattiF@zpzbx.onmicrosoft.com" -messageId $email.id
}
#>