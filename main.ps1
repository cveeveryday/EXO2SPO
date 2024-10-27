Import-Module ..\EXOFunctions\exofunctions.psm1
Import-Module ..\SPFunctions\spofunctions.psm1
Import-Module ..\afterhoursService\ahservicefunctions.psm1

.\variables.ps1

$token = Get-GraphToken -appID $appID -clientSecret $clientSecret -tenantID $tenantID
<#
New-SPListFromObject -token $token -siteName "afterhours" -listName "Patty's Emails" #-colunmns @("From","To","DateReceived","Subject","Body","isRead") 
Update-SPListColumnName -token $token -siteName "afterhours" -listName "PattysEmails" -OldColumnName "Title" -NewColumnName "MessageId"
Add-SPListColumn  -token $token  -siteName "afterhours"  -listName "PattysEmails"  -ColumnName "Subject"  -ColumnType "Text"
Add-SPListColumn -token $token -siteName "afterhours" -listName "PattysEmails" -ColumnName "Sender" -ColumnType "Text"
Add-SPListColumn -token $token -siteName "afterhours" -listName "PattysEmails" -ColumnName "To" -ColumnType "Text"
Add-SPListColumn -token $token -siteName "afterhours" -listName "PattysEmails" -ColumnName "Body" -ColumnType "multilineText"
Add-SPListColumn  -token $token  -siteName "afterhours"  -listName "PattysEmails"  -ColumnName "DateReceived"  -ColumnType "dateTime"
Add-SPListColumn -token $token  -siteName "afterhours"  -listName "PattysEmails"  -ColumnName "Status"  -ColumnType "Text"

$list = Get-SPLists  -token $token  -siteName "afterhours"  -listName "PattysEmails"
#$items = Get-SPListItems -accessToken $token -siteName "afterhours" -listName "PattysEmails" -ColumnName "Status" -filter "In Progress"


New-SPListFromObject -token $token -siteName "afterhours" -listName "staffEdmonton" #-colunmns @("From","To","DateReceived","Subject","Body","isRead") 
Update-SPListColumnName -token $token -siteName "afterhours" -listName "staffEdmonton" -OldColumnName "Title" -NewColumnName "PhoneNumber"
Add-SPListColumn  -token $token  -siteName "afterhours"  -listName "staffEdmonton" -ColumnName "Threshold" -ColumnType "Number"
Add-SPListColumn  -token $token  -siteName "afterhours"  -listName "staffEdmonton" -ColumnName "Frequency" -ColumnType "Number"
Add-SPListColumn  -token $token  -siteName "afterhours"  -listName "staffEdmonton" -ColumnName "Employee" -ColumnType "Text"
Add-SPListColumn -token $token -siteName "afterhours" -listName "staffEdmonton" -ColumnName "Enabled" -ColumnType "Boolean"
#>

$message = Get-OldestVoiceMailMessage -accessToken $token -emailAddress "PattiF@zpzbx.onmicrosoft.com"
if ($message) {
  $onCallStaff = Get-SPListItems -accessToken $token -siteName "afterhours" -listName "staffEdmonton"
  If ($onCallStaff)
  {
    $onCallStaffDetails = @()
    $onCallStaff | ForEach-Object {
      $onCallStaffDetails += (Get-SPListItem -accessToken $token -siteName "afterhours" -listName "staffEdmonton" -itemId $_.id).fields
    }
    $staffToCall = Get-onCallStaff -emailReceivedDate $message.receivedDateTime -oncallStaff $onCallStaffDetails
    if ($staffToCall) {
      $staffToCall | ForEach-Object {
        Write-Host 'Calling ' $_.Employee ' with phone number ' $_.Title
      }
    }
  }
}


<#
$folderId = (Get-MailFolder -accessToken $token -emailAddress "PattiF@zpzbx.onmicrosoft.com" -folderName "Inbox").id
$emailsAll = Get-MailMessages -accessToken $token -emailAddress "PattiF@zpzbx.onmicrosoft.com"  -folderId $folderId -limit 199
$emailsProcessed = @()
$emailsStaged = @()
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

Add-ObjectToSPList -token $token -siteName "afterhours" -listName "PattysEmails" -objects $emailsProcessed

$emailsStaged = Get-SPListItems -accessToken $token   -siteName "afterhours"  -listName "PattysEmails" -ColumnName "Status"  -filter "Staged"

foreach ($email in $emailsStaged) {
  $email = Get-SPListItem -accessToken $token -siteName "afterhours" -listName "PattysEmails" -itemId $email.ID
  $newMessageId = Move-MailMessage -accessToken $token -emailAddress "PattiF@zpzbx.onmicrosoft.com" -folderName "In Progress" -messageId $email.fields.Title
  If ($newMessageId)
  {
    $result = Update-SPListItems -accessToken $token -siteName "afterhours" -listName "PattysEmails" -ColumnNameFilter "Title" -FilterValue $email.fields.Title  -ColumnName "Title" -newValue $newMessageId
    $result = Update-SPListItems -accessToken $token -siteName "afterhours" -listName "PattysEmails" -ColumnNameFilter "Title" -FilterValue $email.fields.Title  -ColumnName "Status" -newValue "Processed"
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