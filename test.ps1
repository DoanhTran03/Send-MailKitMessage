Import-Module "D:\Project\Powershell\Send-MailKitMessage\MailKitMessage.psm1"
$Credential = Import-Clixml "C:\Scripts\Creds\outlook.xml"
$body = "<h1>Hello</h1>"
Send-MailKitMessage -From "gofortest79@outlook.com" -To "gofortest79@outlook.com" -Subject "Hello" -Body $body -SMTPServer "smtp-mail.outlook.com" -Port 587 -Credential $Credential -Attachment "C:\test.png" -BodyAsHTML