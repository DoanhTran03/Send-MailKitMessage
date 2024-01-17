Import-Module "D:\Project\Powershell\Send-MailKitMessage\MailKitMessage.psm1"
$Credential = Import-Clixml "C:\Scripts\Creds\outlook.xml"
Send-MailKitMessage -From "gofortest79@outlook.com" -To "gofortest79@outlook.com" -Subject "Hello" -Body "This is just a test" -SMTPServer "smtp-mail.outlook.com" -Port 587 -Credential $Credential -Attachment "C:\test.png"