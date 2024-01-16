#Import MailKit and MimeKit package from : https://www.nuget.org/api/v2/
Add-Type -Path "D:\Package\MailKit\lib\netstandard2.0\MailKit.dll"
Add-Type -Path "D:\Package\MimeKit\lib\netstandard2.0\MimeKit.dll"

#Create support object based on Mailkit document : http://www.mimekit.net/docs/html/Introduction.htm
$Message = New-Object MimeKit.MimeMessage
$MailboxAddress = New-Object MimeKit.MailboxAddress("Ngoc Doanh", "N/A")
$Body_Text = New-Object MimeKit.TextPart("Hello this is test content")

#Create message object from Mailkit namespace
$Message.From.Add($MailboxAddress::Parse("gofortest79@outlook.com"))
$Message.To.Add($MailboxAddress::Parse("gofortest79@outlook.com"))
$Message.Subject = "Test Subject"
$Message.Body = $Body_Text
$Message