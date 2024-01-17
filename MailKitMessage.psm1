Add-Type -Path "D:\Package\MailKit\lib\netstandard2.0\MailKit.dll"
Add-Type -Path "D:\Package\MimeKit\lib\netstandard2.0\MimeKit.dll"

function Send-MailKitMessage {
    param(
        [Parameter(Mandatory)][String] $From,
        [Parameter(Mandatory)][String] $To,
        [Parameter()]
        [Parameter()] $CC,
        [Parameter()][String] $Subject = "",
        [Parameter()][String] $Body = "",
        [Parameter(Mandatory)][string] $SMTPServer,
        [Parameter()][Int16] $Port, 
        [Parameter()]$Credential
    )
    
    $Message = New-Object MimeKit.MimeMessage
    $SMTP = New-Object MailKit.Net.Smtp.SmtpClient

    $MailboxAddress = New-Object MimeKit.MailboxAddress("Ngoc Doanh", "N/A")
    $Message.From.Add($MailboxAddress::Parse($From))
    $Message.To.Add($MailboxAddress::Parse($To))
    
    if ($CC) {
        foreach ($address in $CC) {
            $Message.Cc.Add($address)
        }
    }

    $Message.Subject = $Subject
    $Message.Body = $Body

    #Create SMTP Object and add make network connection
    $Option = New-Object MailKit.Security.SecureSocketOptions

    $SMTP = New-Object MailKit.Net.Smtp.SmtpClient
    $SMTP.Connect($SMTPServer, $Port, $Option::StartTls)
    $SMTP.Authenticate($Credential)
    $SMTP.Send($Message)
    $SMTP.Disconnect($true) 
}