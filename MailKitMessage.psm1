Add-Type -Path "D:\Package\MailKit\lib\netstandard2.0\MailKit.dll"
Add-Type -Path "D:\Package\MimeKit\lib\netstandard2.0\MimeKit.dll"

function Send-MailKitMessage {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)][String] $From,
        [Parameter(Mandatory)][String] $To,
        [Parameter()]$CC,
        [Parameter()][String] $Subject = "",
        [Parameter()][String] $Body = "",
        [Parameter(Mandatory)][string] $SMTPServer,
        [Parameter(Mandatory)][Int16] $Port, 
        [Parameter(Mandatory)]$Credential
    )
    
    $Message = New-Object MimeKit.MimeMessage
    $SMTP = New-Object MailKit.Net.Smtp.SmtpClient

    #From and To
    $MailboxAddress = New-Object MimeKit.MailboxAddress("Ngoc Doanh", "N/A")
    $Message.From.Add($MailboxAddress::Parse($From))
    $Message.To.Add($MailboxAddress::Parse($To))
    
    #CC
    if ($CC) {
        foreach ($address in $CC) {
            $Message.Cc.Add($address)
        }       
    }

    #Subject
    $Message.Subject = $Subject

    #Body
    $Body_Text = New-Object MimeKit.TextPart
    $Body_Text.Text = $Body
    $Message.Body = $Body_Text

    #Create SMTP Object and add make network connection
    $Option = New-Object MailKit.Security.SecureSocketOptions

    $SMTP = New-Object MailKit.Net.Smtp.SmtpClient
    $SMTP.Connect($SMTPServer, $Port, $Option::StartTls)
    $SMTP.Authenticate($Credential)
    $SMTP.Send($Message)
    $SMTP.Disconnect($true) 
}