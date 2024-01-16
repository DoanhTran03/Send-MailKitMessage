Add-Type -Path "D:\Package\MailKit\lib\netstandard2.0\MailKit.dll"
Add-Type -Path "D:\Package\MimeKit\lib\netstandard2.0\MimeKit.dll"

function Send-MailKitMessage {
    param(
        [Parameter(Mandatory)][String] $From,
        [Parameter(Mandatory)][String] $To,
        [Parameter()] $CC,
        [Parameter()][String] $Subject = "",
        [Parameter()][String] $Body = "",
        [Parameter(Mandatory)][string] $SMTPServer,
        [Parameter()][Int16] $Port, 
        [Parameter()]$Credential
    )
}