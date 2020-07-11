#Author: Prabhat Nigam
#Microsoft Architect and CTO @ Golden Five Consulting
#Date: 08/10/2017
#Description: Extract Password Never Expires enabled user list and email to the configured user.
#Disclaimer: Please use as your own risk
#Please update SMTPHost, From,To, ReportPath, Subjet, $Body#

$smtphost = "SMTPHOST IP"
$from = "sender email"
$to = "recepient email"
$Reportpath = "c:\temp\PNE-$((Get-Date).ToString('MM-dd-yyyy')).csv"
Import-Module ActiveDirectory
get-aduser -filter * -properties * | where { $_.passwordNeverExpires -eq "true" } | where {$_.enabled -eq "true"} | Select Name,DistinguishedName,PasswordNeverExpires,@{Name="Lastlogon"; Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}} | export-csv $Reportpath -force -NoTypeInformation
$subject = "Password Never Expire Report" 
$body = "Please find the attached Password Never Expires marked user report."
$smtp= New-Object System.Net.Mail.SmtpClient $smtphost 
$msg = New-Object System.Net.Mail.MailMessage $from, $to, $subject, $body 
$msg.isBodyhtml = $true 
$msg.Attachments.Add($Reportpath)
$smtp.send($msg) 
