#       Author: Prabhat Nigam. 
#       Twitter: @PrabhatNigamXHG
#       Blog: MSExchangeGuru.com
#       Email: Prabhat.Nigam@GoldenFive.net
#       Website: www.GoldenFiveConsulting.com
#       Date: 02/20/2016.
#       Description: Change the IP of the existing DNS host record.
#       Version: 1.0
#       Disclaimer: Use it an your own risk. I have tested and used it in my lab and production but every setup is different so it is recommended to test the script in your lab before using it.

$zonename = Read-host "Please enter DNS zone name"
$Hostname = Read-host "Please enter Host record name"
$oldobj = get-dnsserverresourcerecord -name $Hostname -zonename $zonename -rrtype "A"
$newobj = get-dnsserverresourcerecord -name $Hostname -zonename $zonename -rrtype "A"
$updateip = Read-host "Please enter new IP Address"
$newobj.recorddata.ipv4address=[System.Net.IPAddress]::parse($updateip)
Set-dnsserverresourcerecord -newinputobject $newobj -oldinputobject $oldobj -zonename $zonename -passthru