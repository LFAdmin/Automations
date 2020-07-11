#       Author: Prabhat Nigam. 
#       Twitter: @PrabhatNigamXHG
#       Blog: MSExchangeGuru.com
#       Email: Prabhat.Nigam@GoldenFive.net
#       Website: www.GoldenFiveConsulting.com
#       Date: 09/06/2016.
#       Description: All Mailbox Size Export from Exchange 2010/2013/2016.
#       Version: 1.0
#       Disclaimer: Use it an your own risk. I have tested and used it in my lab and production but every setup is different so it is recommended to test the script in your lab before using it.

Get-Mailbox -Resultsize Unlimited | Get-MailboxStatistics | select DisplayName,Database,Totalitemsize,Totaldeleteditemsize | export-csv c:\temp\MailboxSize-$((Get-Date).ToString('MM-dd-yyyy')).csv -NoTypeInformation
