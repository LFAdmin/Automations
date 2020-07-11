#       Author: Prabhat Nigam. 
#       Twitter: @PrabhatNigamXHG
#       Blog: MSExchangeGuru.com
#       Email: Prabhat.Nigam@GoldenFive.net
#       Website: www.GoldenFiveConsulting.com
#       Date: 02/10/2019.
#       Description: Post Patching or Restart Exchange Role transfer from server A to Server B for Exchange 2013 and 2016 .
#       Version: 1.0
#       Disclaimer: Use it an your own risk. I have tested and used it in my lab and production but every setup is different so it is recommended to test the script in your lab before using it.
$PS = "EdgeTransportServer"

#Role Back all roles
Set-ServerComponentState $PS -Component ServerWideOffline -State Active -Requester Maintenance

Set-ServerComponentState $PS -Component HubTransport -State Active -Requester Maintenance
Restart-Service MSExchangeTransport

#Change the directory back to temp
CD c:\temp

#End of script