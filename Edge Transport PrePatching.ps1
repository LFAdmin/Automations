#       Author: Prabhat Nigam. 
#       Twitter: @PrabhatNigamXHG
#       Blog: MSExchangeGuru.com
#       Email: Prabhat.Nigam@GoldenFive.net
#       Website: www.GoldenFiveConsulting.com
#       Date: 02/10/2019.
#       Description: Pre Patching or Restart Exchange Role transfer from server A to Server B for Exchange 2013 and 2016 Edge Transport.
#       Version: 1.0


$PS = "EdgeTransportServer"

#Drain active mail queues on the server
Set-ServerComponentState $PS -Component HubTransport -State Draining -Requester Maintenance

#Transport Services Restart
Restart-Service MSExchangeTransport

#To put the server in maintenance mode run:
Set-ServerComponentState $PS -Component ServerWideOffline -State Inactive -Requester Maintenance

#End of script 