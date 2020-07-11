#       Author: Prabhat Nigam. 
#       Twitter: @PrabhatNigamXHG
#       Blog: MSExchangeGuru.com
#       Email: Prabhat.Nigam@GoldenFive.net
#       Website: www.GoldenFiveConsulting.com
#       Date: 09/03/2016.
#       Description: Pre Patching or Restart Exchange Role transfer from server A to Server B for Exchange 2013 and 2016 .
#       Version: 1.0
#       Disclaimer: Use it an your own risk. I have tested and used it in my lab and production but every setup is different so it is recommended to test the script in your lab before using it.

$PS = Read-host "Please type hostname of the Passive Server"
$AS = Read-host "Please type FQDN of the Active Server"

#Drain active mail queues on the mailbox server
Set-ServerComponentState $PS -Component HubTransport -State Draining -Requester Maintenance

#To redirect messages pending delivery in the local queues to another Mailbox server run:
#Note: The Target Server value has to be the target server’s FQDN and that the target server shouldn’t be in maintenance mode
Redirect-Message -Server $PS -Target $AS -confirm:$False

#Transport Services Restart
Restart-Service MSExchangeFrontEndTransport
Restart-Service MSExchangeTransport

#To move all active databases currently hosted on the DAG member to other DAG members, run
Set-MailboxServer $PS -DatabaseCopyAutoActivationPolicy Unrestricted
Set-MailboxServer $PS -DatabaseCopyActivationDisabledAndMoveNow $True
Move-ActiveMailboxDatabase -server $PS -activateonserver $AS

#To prevent the server from hosting active database copies, run
Set-MailboxServer $PS -DatabaseCopyAutoActivationPolicy Blocked

#To put the server in maintenance mode run:
Set-ServerComponentState $PS -Component ServerWideOffline -State Inactive -Requester Maintenance

#To prevents the node from being and becoming the PAM, pause the cluster node by running
Suspend-ClusterNode $PS

#End of script 