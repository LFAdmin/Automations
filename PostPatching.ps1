#       Author: Prabhat Nigam. 
#       Twitter: @PrabhatNigamXHG
#       Blog: MSExchangeGuru.com
#       Email: Prabhat.Nigam@GoldenFive.net
#       Website: www.GoldenFiveConsulting.com
#       Date: 09/03/2016.
#       Description: Post Patching or Restart Exchange Role transfer from server A to Server B for Exchange 2013 and 2016 .
#       Version: 1.0
#       Disclaimer: Use it an your own risk. I have tested and used it in my lab and production but every setup is different so it is recommended to test the script in your lab before using it.
$PS = Read-host "Please type the hostname of the rebooted server"
$DG = Read-host "Please type DAG NAME"

#Role Back all roles
Set-ServerComponentState $PS -Component ServerWideOffline -State Active -Requester Maintenance
Resume-ClusterNode $PS
Set-MailboxServer $PS -DatabaseCopyActivationDisabledAndMoveNow $False
Set-MailboxServer $PS -DatabaseCopyAutoActivationPolicy Unrestricted
Set-ServerComponentState $PS -Component HubTransport -State Active -Requester Maintenance
Restart-Service MSExchangeTransport
Restart-Service MSExchangeFrontEndTransport

#Once completed run the below mentioned commands to balance the database and move the PAM
cd $exscripts
.\RedistributeActiveDatabases.ps1 -DagName $DG -BalanceDbsByActivationPreference -ShowFinalDatabaseDistribution -confirm:$false

#Move the cluster group to this server
Get-ClusterGroup "Cluster Group" | Move-ClusterGroup -Node $PS

#Change the directory back to temp
CD c:\temp

#End of script