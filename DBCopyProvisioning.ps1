#       Date: 10/24/2016
# 	    Blog: MSExchangeGuru.com
#       Email: Prabhat.Nigam@GoldenFive.net
#       Website: www.GoldenFiveConsulting.com
#       Description: Database Copy creation Script
#       Version: 1.0
#       Disclaimer: 	Use it an your own risk. I have tested and used it in my lab and production 
#			            but every setup is different so it is recommended to test the script in your lab before using it.	
#	    Important: 	    Restart of Information Store is required to mount the databases and to create 2nd copy.
#			            So Plan to run the script during change window.

#2nd Database Copy creation script should be run post primary database and indexing are showing healthy with zero log & replay queue. 
#Check by running the command:	Get-MailboxDatabaseCopyStatus *\*

#Define Variables
$Name = Import-csv -Path C:\temp\NewDB.csv
$DG = $Name.DAG

#Creating Second Copy of the Database
Import-CSV C:\temp\NewDB.csv\ | ForEach {Add-MailboxDatabaseCopy -Identity $_.name -MailboxServer $_.Server2 -ActivationPreference $_.AP}

#Switch Server
Move-ActiveMailboxDatabase -Server $_.Server2 -activateonServer $_.Server1

#Restarting Information Store Service
Stop-Service MSExchangeIS
Start-Service MSExchangeIS

#Redistribute the Database as per Activation Preference
cd $exscripts
.\RedistributeActiveDatabases.ps1 -DagName $DG -BalanceDbsByActivationPreference -ShowFinalDatabaseDistribution -confirm:$false

#End of script