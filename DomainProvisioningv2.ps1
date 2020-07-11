#       Author: Prabhat Nigam
#       Date: 10/24/2016
# 	    Blog: MSExchangeGuru.com
#       Email: Prabhat.Nigam@GoldenFive.net
#       Website: www.GoldenFiveConsulting.com
#       Description: Exchange 2013/2016 New Domain Provisioning
#       Version: 1.0
#       Disclaimer: 	Use it an your own risk. I have tested and used it in my lab and production 
#			            but every setup is different so it is recommended to test the script in your lab before using it.
#	    Important: 	    Restart of Information Store is required to mount the databases and to create 2nd copy.
#			            So Plan to run the script during change window.
#       Updated: 10/25/2016: MVP Vikas Sukhija: RecipientFilter updated from "CustomAttribute2 -eq '$Domain'" to "((Alias -ne $null) -and (CustomAttribute2 -eq '$Domain'))"


#Define Variables
$Name = Import-csv -Path C:\temp\NewDB.csv
$GAL = $Name.name +"_GAL"
$OAB = $Name.name +"_OAB2016"
$Domain = $Name.Domain
$AL = $Name.name +"_AL"
$RAL = $Name.name +"_RoomAL"
$ABP = $Name.name +"_ABP"
$str = new-object System.Text.Stringbuilder
$noll = "$" + "null"
$str= "{" + "((Alias -ne $noll) -and (CustomAttribute2 -eq '$Domain'))" + "}"

# Create New Accepted Domain
New-AcceptedDomain -Name $Domain -DomainName $Domain -DomainType Authoritative

#Create New Global AddressList. Ypu can update recipientFilter as per the requirement or change the custom attribute.
New-GlobalAddressList -Name $GAL -RecipientFilter $str | out-null

#Create New Room AddressList
New-AddressList -Name $RAL -RecipientFilter "$str -and RecipientDisplayType -eq 'ConferenceRoomMailbox' -or RecipientDisplayType -eq 'SyncedConferenceRoomMailbox'" | out-null

#Create New AddressList
New-AddressList -Name $AL -RecipientFilter $str | out-null

#Create New Offlice Address Book
New-OfflineAddressBook -Name $OAB -AddressLists $GAL -VirtualDirectories $null -GlobalWebDistributionEnabled $true | out-null

#Create New Address Book Policy
New-AddressBookPolicy -Name $ABP -AddressLists $AL -OfflineAddressBook $OAB -GlobalAddressList $GAL -RoomList $RAL | out-null

#Create New Mailbox Database
Import-CSV C:\temp\NewDB.csv | ForEach {New-mailboxdatabase -Name $Name.name -LogFolderPath $_.LogFolderPath -EdbFilePath $_.EdbFilePath -OfflineAddressBook $OAB -Server $_.Server1}

#Switch Server
Move-ActiveMailboxDatabase -Server $_.Server1 -activateonServer $_.Server2

#Restarting Information Store Service
Stop-Service MSExchangeIS
Start-Service MSExchangeIS

#Mounting Database
Mount-Database $Name.name

#Updating Databsase Quota Limits
Import-CSV C:\temp\NewDB.csv | ForEach {Set-MailboxDatabase -Identity $Name.name -ProhibitSendReceiveQuota $_.ProhibitSendReceiveQuota -ProhibitSendQuota $_.ProhibitSendQuota -IssueWarningQuota $_.IssueWarningQuota}

#End of script - 2nd Database Copy creation script should be run post primary database and indexing are showing healthy.
