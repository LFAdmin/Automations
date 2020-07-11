#       Author: Prabhat Nigam
#       Date: 06/03/2015
#       Description: Exchange 2013 URLs Configuration Script
$InternalURL = Read-Host "Type internal url starting with https://"
$ExternalURL = Read-host "Type external url starting with https://"
$Autodiscoverserviceinternaluri = Read-host "Type emaildomain.com starting with https://autodiscover."

Get-Clientaccessserver | Set-ClientAccessServer –AutodiscoverServiceInternalUri “$Autodiscoverserviceinternaluri/autodiscover/autodiscover.xml”

Get-webservicesvirtualdirectory | Set-webservicesvirtualdirectory –internalurl “$InternalURL/ews/exchange.asmx” –Externalurl “$ExternalURL/ews/exchange.asmx”
Get-oabvirtualdirectory | Set-oabvirtualdirectory –internalurl “$InternalURL/oab” –Externalurl “$ExternalURL/oab”
Get-owavirtualdirectory | Set-owavirtualdirectory –internalurl “$InternalURL/owa” –Externalurl “$ExternalURL/owa”
Get-ecpvirtualdirectory | Set-ecpvirtualdirectory –internalurl “$InternalURL/ecp” –Externalurl “$ExternalURL/ecp”
Get-ActiveSyncVirtualDirectory | Set-ActiveSyncVirtualDirectory -InternalUrl "$InternalURL/Microsoft-Server-ActiveSync" -ExternalUrl "$ExternalURL/Microsoft-Server-ActiveSync"

#get commands to  to doublecheck the config

get-webservicesvirtualdirectory | ft identity,Externalurl,Internalurl
get-oabvirtualdirectory | ft identity,Externalurl,Internalurl
get-owavirtualdirectory | ft identity,Externalurl,Internalurl
get-ecpvirtualdirectory | ft identity,Externalurl,Internalurl
get-ActiveSyncVirtualDirectory | ft identity,Externalurl,Internalurl
get-ClientAccessServer | ft identity,AutodiscoverServiceInternalUri

#End of script