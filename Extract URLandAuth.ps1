#       Author: Prabhat Nigam | Microsoft MVP | CTO | Chief Architect
#       Company: Golden Five Consulting
#       Date: 04/26/2016
#       Description: Exchange 2013 URLs and Authentication Export Script

$FormatEnumerationLimit =-1
$ErrorActionPreference="SilentlyContinue"
$ErrorActionPreference = "Continue"

$EWS = get-webservicesvirtualdirectory | FT identity,Server,Externalurl,Internalurl,BasicAuthentication,WindowsAuthentication,OAuthAuthentication,@{label="InternalAuthenticationMethods";expression={[string]::join(“;”,$_.InternalAuthenticationMethods)}},@{label="ExternalAuthenticationMethods";expression={[string]::join(“;”,$_.ExternalAuthenticationMethods)}} -AutoSize | out-string -Width 660
$OAB = get-oabvirtualdirectory | FT identity,Server,Externalurl,Internalurl,BasicAuthentication,WindowsAuthentication,OAuthAuthentication,@{label="InternalAuthenticationMethods";expression={[string]::join(“;”,$_.InternalAuthenticationMethods)}},@{label="ExternalAuthenticationMethods";expression={[string]::join(“;”,$_.ExternalAuthenticationMethods)}} -AutoSize | out-string -Width 660
$OWA = get-owavirtualdirectory | FT identity,Server,Externalurl,Internalurl,BasicAuthentication,WindowsAuthentication,OAuthAuthentication,@{label="InternalAuthenticationMethods";expression={[string]::join(“;”,$_.InternalAuthenticationMethods)}},@{label="ExternalAuthenticationMethods";expression={[string]::join(“;”,$_.ExternalAuthenticationMethods)}} -AutoSize | out-string -Width 660
$ECP = get-ecpvirtualdirectory | FT identity,Server,Externalurl,Internalurl,BasicAuthentication,WindowsAuthentication,OAuthAuthentication,@{label="InternalAuthenticationMethods";expression={[string]::join(“;”,$_.InternalAuthenticationMethods)}},@{label="ExternalAuthenticationMethods";expression={[string]::join(“;”,$_.ExternalAuthenticationMethods)}} -AutoSize | out-string -Width 660
$EAS = get-ActiveSyncVirtualDirectory | FT identity,Server,Externalurl,Internalurl,BasicAuthEnabled,WindowsAuthEnabled,ClientCertAuth,@{label="InternalAuthenticationMethods";expression={[string]::join(“;”,$_.InternalAuthenticationMethods)}},@{label="ExternalAuthenticationMethods";expression={[string]::join(“;”,$_.ExternalAuthenticationMethods)}} -AutoSize | out-string -Width 660
$MAPI = get-mapivirtualdirectory | FT identity,Server,Externalurl,Internalurl,@{label="IISAuthenticationMethods";expression={[string]::join(“;”,$_.IISAuthenticationMethods)}},@{label="InternalAuthenticationMethods";expression={[string]::join(“;”,$_.InternalAuthenticationMethods)}},@{label="ExternalAuthenticationMethods";expression={[string]::join(“;”,$_.ExternalAuthenticationMethods)}} -AutoSize | out-string -Width 660
$OA = get-outlookanywhere | FT identity,Server,Externalhostname,Internalhostname,@{label="IISAuthenticationMethods";expression={[string]::join(“;”,$_.IISAuthenticationMethods)}},InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod -AutoSize | out-string -Width 660
$AuD = get-ClientAccessService | FT identity,AutodiscoverServiceInternalUri -AutoSize

$OBJ = new-object PSObject
$OBJ | add-member NoteProperty Exchange_Web_Services $EWS
$OBJ | add-member NoteProperty Offline_Address_Book $OAB
$OBJ | add-member NoteProperty Outlook_Web_Access $OWA
$OBJ | add-member NoteProperty Exchange_Control_Panel $ECP
$OBJ | add-member NoteProperty Active_Sync $EAS
$OBJ | add-member NoteProperty MAPI $MAPI
$OBJ | add-member NoteProperty Outlook_Anywhere $OA
$OBJ | add-member NoteProperty Auto_Discover ($AuD | Out-String).Trim()

$OBJ | export-csv "output.csv" -notypeinformation

$FormatEnumerationLimit = 16
