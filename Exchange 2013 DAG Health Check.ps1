#############################################################################
#       Author: Mahesh Sharma
#       Modified by 1: Vikas SUkhija 
#       Modified by 2: Prabhat Nigam on 4/22/2015 and 6/3/2015    
#       Date: 06/10/2013
#	    Modified:06/19/2013 - made it to run from any path
#       Modified:02/09/2014 - started modifying it for exchange 2010
#       Modified:02/18/2014 - modified to include all mailox servers in test mailflow
#       Modified:05/22/2014 - added activation prefrence
#	    Modified:09/09/2014 - included DAG DB backups status
#	    Modified:04/15/2015 - updated for Exchange 2013
#       Modified:04/22/2015 - Added report as attachment, suppress get-transportserver warning and powershell url
#       Modified:06/03/2015 - Added some more change for the convinence of the users.
#       Description: ExChange Health Status
#############################################################################

########################### Add Exchange Shell##############################

If ((Get-PSSnapin | where {$_.Name -match "Microsoft.Exchange.Management.PowerShell"}) -eq $null)
{
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
}

###########################Define Variables##################################
#Make sure to keep the script in the same directory as reportpath.
$reportpath = "c:\temp\2013Report.htm" 

if((test-path $reportpath) -like $false)
{
new-item $reportpath -type file
}
#You can replace the text "Read-host "Please Provide Your SMTP Relay Servers FQDN"" with the FQDN of the SMTP server
$smtphost = Read-host "Please Provide Your SMTP Relay Servers FQDN like servername.domain.com"
#You can replace the text "Read-host "Please Provide Sender's email address"" with the Sender email address
$from = Read-host "Please Provide Sender's email addresslike sender@domain.com"
#You can replace the text "Read-host "Please Provide Recipient's email address"" with the Recepient email address
$to = Read-host "Please Provide Recipient's email address like recipient@domain.com"
$hrs = (get-date).Addhours(-24) 
#Update the URL for the remote session with the powershell internalurl which you can get by running Get-PowerShellVirtualDirectory | fl internalurl
$url = "http://dchyb.g5.com/Powershell"

###############################HTml Report Content############################
$report = $reportpath

Clear-Content $report 
Add-Content $report "<html>" 
Add-Content $report "<head>" 
Add-Content $report "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>" 
Add-Content $report '<title>Exchange Status Report</title>' 
add-content $report '<STYLE TYPE="text/css">' 
add-content $report  "<!--" 
add-content $report  "td {" 
add-content $report  "font-family: Tahoma;" 
add-content $report  "font-size: 11px;" 
add-content $report  "border-top: 1px solid #999999;" 
add-content $report  "border-right: 1px solid #999999;" 
add-content $report  "border-bottom: 1px solid #999999;" 
add-content $report  "border-left: 1px solid #999999;" 
add-content $report  "padding-top: 0px;" 
add-content $report  "padding-right: 0px;" 
add-content $report  "padding-bottom: 0px;" 
add-content $report  "padding-left: 0px;" 
add-content $report  "}" 
add-content $report  "body {" 
add-content $report  "margin-left: 5px;" 
add-content $report  "margin-top: 5px;" 
add-content $report  "margin-right: 0px;" 
add-content $report  "margin-bottom: 10px;" 
add-content $report  "" 
add-content $report  "table {" 
add-content $report  "border: thin solid #000000;" 
add-content $report  "}" 
add-content $report  "-->" 
add-content $report  "</style>" 
Add-Content $report "</head>" 
Add-Content $report "<body>" 
add-content $report  "<table width='100%'>" 
add-content $report  "<tr bgcolor='Lavender'>" 
add-content $report  "<td colspan='7' height='25' align='center'>" 
add-content $report  "<font face='tahoma' color='#003399' size='4'><strong>DAG Active Manager</strong></font>" 
add-content $report  "</td>" 
add-content $report  "</tr>" 
add-content $report  "</table>" 
 
add-content $report  "<table width='100%'>" 
Add-Content $report  "<tr bgcolor='IndianRed'>" 
Add-Content $report  "<td width='10%' align='center'><B>Identity</B></td>" 
Add-Content $report  "<td width='5%' align='center'><B>PrimaryActiveManager</B></td>" 
Add-Content $report  "<td width='20%' align='center'><B>OperationalMachines</B></td>" 
 

Add-Content $report "</tr>" 

##############################Get ALL DAG's##################################

$inputdag = @()

$indag = Get-DatabaseAvailabilityGroup 

foreach($dg in $indag)
{
  $mem= $dg.Servers
  foreach($m in $mem){
	if((Get-ExchangeServer $m.Name).AdminDisplayVersion -like "*15.0*"){
        $inputdag += $dg.Name}
	}
}

$inputdag = $inputdag | select -uniq


################################################################################################################
################################################################################################################


$dagList = $inputdag
$TestMailFlow =  Get-ExchangeServer | where{$_.ServerRole -like "*Mailbox*"}

$report = $reportpath



##########################################################################################################
##############################################Check PAM###################################################


foreach ($dag in $dagList) 
{ 
 
		
		$FullStatus = Get-DatabaseAvailabilityGroup -Status $dag

		Foreach ($status in $Fullstatus)
		{

			
		$Identity = $status.identity
	        $PrimaryActiveManager =  $status.PrimaryActiveManager
                $Servers = $status.Servers
		Add-Content $report "<tr>" 
		Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B> $Identity</B></td>" 
         	Add-Content $report "<td bgcolor= 'Aquamarine' align=center>  <B>$PrimaryActiveManager</B></td>" 
		Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$Servers</B></td>" 
		Add-Content $report "</tr>" 			


		}
        	
		
}

##################################################################################################################
############################################## Mailbox Database Status ###########################################


add-content $report  "<tr bgcolor='Lavender'>" 
add-content $report  "<td colspan='7' height='25' align='center'>" 
add-content $report  "<font face='tahoma' color='#003399' size='4'><strong>Mailbox Database Status</strong></font>" 
add-content $report  "</td>" 
add-content $report  "</tr>"

add-content $report  "</tr>" 
add-content $report  "</table>" 
add-content $report  "<table width='100%'>" 
Add-Content $report "<tr bgcolor='IndianRed'>"
Add-Content $report  "<td width='25%' align='center'><B>databaseName</B></td>" 
Add-Content $report "<td width='25%' align='center'><B>Status</B></td>" 
Add-Content $report "<td width='25%' align='center'><B>ActiveCopy</B></td>" 
Add-Content $report  "<td width='25%' align='center'><B>CopyQueuelength</B></td>" 
Add-Content $report  "<td width='25%' align='center'><B>ReplayQueueLength</B></td>" 
Add-Content $report  "<td width='25%' align='center'><B>LastInspectedLogTime</B></td>" 
Add-Content $report  "<td width='25%' align='center'><B>ContentIndexState</B></td>" 
Add-Content $report "</tr>" 


$mbxdb = Get-MailboxDatabase | Get-MailboxDatabaseCopyStatus 

$mbxdb = $mbxdb | Sort-Object Status -Descending


foreach ($db in $mbxdb)
{

$dbname = $db.name
foreach($dbn in $dbname)

    { $stcopy=Get-MailboxDatabaseCopyStatus $dbn
      $srv = $stcopy.ActiveDatabaseCopy
      $mbxdbname = $stcopy.DatabaseName
      $mbxdbname1 = get-mailboxdatabase $mbxdbname
      $acpref = ($mbxdbname1 | select -ExpandProperty ActivationPreference  | where {$_.key -like "$srv"}).value
      }



$status = $db.Status
$ActiveCopy = $db.ActiveCopy
$CopyQueuelength = $db.CopyQueuelength
$ReplayQueueLength = $db.ReplayQueueLength
$LastInspectedLogTime = $db.LastInspectedLogTime
$ContentIndexState = $db.ContentIndexState

$result = $flow.TestMailflowResult
$time = $Flow.MessageLatencyTime
$remote =  $Flow.IsRemoteTest
Add-Content $report "<tr>" 
Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$dbname</B></td>" 

       if ((($status -eq "Mounted") -and ($acpref -eq 1)) -or ($status -eq "Healthy"))
            {
            Add-Content $report "<td bgcolor= 'Aquamarine' align=center>  <B>$status</B></td>" 
            }
       elseif((($status -eq "Mounted") -and ($acpref -ne 1)) -or ($status -eq "Healthy"))
            {
           Add-Content $report "<td bgcolor= 'yellow' align=center>  <B>$status</B></td>"
            }
        
       else
            {
            Add-Content $report "<td bgcolor= 'Red' align=center>  <B>$status</B></td>" 
            }
Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$ActiveCopy</B></td>" 
Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$CopyQueuelength</B></td>" 
Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$ReplayQueueLength</B></td>" 
Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$LastInspectedLogTime</B></td>" 

           if ($ContentIndexState -eq"Healthy")

            {
		Add-Content $report "<td bgcolor= 'Aquamarine' align=center>  <B>$ContentIndexState</B></td>" 
            }
           else
            {
		Add-Content $report "<td bgcolor= 'Red' align=center>  <B>$ContentIndexState</B></td>" 
            }

Add-Content $report "</tr>"

}

#################################################################################################################
##############################################DAG DB Backup Status###############################################

add-content $report  "<tr bgcolor='Lavender'>" 
add-content $report  "<td colspan='7' height='25' align='center'>" 
add-content $report  "<font face='tahoma' color='#003399' size='4'><strong>DAG Database Backup Status</strong></font>" 
add-content $report  "</td>" 
add-content $report  "</tr>"

add-content $report  "</tr>" 
add-content $report  "</table>" 
add-content $report  "<table width='100%'>" 
Add-Content $report "<tr bgcolor='IndianRed'>"
Add-Content $report  "<td width='10%' align='center'><B>Database</B></td>" 
Add-Content $report  "<td width='5%' align='center'><B>BackupInProgress</B></td>" 
Add-Content $report  "<td width='10%' align='center'><B>SnapshotLastFullBackup</B></td>" 
Add-Content $report  "<td width='5%' align='center'><B>SnapshotLastCopyBackup</B></td>" 
Add-Content $report  "<td width='10%' align='center'><B>LastFullBackup</B></td>" 
Add-Content $report  "<td width='5%' align='center'><B>RetainDeletedItemsUntilBackup</B></td>"

Add-Content $report "</tr>" 

$dbst= Get-MailboxDatabase | where{$_.MasterType -like "DatabaseAvailabilityGroup"}

$dbst | foreach{$st=Get-MailboxDatabase $_ -status
$dbname =  $st.Name
$dbbkprg = $st.BackupInProgress
$dbsnpl = $st.SnapshotLastFullBackup
$dbsnplc= $st.SnapshotLastCopyBackup
$dblfb = $st.LastFullBackup
$dbrd = $st.RetainDeletedItemsUntilBackup
		Add-Content $report "<tr>" 
		Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$dbname</B></td>" 
         	Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$dbbkprg</B></td>" 
		Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$dbsnpl</B></td>" 
         	Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$dbsnplc</B></td>" 
if($dblfb -lt $hrs)
{
		Add-Content $report "<td bgcolor= 'Red' align=center>  <B>$dblfb</B></td>"
}
else
{
   		Add-Content $report "<td bgcolor= 'Aquamarine' align=center>  <B>$dblfb</B></td>"  
}    	

		Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$dbrd</B></td>" 
		Add-Content $report "</tr>" 


}


##################################################################################################################
############################################## Test mail Flow For DAG ############################################


add-content $report  "<tr bgcolor='Lavender'>" 
add-content $report  "<td colspan='7' height='25' align='center'>" 
add-content $report  "<font face='tahoma' color='#003399' size='4'><strong>Mail Flow Test Report</strong></font>" 
add-content $report  "</td>" 
add-content $report  "</tr>"

add-content $report  "</tr>" 
add-content $report  "</table>" 
add-content $report  "<table width='100%'>" 
Add-Content $report "<tr bgcolor='IndianRed'>"
Add-Content $report  "<td width='25%' align='center'><B>Server</B></td>" 
Add-Content $report  "<td width='25%' align='center'><B>Result</B></td>" 
Add-Content $report "<td width='25%' align='center'><B>Message Latency Time</B></td>" 
Add-Content $report  "<td width='25%' align='center'><B>IsRemoteTest</B></td>" 
Add-Content $report "</tr>" 


Foreach ($server in $TestMailFlow)
{

$server = $server.Name
write-host $server

$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $url

$flow = Invoke-Command -Session $session {Test-Mailflow}

#test-mailflow $server

if($flow -ne $null)
{

$result = $flow.TestMailflowResult
$time = $Flow.MessageLatencyTime
$remote =  $Flow.IsRemoteTest
Add-Content $report "<tr>" 
Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$server</B></td>" 
if ($result -eq "Success")
{
Add-Content $report "<td bgcolor= 'Aquamarine' align=center>  <B> $result</B></td>"
}
else
{
Add-Content $report "<td bgcolor= 'Red' align=center>  <B> $result</B></td>"
} 
Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$time</B></td>" 
Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$remote</B></td>" 

Add-Content $report "</tr>"
}


}



#####################################################################################################################
############################################## Get Queue For HUB Servers ############################################


add-content $report  "<tr bgcolor='Lavender'>" 
add-content $report  "<td colspan='7' height='25' align='center'>" 
add-content $report  "<font face='tahoma' color='#003399' size='4'><strong>Mail Queue Status</strong></font>" 
add-content $report  "</td>" 
add-content $report  "</tr>"

add-content $report  "</tr>" 
add-content $report  "</table>" 
add-content $report  "<table width='100%'>" 
Add-Content $report "<tr bgcolor='IndianRed'>"
Add-Content $report  "<td width='25%' align='center'><B>Identity</B></td>" 
Add-Content $report "<td width='25%' align='center'><B>Delivery Type</B></td>" 
Add-Content $report  "<td width='25%' align='center'><B>Status</B></td>" 
Add-Content $report "<td width='25%' align='center'><B>Message Count</B></td>" 
Add-Content $report  "<td width='25%' align='center'><B>Next Hop Domain</B></td>"
Add-Content $report "</tr>" 

$GetHub = Get-TransportServer -WarningAction:silent | get-Queue

Foreach ($Queue in $GetHub)
{

$Identity = $Queue.Identity
$DeliveryType = $Queue.DeliveryType
$Status = $Queue.Status
$MSgCount =  $Queue.Messagecount
$NextHopDomain = $Queue.NextHopDomain


Add-Content $report "<tr>" 
Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B> $Identity</B></td>" 
Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$DeliveryType</B></td>" 
Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$Status</B></td>" 
Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$MSgCount</B></td>" 
Add-Content $report "<td bgcolor= 'GainsBoro' align=center>  <B>$NextHopDomain</B></td>" 

Add-Content $report "</tr>"


}

###########################################################################################################################
######################################################### Send Mail #######################################################


Add-content $report  "</table>" 
Add-Content $report "</body>" 
Add-Content $report "</html>"

 
$subject = "Exchange Status Check Report" 
$body = Get-Content $reportpath
$smtp= New-Object System.Net.Mail.SmtpClient $smtphost 
$msg = New-Object System.Net.Mail.MailMessage $from, $to, $subject, $body 
$msg.isBodyhtml = $true 
$msg.Attachments.Add($Reportpath)
$smtp.send($msg) 

###################################################Exchange Test Complete##################################################


