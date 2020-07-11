#############################################################################
#       Author: Prabhat Nigam    
#       Twitter: @PrabhatNigamXHG
#       Blog: MSExchangeGuru.com
#       Email: Prabhat.Nigam@GoldenFive.net
#       Website: www.GoldenFiveConsulting.com
#       Version: 1.0
#       Date: 08/10/2017
#	    Description: Selected DL Member Export to one CSV file
#       Disclaimer: Use it an your own risk. I have tested and used it in my lab and production but every setup is different so it is recommended to test the script in your lab before using it.
#############################################################################

$Groups = Get-Content c:\temp\dl.csv | Get-DistributionGroup
$Groups | ForEach-Object {
$group = $_.Name
$members = ''
Get-DistributionGroupMember $group | ForEach-Object {
        If($members) {
              $members=$members + ";" + $_.Name
           } Else {
              $members=$_.Name
           }
  }
New-Object -TypeName PSObject -Property @{
      GroupName = $group
      Members = $members
     }
} | Export-CSV "c:\temp\DLExport.csv" -NoTypeInformation -Encoding UTF8