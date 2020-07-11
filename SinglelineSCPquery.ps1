#Written By#
#Prabhat Nigam - Microsoft Solution Archtitect#
#Microsoft Profile - http://social.technet.microsoft.com/profile/prabhatnigam/#
Import-module ActiveDirectory
Get-ADObject -LDAPFilter '(&(objectClass=serviceConnectionPoint)(|(keywords=67661d7F-8FC4-4fa7-BFAC-E1D7794C1F68)(keywords=77378F46-2C66-4aa9-A6A6-3E7A48B19596)))' -SearchScope Subtree -SearchBase 'CN=Configuration,DC=humed,DC=com' | Get-ADObject -Properties WhenCreated,ServiceBindingInformation,Keywords | ft Name,WhenCreated,ServiceBindingInformation,Keywords -Autosize
