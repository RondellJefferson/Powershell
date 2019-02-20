#Gets my credentials and uses to login Exchange Server Username = Full Admin Name <USERNAME>@ptown.local
#Setup the connection to On-Premise Exchange Server.
$myUsername = "rd22@ptown.local"
$myUsername = "rjefferson@partstown.com"
$password1 = Get-Content "C:\Users\rjefferson\Documents\code\passrd22.txt" | ConvertTo-SecureString 
$MyCreds = new-object -typename System.Management.Automation.PSCredential("rd22@ptown.local",$password1)
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://PRDEXCHANGE.ptown.local/PowerShell/ -Authentication Kerberos -Credential $MyCreds

#Makes the connection to the On-Premise Exchange Server
Import-PSSession $Session
$users = Import-Csv "C:\Users\rjefferson\Documents\code\ExternalContacts\ExternalContacts.csv" 
foreach ($u in $users){
$emailAddress = $u.Email
Add-DistributionGroupMember -identity "PX Daily Inventory" -Member $emailAddress
}