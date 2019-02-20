<#$users = get-content "C:\Users\rjefferson\Documents\Project\list1.txt"
ForEach ($u in $users){
$fullname = $u
$SplitName = $fullname.Split()
$firstName = $SplitName[0]
$lastName = $SplitName[1]
$Filter = "givenName -like ""*$firstName*"" -and sn -like ""$lastName"""
$copyUser = Get-ADUser -Filter $Filter | Select-Object samAccountName -ExpandProperty samAccountName 
try{
Get-ADUser $copyUser | Set-ADUser -City "Addison"
} catch { $u }

}
#>
<#
$password1 = Get-Content "C:\Users\rjefferson\Documents\code\passrd22.txt" | ConvertTo-SecureString 
$MyCreds = new-object -typename System.Management.Automation.PSCredential("rd22@ptown.local",$password1)
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://PRDEXCHANGE.ptown.local/PowerShell/ -Authentication Kerberos -Credential $MyCreds
Import-PSSession $Session
#>


$newUsers = Get-ADUser -Filter 'City -like "Addison"' -Properties samaccountname | Select-Object samaccountname -ExpandProperty samaccountname
ForEach ($new in $newUsers){
Add-DistributionGroupMember -Identity addisonil@partstown.com -Member $new
}