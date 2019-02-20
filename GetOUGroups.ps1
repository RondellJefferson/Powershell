	
#Get-ADUser -filter * -SearchBase "CN=Users,DC=ptown,DC=local" | Select-Object -ExpandProperty name | Set-Content C:\Users\rjefferson\Documents\code\AllUserAccounts\Users.txt
#$usersInOU = Get-ADOrganizationalUnit -Filter * -SearchBase "OU=Departments,OU=Parts Town Users,DC=ptown,DC=local" | Select-Object -ExpandProperty distinguishedName | Set-Content "C:\Users\rjefferson\Documents\code\AllUserAccounts\PartstownUsersOU.txt"
$userInOU = Get-Content "C:\Users\rjefferson\Documents\code\AllUserAccounts\PartstownUsersOU.txt"
Remove-Item "C:\Users\rjefferson\Documents\code\AllUserAccounts\PartstownUsers.txt"
foreach ($OU in $userInOU){
    $findADUser = Get-ADUser -filter * -SearchBase "$OU" | Select-Object -ExpandProperty samaccountname | add-Content "C:\Users\rjefferson\Documents\code\AllUserAccounts\PartstownUsers.txt"
    }

Remove-Item "C:\Users\rjefferson\Documents\code\AllUserAccounts\PartstownUsers.txt"
$findADUser = Get-ADUser -filter * -SearchBase "OU=Departments,OU=Parts Town Users,DC=ptown,DC=local" | Select-Object -ExpandProperty samaccountname | add-Content "C:\Users\rjefferson\Documents\code\AllUserAccounts\PartstownUsers.txt"