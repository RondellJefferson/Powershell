$readName = Read-Host "What Is The User's Display Name"
$newLocation = Read-Host "What Is The New Office Location Number"
$getUserPlease = Get-ADUser -LDAPFilter "(name=$readName)"
$theUsersName = $getUserPlease | select samAccountName | Select-Object -ExpandProperty samAccountName
Get-ADUser -Identity $theUsersName | Set-ADUser -Office $newLocation