$aUser = Get-ADUser -Filter * -SearchBase "OU=Warehouse Users,OU=Departments,OU=Parts Town Users,DC=ptown,DC=local" -Properties samAccountName, displayName
ForEach ($u in $aUser){
    $userDisplay = Get-ADUser $u -Properties displayname | Select-Object displayname
    $userAccount = Get-ADUser $u -Properties samAccountName | Select-Object samAccountName
    $userDisplayString = $userDisplay.displayname.toString()
    $userAccountString = $userAccount.samAccountName.toString()
    #$rename = $userString | ForEach-Object {$_ + ' (IT)'}
    $userAccountString | Set-ADUser -Replace @{jabberName="$($userDisplayString) (DC)"}
}