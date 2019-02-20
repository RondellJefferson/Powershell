$users = Import-Csv BastianUsers.csv

Function Get-LogonTime{
    param ($aUser)
    #Receive the user's LastDomainLogin Time.
    $whenLogonTime = Get-ADUser $u.UserID -Properties lastlogonTimeStamp | select-object -ExpandProperty lastLogonTimeStamp
    $LoginTime = [datetime]::FromFileTime($whenLogonTime).ToString('g')
    return $LoginTime
}

Function Get-ExpireDate{
    param ($aUser)
    #Receive the password expire date and convert it to normal time.
    $expirePass = Get-ADUser $u.UserID -Properties msDS-UserPasswordExpiryTimeComputed | Select-Object -ExpandProperty msDS-UserPasswordExpiryTimeComputed
    $giveMeExpireDate = [datetime]::FromFileTime($expirePass).ToString('g')
    return $giveMeExpireDate
}
ForEach ($u in $users){
#Function to receive the Users Full Name
$theUsersName = Get-ADUser $u.UserID -Properties name | Select-Object -ExpandProperty name
$theLoginTime = Get-LogonTime $u.UserID
$theExpirationDate = Get-ExpireDate $u.UserID

Write-Host "The Last Time $theUsersName Logged on was $theLoginTime with username $($u.UserID) and Password Expiration date is on $theExpirationDate"
}