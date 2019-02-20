$myFirst = "Rondell".ToUpper()
$myLast = "Jefferson".ToUpper()
#$Filter = "givenName -like ""*$myFirst*"" -and sn -like ""$myLast"""
#$copiedUser = Get-ADUser -Filter $Filter | Select-Object samAccountName | select -expandproperty samAccountName

$fullname = "$myFirst $myLast"
$SplitName = $fullname.Split()
$splitFirstName = $firstName.ToCharArray()
$splitLastName = $lastName.ToCharArray()
$newInitial = $splitFirstName[0] + $splitLastName[0] + $splitLastName[-1].ToString().toUpper()
$newInitial