Set-Location "S:\IT Files\NewUsers"
$users = Import-Csv DisableUserAD.csv
foreach ($u in $users){ 
    $myFirst = $u.First
    $myLast = $u.Last
    $Filter = "givenName -like ""*$myFirst*"" -and sn -like ""$myLast"""
    $theUserName = Get-ADUser -Filter $Filter | Select-Object samAccountName | select -expandproperty samAccountName
    $enabledOrNaw = Get-ADUser -Filter $Filter -Properties enabled | Select-Object enabled | select -ExpandProperty enabled 
    if ($enabledOrNaw -eq $true){
        Disable-ADAccount -identity $theUsername 
        Write-Host "$myFirst $myLast has been disabled Today"
    } else {
        Write-Host "$myFirst $myLast has already been disabled"
    }

}