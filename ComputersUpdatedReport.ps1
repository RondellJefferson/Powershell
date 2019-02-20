$env:desktop
Get-HotFix -ComputerName desktop80 | Select-Object HotFixID,Description,InstalledOn | Where-Object {$_.InstalledOn -gt "10/1/2017"}
#$aString = $pcUpdated.installedon.Equals("6/16/2018")
#$isPCUpdated = $aString.ToString()

Test-Connection -ComputerName desktop83

get-wmiobject -ComputerName LAP-PERRIH -class win32_quickfixengineering | Select-Object HotFixID,Description,InstalledOn | Where-Object {$_.InstalledOn -gt "6/1/2017"} | sort -Descending