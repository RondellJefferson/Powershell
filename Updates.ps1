get-wmiobject -computername 10.10.80.54 -class win32_quickfixengineering | Where-Object {$_.InstalledOn -gt "10/1/2017"}

