$theUserName = "JEDI419"
Get-EventLog -LogName security -ComputerName DRDC01 -InstanceId 4740 -Message *$theUserName* | ft message -Wrap