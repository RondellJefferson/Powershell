<# 
$env:COMPUTERNAME = Read-Host 'What Is The Computer Name'
#$pass = Read-Host 'What is your password?' -AsSecureString
psexec -s \\$env:COMPUTERNAME -u WBUser cmd
#Start-Sleep -s 35
shutdown /r /t 000
#exit
#>
$computers = (Get-Content C:\Users\rjefferson\Documents\code\list.txt)
$computers | Out-GridView -OutputMode Multiple | Restart-Computer -Force