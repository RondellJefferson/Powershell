$myUsername = "rjefferson@partstown.com"
$password = Get-Content "C:\Users\rjefferson\Documents\code\pass.txt" | ConvertTo-SecureString 
$Creds = new-object -typename System.Management.Automation.PSCredential("rjefferson@partstown.com",$password)

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Creds -Authentication Basic -AllowRedirection
Import-PSSession $Session

#Get-PSSession | Remove-PSSession
