#$File = "C:\temp\newPassword1.txt"
#$Password = Read-Host -AsSecureString| ConvertTo-SecureString -AsPlainText -Force
#$Password | ConvertFrom-SecureString | Out-File $File

	
(get-credential).password | ConvertFrom-SecureString | set-content "C:\Users\rjefferson\Documents\code\passrd22.txt"