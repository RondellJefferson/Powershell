Copy-Item C:\Users\rjefferson\Documents\code -Destination "\\ptown.local\files\User Data\HomeDrives\rjefferson\PC\Doc\NewPC\Codes" -Recurse -Force
Copy-Item C:\Users\rjefferson\Documents\programming -Destination "\\ptown.local\files\User Data\HomeDrives\rjefferson\PC\Docs\NewPC\Programming" -Recurse -Force

#Get-ChildItem -Path C:\Users\rjefferson\Newfolder\*aTransferFile*.txt | Move-Item -Destination C:\Users\rjefferson\Newfolder\newFolders
#Set-Location C:\Users\rjefferson\Newfolder\newFolders
#Get-ChildItem -Filter "*aTransferFile*" | Rename-Item -NewName {$_.name -replace '\.txt$','.csv' }
#rename  *.txt   *.csv