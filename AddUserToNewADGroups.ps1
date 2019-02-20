#Use username
$userToChange = ashalton
$copyUser = kstern

Get-ADUser -Identity $userToChange -Properties memberof |
#Grab the object of MemberOf
Select-Object -ExpandProperty memberof | Get-ADGroup | Remove-ADGroupMember -Members $userToChange -Confirm:$false
#Copy an AD User GroupMembership
#Grab the User and select MemberOf Tab
    Get-ADUser -Identity $copyUser -Properties memberof |
#Grab the object of MemberOf
    Select-Object -ExpandProperty memberof |
#Use the object and place it with the user below
    Add-ADGroupMember -Members $userToChange