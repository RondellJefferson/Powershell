Get-ADUser -Identity eashcraft -Properties memberof |
#Grab the object of MemberOf
Select-Object -ExpandProperty memberof | Get-ADGroup | Remove-ADGroupMember -Members eashcraft -Confirm:$false