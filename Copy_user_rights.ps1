Get-ADUser -Identity user1maitre -Properties memberof | Select-Object -ExpandProperty memberof | Add-ADGroupMember -Members user2client
