Get-ADUser -Identity user1maitre -Properties memberof | Select-Object -ExpandProperty memberof | Add-ADGroupMember -Members user2client

Get-ADGroupMember "GroupeA" | Get-ADUser | ForEach-Object {Add-ADGroupMember -Identity "GroupeB" -Members $_}
