$disabledAccounts = Search-ADAccount -AccountDisabled -UsersOnly | Get-ADUser -Properties Description | Select-Object Name, DistinguishedName, Description

#Exclure des comptes venant de certaines OU
$disabledAccounts | ?{$_.DistinguishedName -notmatch "OU=XXXX"} | Out-GridView
