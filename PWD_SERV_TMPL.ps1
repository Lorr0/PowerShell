# Importer le module Active Directory
Import-Module ActiveDirectory

# Définir la date limite pour le contrôle des mots de passe (15 jours à partir d'aujourd'hui)
$limite = (Get-Date).AddDays(15)

# Récupérer tous les comptes d'utilisateurs de l'OU spécifiée
$utilisateurs = Get-ADUser -Filter * -SearchBase "#chemin d'OU des comptes a vérifier" -Properties "SamAccountName", "PasswordLastSet", "msDS-UserPasswordExpiryTimeComputed", "PasswordNeverExpires"

# Initialiser une variable pour vérifier s'il y a des comptes à expiration imminente
$comptesExpirant = $false

# Créer un tableau pour stocker les données
$tableauDonnees = @()

# Parcourir les utilisateurs et afficher la date d'expiration du mot de passe
foreach ($utilisateur in $utilisateurs) {
    if (!$utilisateur.PasswordNeverExpires) {
        $passwordExpiry = [System.DateTime]::FromFileTime($utilisateur.'msDS-UserPasswordExpiryTimeComputed')

        if ($passwordExpiry -ge (Get-Date) -and $passwordExpiry -le $limite) {
            # Le mot de passe va expirer dans les 15 prochains jours
            $joursRestants = ($passwordExpiry - (Get-Date)).Days
            $donnees = [PSCustomObject]@{
                Nom = $utilisateur.SamAccountName
                "Date d'expiration" = $passwordExpiry.ToShortDateString()
                "Jours restants" = $joursRestants
            }
            $tableauDonnees += $donnees
            $comptesExpirant = $true
        }
    }
}

### Message ###################################################################

$Msg = New-Object System.Net.Mail.MailMessage
 
$Msg.From = "#Compte envoyant mail"

$Msg.To.Add("#Compte receveur mail") 
    
### Message Body ######Si il y'a des comptes expirant la fonction "if" par mais si il n'ya personne la fonction "else" indiquant qu'il n'ya pas de compte expirant #####
if ($comptesExpirant) {

$MsgHTML = "<html>
            <style type=`"text/css`">
            <!--
            body {
	            background-color: #E0E0E0;
	            font-family: sans-serif;
            }
            table, th, td {
	            background-color: white;
	            border-collapse: collapse;
	            border: 1px solid black;
	            padding: 5px;
            }
            -->
            </style>
            <body>"

$MsgHTML += "Liste des comptes ayant un mot de passe expirant sous 15 jours.<br />
            <br />
            <br />
            Nom des comptes  -  Date d'expiration  -  Nombre de jours avant expiration
            <br />
            <br />"

$MsgHTML += $tableauDonnees | ConvertTo-Html -Fragment 

$MsgHTML += "<br />
            Certains comptes de services néccessite une intervention rapide.<br />
            <br />"




$MsgHTML += "</body></html>"
        }
else {
    $MsgHTML2 = "<html>
            <style type=`"text/css`">
            <!--
            body {
	            background-color: #E0E0E0;
	            font-family: sans-serif;
            }
            table, th, td {
	            background-color: white;
	            border-collapse: collapse;
	            border: 1px solid black;
	            padding: 5px;
            }
            -->
            </style>
            <body>"

$MsgHTML2 += "Liste des comptes ayant un mot de passe expirant sous 15 jours.<br />
            <br />
            <br />
            Pas de comptes expirant :)
            <br />
            <br />"

$MsgHTML2 += "</body></html>"
}        
###/Message Body ##############################################################

$Msg.Subject = "Liste des comptes ayant un mot de passe expirant sous 15 jours."   
#$Msg.Body = $MsgBodyTxt
$Msg.Body = if ($comptesExpirant) {
    $MsgHTML 
} 
else {
    $MsgHTML2
}
$Msg.BodyEncoding =$([system.text.encoding]::utf8)
#$Msg.IsBodyHtml = $false
$Msg.IsBodyHtml = $true

# IDFT SMTP

$Username = "#COMPTE SMTP"
$Password = "#MDP SMTP"
$SmtpSrv = "#SERV SMTP"
$SmtpPort = "#Port"

$Smtp = New-Object Net.Mail.SmtpClient($SmtpSrv,$SmtpPort)
$Smtp.Credentials = New-Object System.Net.NetworkCredential($Username,$Password)
$Smtp.Send($Msg)
$Msg.Dispose()
 
