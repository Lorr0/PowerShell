# Importer le module Active Directory
Import-Module ActiveDirectory

# Définir la date limite pour le contrôle des mots de passe (15 jours à partir d'aujourd'hui)
$limite = (Get-Date).AddDays(15)

# Récupérer tous les comptes d'utilisateurs de l'OU spécifiée
$utilisateurs = Get-ADUser -Filter * -SearchBase "#CHEMIN DES COMPTES A INSPECTER" -Properties "SamAccountName", "PasswordLastSet", "msDS-UserPasswordExpiryTimeComputed", "PasswordNeverExpires", "extensionAttribute8", "Description"

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
            $description = if ($utilisateur.Description) { $utilisateur.Description } else { "" }  # Vérifier si Description est nulle
            $donnees = [PSCustomObject]@{
                Nom = $utilisateur.SamAccountName.ToLower()  # Convertir en minuscules
                Description = $description  # Ajouter la description (ou une chaîne vide si nulle)
                "Date d'expiration" = $passwordExpiry.ToShortDateString()
                "Jours restants" = $joursRestants
                "extensionAttribute8" = if ($utilisateur.extensionAttribute8) { $utilisateur.extensionAttribute8.ToLower() } else { "" }  # Vérifier si extensionAttribute8 est nulle
            }
            $tableauDonnees += $donnees
            $comptesExpirant = $true
        }
    }
}

# Tri du tableau par les jours restants en ordre croissant
$tableauDonnees = $tableauDonnees | Sort-Object -Property "Jours restants"

# Créer la liste de destinataires (ajoutez les adresses e-mail ici)
$destinataires = @("adresse@email.com")

# Ajouter l'adresse e-mail additionnelle si elle existe
foreach ($donnees in $tableauDonnees) {
    if ($donnees."extensionAttribute8") {
        $destinataires += $donnees."extensionAttribute8"
    }
}

# Créer le message
$Msg = New-Object System.Net.Mail.MailMessage
$Msg.From = "adresse@email.com"

# Ajouter les destinataires (pour et CC)
foreach ($destinataire in $destinataires) {
    $Msg.To.Add($destinataire.ToLower())  # Convertir en minuscules
}

# Message Body
if ($comptesExpirant) {
    $MsgHTML = @"
<html>
    <style>
        body {
            background-color: #f5f5f5;
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 0;
        }
        .container {
            background-color: #ffffff;
            border: 1px solid #e1e1e1;
            border-radius: 5px;
            margin: 20px auto;
            max-width: 1000px; /* Augmenter la largeur du conteneur */
            padding: 20px;
        }
        h1 {
            color: #333;
        }
        p {
            color: #666;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        }
        th, td {
            border: 1px solid #e1e1e1;
            padding: 35px; /* Augmenter la taille de la cellule */
            text-align: left;
        }
    </style>
    <body>
        <div class="container">
            <h1>Liste des comptes expirant dans les 15 prochains jours</h1>
            <table>
                <tr>
                    <th>Nom des comptes</th>
                    <th>Date d'expiration</th>
                    <th>Jours restants</th>
                     <th>Description</th>
                </tr>
"@

    foreach ($donnees in $tableauDonnees) {
        $MsgHTML += @"
                <tr>
                    <td>$($donnees.Nom)</td>
                    <td>$($donnees."Date d'expiration")</td>
                    <td>$($donnees."Jours restants")</td>
                    <td>$($donnees.Description)</td>
                </tr>
"@
    }

    $MsgHTML += @"
            </table>
            <p>Certains comptes de services nécessitent une intervention rapide.</p>
        </div>
    </body>
</html>
"@
}
else {
    $MsgHTML2 = @"
<html>
    <style>
        body {
            background-color: #f5f5f5;
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 0;
        }
        .container {
            background-color: #ffffff;
            border: 1px solid #e1e1e1;
            border-radius: 5px;
            margin: 20px auto;
            max-width: 600px;
            padding: 20px;
        }
        h1 {
            color: #333;
        }
        p {
            color: #666;
        }
    </style>
    <body>
        <div class="container">
            <h1>Liste des comptes</h1>
            <p>Aucun compte n'expire dans les 15 prochains jours. Votre système est à jour.</p>
        </div>
    </body>
</html>
"@
}

# Définir le sujet du message
$Msg.Subject = "Expiration mot de passe compte SVC"

# Sélectionner le corps du message en fonction de la présence de comptes expirants
$Msg.Body = if ($comptesExpirant) {
    $MsgHTML 
} 
else {
    $MsgHTML2
}

# Définir l'encodage du corps du message
$Msg.BodyEncoding = $([system.text.encoding]::utf8)

# Indiquer que le corps du message est au format HTML
$Msg.IsBodyHtml = $true

# Configuration des paramètres SMTP
$Username = "votre_nom_d_utilisateur_SMTP"
$Password = "votre_mot_de_passe_SMTP"
$SmtpSrv = "serveur_SMTP"
$SmtpPort = "port_SMTP"

# Créer un client SMTP
$Smtp = New-Object Net.Mail.SmtpClient($SmtpSrv, $SmtpPort)
$Smtp.Credentials = New-Object System.Net.NetworkCredential($Username, $Password)

# Envoyer le message
$Smtp.Send($Msg)

# Libérer les ressources utilisées par l'objet MailMessage
$Msg.Dispose()
