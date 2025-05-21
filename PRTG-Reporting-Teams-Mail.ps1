# === CONFIGURATION ===
$mois = (Get-Date -Format "yyyy-MM")
$csvStatsPath = "C:\PRTG_STATS\prtg_errors_stats_$mois.csv"
$cheminHtmlRapport = "C:\PRTG_STATS\rapport_prtg_$mois.html"
$dossierArchives = "C:\PRTG_STATS\Archives"

# === VERIFICATION DU CSV ===
if (!(Test-Path $csvStatsPath)) {
    Write-Error "Le fichier $csvStatsPath est introuvable."
    exit 1
}

# === IMPORT ET GROUPEMENT DES DONNEES ===
$donnees = Import-Csv -Path $csvStatsPath

$donneesGroupees = $donnees | Group-Object -Property Erreur | ForEach-Object {
    [PSCustomObject]@{
        Erreur = $_.Name
        Count  = ($_.Group | Measure-Object -Property Count -Sum).Sum
    }
}

$totalErreurs = ($donneesGroupees | Measure-Object -Property Count -Sum).Sum

# === GENERATION DU HTML SANS ACCENTS ===
$html = @"
<!DOCTYPE html>
<html lang=\"fr\">
<head>
    <meta charset=\"utf-8\">
    <title>Rapport PRTG - $mois</title>
    <style>
        body { font-family: Arial, sans-serif; background-color: #ffffff; color: #222; margin: 40px; }
        h1 { font-size: 24px; color: #333; }
        .subtitle { font-size: 14px; margin-bottom: 20px; color: #555; }
        table { border-collapse: collapse; margin-top: 10px; }
        th, td { border: 1px solid #ccc; padding: 10px 15px; text-align: left; white-space: nowrap; }
        th { background-color: #2c3e50; color: #fff; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        tfoot td { font-weight: bold; border-top: 2px solid #2c3e50; }
        footer { margin-top: 40px; font-size: 12px; color: #888; }
    </style>
</head>
<body>
    <h1>Rapport mensuel PRTG - $mois</h1>
    <p class=\"subtitle\">Resume des erreurs detectees ce mois-ci par PRTG.</p>
    <table>
        <thead><tr><th>Erreur</th><th>Nombre</th></tr></thead>
        <tbody>
"@

foreach ($ligne in $donneesGroupees) {
    $erreur = [System.Web.HttpUtility]::HtmlEncode($ligne.Erreur)
    $count = $ligne.Count
    $html += "            <tr><td>$erreur</td><td>$count</td></tr>`n"
}

$html += @"
        </tbody>
        <tfoot>
            <tr><td>Total des erreurs</td><td>$totalErreurs</td></tr>
        </tfoot>
    </table>
    <footer>Genere le $(Get-Date -Format "yyyy-MM-dd HH:mm")</footer>
</body>
</html>
"@

# === SAUVEGARDE DU HTML ===
$utf8BOM = New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllText($cheminHtmlRapport, $html, $utf8BOM)

# === ARCHIVAGE ===
if (!(Test-Path $dossierArchives)) {
    New-Item -Path $dossierArchives -ItemType Directory | Out-Null
}
$cheminArchive = Join-Path $dossierArchives "rapport_prtg_$mois.html"
Copy-Item -Path $cheminHtmlRapport -Destination $cheminArchive -Force

$destinataires = @(
    "MAIL 1",
    "MAIL 2"
)

Send-MailMessage `
    -From "MAIL OWNER" `
    -To $destinataires `
    -Subject "[PRTG] Rapport mensuel - $mois" `
    -BodyAsHtml $html `
    -SmtpServer "smtps.office365" `
    -Port xx `
    -UseSsl 

Write-Host "✅ Rapport HTML sauvegarde, archive et envoye par e-mail."
