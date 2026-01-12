# === CONFIGURATION ===
$mois = (Get-Date -Format "yyyy-MM")
$csvStatsPath = "C:\PRTG_STATS\prtg_errors_stats_$mois.csv"
$cheminHtmlRapport = "C:\PRTG_STATS\rapport_prtg_$mois.html"

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
<html lang="en">
<head>
    <meta charset="utf-8">
    <title>PRTG Report - $mois</title>
    <style>
        body {
            font-family: "Segoe UI", Arial, sans-serif;
            background-color: transparent;
            color: #222;
            margin: 40px;
        }

        h1 {
            font-size: 28px;
            color: #2c3e50;
            margin-bottom: 5px;
        }

        .subtitle {
            font-size: 15px;
            margin-bottom: 20px;
            color: #555;
        }

        .total-errors {
            background-color: #ffe6e6;
            border: 1px solid #ffcccc;
            padding: 10px 15px;
            font-size: 16px;
            font-weight: bold;
            color: #a00000;
            border-radius: 5px;
            display: inline-block;
            margin-bottom: 20px;
        }

        table {
            border-collapse: collapse;
            margin-top: 10px;
            background-color: #fff;
            border-radius: 6px;
            box-shadow: 0 0 10px rgba(0,0,0,0.05);
            overflow: hidden;
        }

        th, td {
            border: 1px solid #ddd;
            padding: 10px 15px;
            text-align: left;
            white-space: nowrap;
        }

        th {
            background-color: #2c3e50;
            color: #fff;
            font-weight: normal;
        }

        tr:nth-child(even) {
            background-color: #f9f9f9;
        }

        footer {
            margin-top: 40px;
            font-size: 12px;
            color: #888;
            text-align: center;
        }
    </style>
</head>
<body>
    <h1>Monthly PRTG Report - $mois</h1>
    <p class="subtitle">Summary of errors detected by PRTG this month.</p>
    <div class="total-errors">Total errors this month: $totalErreurs</div>

    <table>
        <thead>
            <tr><th>Error</th><th>Count</th></tr>
        </thead>
        <tbody>
"@

foreach ($ligne in $donneesGroupees) {
    $erreur = [System.Web.HttpUtility]::HtmlEncode($ligne.Erreur)
    $count = $ligne.Count
    $html += "            <tr><td>$erreur</td><td>$count</td></tr>`n"
}

$html += @"
        </tbody>
    </table>

    <footer><p>Generated on $(Get-Date -Format "yyyy-MM-dd HH:mm")</p></footer>
</body>
</html>
"@


# === SAUVEGARDE DU HTML ===
$utf8BOM = New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllText($cheminHtmlRapport, $html, $utf8BOM)

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
