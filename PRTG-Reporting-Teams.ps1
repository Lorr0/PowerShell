# === PARAMÈTRES À DÉFINIR ===
$csvFolder = "C:\chemin\vers\dossierCSV"
$prtgServer = "https://prtg.example.com"
$prtgUser = "utilisateur"
$prtgPasshash = "votre_passhash"
$webhookUrl = "https://webhook.teams..."

# === CRÉATION DU DOSSIER SI NÉCESSAIRE ===
if (-not (Test-Path $csvFolder)) {
    New-Item -ItemType Directory -Path $csvFolder | Out-Null
}

# === DATES ===
$moisActuel = Get-Date -Format "yyyy-MM"
$dateDuJour = Get-Date -Format "yyyy-MM-dd"
$csvStatsPath = Join-Path $csvFolder "prtg_errors_stats_$moisActuel.csv"

# === CAPTEURS À EXCLURE (INSENSIBLE À LA CASSE) ===
$excludedSensors = @(
    "Memory: Physical Memory",
    "Charge CPU",
    "Memory",
    "Disponibilité du système Windows",
    "Statut des mises à jour Windows"
) | ForEach-Object { $_.ToLower() }

# === RÉCUPÉRATION DES CAPTEURS EN ERREUR ===
$uri = "$prtgServer/api/table.json?content=sensors&output=json&filter_status=5&username=$prtgUser&passhash=$prtgPasshash"

try {
    $response = Invoke-RestMethod -Uri $uri -Method Get -UseBasicParsing
    $sensors = $response.sensors
} catch {
    Write-Error "❌ Erreur lors de la récupération des capteurs PRTG : $_"
    exit 1
}

# === FILTRAGE DES CAPTEURS EXCLUS ===
$sensors = $sensors | Where-Object {
    $excludedSensors -notcontains $_.sensor.ToLower()
}

# === GÉNÉRATION DU MESSAGE ===
if ($sensors.Count -eq 0) {
    $message = "✅ Aucun capteur en erreur actuellement."
} else {
    $listeCapteurs = $sensors | ForEach-Object {
        "❌ $($_.device) - $($_.sensor)"
    }
    $message = $listeCapteurs -join "||"
}

# === ENVOI DU MESSAGE DANS TEAMS ===
$body = @{
    title = "🚨 Capteurs PRTG en erreur"
    text  = $message
}
$jsonBody = $body | ConvertTo-Json -Depth 5
$utf8NoBom = New-Object System.Text.UTF8Encoding $false
$bodyBytes = $utf8NoBom.GetBytes($jsonBody)

try {
    Invoke-RestMethod -Uri $webhookUrl -Method Post -Body $bodyBytes -ContentType 'application/json'
    Write-Host "✅ Message envoyé dans Teams."
} catch {
    Write-Error "❌ Erreur lors de l'envoi vers Teams : $_"
}

# === SAUVEGARDE DANS LE CSV UNIQUEMENT SI ERREURS ===
if ($sensors.Count -eq 0) {
    Write-Host "✅ Aucun capteur en erreur. Aucune écriture dans le CSV."
    return
}

# Nettoyage des erreurs
$erreurs = $message -split '\|\|' | ForEach-Object { $_.Trim().TrimStart('❌').Trim() }

# Chargement du CSV existant ou création d’un tableau vide
if (Test-Path $csvStatsPath) {
    $stats = Import-Csv -Path $csvStatsPath
} else {
    $stats = @()
}

# Mise à jour des compteurs (éviter les doublons du jour)
foreach ($err in $erreurs) {
    $entry = $stats | Where-Object { $_.Erreur -eq $err -and $_.Date -eq $dateDuJour }
    if (-not $entry) {
        $stats += [PSCustomObject]@{
            Erreur = $err
            Date   = $dateDuJour
            Count  = 1
        }
    }
}

# Sauvegarde du fichier
$stats | Sort-Object Date, Erreur | Export-Csv -Path $csvStatsPath -NoTypeInformation -Encoding UTF8
Write-Host "📁 Statistiques enregistrées dans : $csvStatsPath"
