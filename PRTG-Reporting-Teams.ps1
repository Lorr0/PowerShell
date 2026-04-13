#Requires -Version 5.1
<#
.SYNOPSIS
    Reporting PRTG → Microsoft Teams via Webhook (Workflow Teams)
.DESCRIPTION
    Récupère les capteurs PRTG en erreur, envoie un rapport structuré dans Teams
    et alimente un CSV de suivi mensuel.
.NOTES
    Auteur  : Gregory B. — MRN / Infrastructure & Systèmes
    Version : 2.0
#>

# ╔══════════════════════════════════════════════════════════════╗
# ║  PARAMÈTRES                                                  ║
# ╚══════════════════════════════════════════════════════════════╝

# --- Authentification PRTG (variables à définir en amont) ---
$csvFolder = "C:\chemin\vers\dossierCSV"
$prtgServer = "https://prtg.example.com"
$prtgUser = "utilisateur"
$prtgPasshash = "votre_passhash"
$webhookUrl = "https://webhook.teams..."
 
# ╔══════════════════════════════════════════════════════════════╗
# ║  CONFIGURATION DU FILTRAGE                                   ║
# ╚══════════════════════════════════════════════════════════════╝

$excludedSensorNames = @(
    "Charge CPU"
    "Memory"
    "Mémoire"
    "Memory: Physical Memory"
    "Disponibilité du système Windows"
    "Statut des mises à jour Windows"
    "Disponibilité"
    "Probe Health"
    "System Health"
    "Core Health"
)

$excludedPatterns = @(
    "*snmp*"
    "*heartbeat*"
    "*probe*health*"
)

$excludedGroups = @(
    # "Labo"
    # "Tests"
)

# Priorité minimale à remonter (1=basse → 5=critique). 0 = tout afficher
$minPriority = 0

# ╔══════════════════════════════════════════════════════════════╗
# ║  INITIALISATION                                              ║
# ╚══════════════════════════════════════════════════════════════╝

if (-not (Test-Path $csvFolder)) {
    New-Item -ItemType Directory -Path $csvFolder -Force | Out-Null
}

$now          = Get-Date
$moisActuel   = $now.ToString("yyyy-MM")
$dateDuJour   = $now.ToString("yyyy-MM-dd")
$heureDuJour  = $now.ToString("HH:mm")
$csvStatsPath = Join-Path $csvFolder "prtg_errors_stats_$moisActuel.csv"

# ╔══════════════════════════════════════════════════════════════╗
# ║  REQUÊTE API PRTG — CAPTEURS EN ERREUR (status=5)           ║
# ╚══════════════════════════════════════════════════════════════╝

$columns = "objid,device,sensor,group,status,lastvalue,priority,downtimesince"
$uri = "$prtgServer/api/table.json" +
       "?content=sensors" +
       "&output=json" +
       "&columns=$columns" +
       "&filter_status=5" +
       "&sortby=priority" +
       "&username=$prtgUser" +
       "&passhash=$prtgPasshash"

try {
    $response = Invoke-RestMethod -Uri $uri -Method Get -UseBasicParsing
    $sensors  = $response.sensors
    Write-Host "[OK] $($sensors.Count) capteur(s) en erreur brut(s) récupéré(s)."
} catch {
    Write-Error "Erreur lors de la récupération des capteurs PRTG : $_"
    exit 1
}

# ╔══════════════════════════════════════════════════════════════╗
# ║  FILTRAGE AVANCÉ                                             ║
# ╚══════════════════════════════════════════════════════════════╝

$excludedLower = $excludedSensorNames | ForEach-Object { $_.ToLower() }
$excludedGroupsLower = $excludedGroups | ForEach-Object { $_.ToLower() }

$filtered = $sensors | Where-Object {
    $name  = $_.sensor.ToLower()
    $group = if ($_.group) { $_.group.ToLower() } else { "" }
    $prio  = if ($_.priority_raw) { [int]$_.priority_raw } else { 0 }

    if ($excludedLower -contains $name) { return $false }

    foreach ($pattern in $excludedPatterns) {
        if ($name -like $pattern) { return $false }
    }

    if ($excludedGroupsLower.Count -gt 0 -and $excludedGroupsLower -contains $group) {
        return $false
    }

    if ($minPriority -gt 0 -and $prio -lt $minPriority) { return $false }

    return $true
}

Write-Host "[OK] $($filtered.Count) capteur(s) après filtrage."

# ╔══════════════════════════════════════════════════════════════╗
# ║  HELPERS                                                     ║
# ╚══════════════════════════════════════════════════════════════╝

function Get-PrioIcon ([int]$prio) {
    switch ($prio) {
        5       { return "🔴" }
        4       { return "🟠" }
        3       { return "🟡" }
        default { return "⚪" }
    }
}

function Get-PrioLabel ([int]$prio) {
    switch ($prio) {
        5       { return "CRITIQUE" }
        4       { return "HAUTE" }
        3       { return "MOYENNE" }
        default { return "BASSE" }
    }
}

# ╔══════════════════════════════════════════════════════════════╗
# ║  CONSTRUCTION DU MESSAGE TEAMS                               ║
# ╚══════════════════════════════════════════════════════════════╝

# Séparateurs visuels Unicode (texte brut)
$sepH = "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
$sepL = "─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─"

if ($filtered.Count -eq 0) {

    # ────── AUCUNE ALERTE ──────
    $title = "✅ PRTG — RAS"
    $lines = @(
        $sepH
        "📊  ETAT DU MONITORING"
        $sepH
        ""
        "✅  Tous les capteurs sont operationnels."
        ""
        $sepL
        "🕐  Rapport du $dateDuJour a ${heureDuJour}"
    )
    $text = $lines -join "||"

} else {

    # ────── ALERTES DÉTECTÉES ──────
    $grouped = $filtered | Group-Object -Property device | Sort-Object Name

    # Compteurs par sévérité
    $nbCritique = @($filtered | Where-Object { [int]$_.priority_raw -ge 5 }).Count
    $nbHaute    = @($filtered | Where-Object { [int]$_.priority_raw -eq 4 }).Count
    $nbMoyenne  = @($filtered | Where-Object { [int]$_.priority_raw -eq 3 }).Count
    $nbBasse    = @($filtered | Where-Object { [int]$_.priority_raw -le 2 }).Count
    $nbDevices  = $grouped.Count

    $lines = @()

    # ── BLOC RÉSUMÉ ──
    $lines += $sepH
    $lines += "📊  RESUME"
    $lines += $sepH
    $lines += ""
    $lines += "⚠️  $($filtered.Count) alerte(s) detectee(s) sur $nbDevices equipement(s)"
    $lines += ""

    # Jauge de sévérité (uniquement catégories non-vides)
    $sevParts = @()
    if ($nbCritique -gt 0) { $sevParts += "🔴 Critique : $nbCritique" }
    if ($nbHaute -gt 0)    { $sevParts += "🟠 Haute : $nbHaute" }
    if ($nbMoyenne -gt 0)  { $sevParts += "🟡 Moyenne : $nbMoyenne" }
    if ($nbBasse -gt 0)    { $sevParts += "⚪ Basse : $nbBasse" }
    $lines += $sevParts -join "   ·   "

    $lines += ""
    $lines += "🕐  $dateDuJour a ${heureDuJour}"
    $lines += $sepH

    # ── DÉTAIL PAR DEVICE ──
    foreach ($grp in $grouped) {
        $deviceName = $grp.Name
        $nbAlertesDevice = $grp.Count

        $sortedSensors = @($grp.Group | Sort-Object {
            if ($_.priority_raw) { -[int]$_.priority_raw } else { 0 }
        })

        $lines += ""
        $lines += "🖥️ $deviceName ($nbAlertesDevice alerte(s))"

        for ($i = 0; $i -lt $sortedSensors.Count; $i++) {
            $s = $sortedSensors[$i]
            $prioRaw = if ($s.priority_raw) { [int]$s.priority_raw } else { 0 }
            $prioStars = ("★" * $prioRaw) + ("☆" * (5 - $prioRaw))

            # Connecteur arbre : └ pour le dernier, ├ pour les autres
            $connector = if ($i -eq $sortedSensors.Count - 1) { "└" } else { "├" }

            # Downtime
            $downtime = if ($s.downtimesince -and $s.downtimesince -ne "") {
                "  ⏱️ $($s.downtimesince)"
            } else { "" }

            # Dernière valeur
            $lastVal = if ($s.lastvalue -and $s.lastvalue -ne "" -and $s.lastvalue -ne "-") {
                "  📈 $($s.lastvalue)"
            } else { "" }

            $lines += "$connector ❌ $($s.sensor) [$prioStars]${downtime}${lastVal}"
        }
    }

    # ── FOOTER ──
    $lines += ""
    $lines += $sepH

    $title = "🚨 PRTG — $($filtered.Count) alerte(s) sur $nbDevices equipement(s)"
    $text  = $lines -join "||"
}

# ╔══════════════════════════════════════════════════════════════╗
# ║  ENVOI VERS TEAMS (Webhook Workflow)                         ║
# ╚══════════════════════════════════════════════════════════════╝

$body = @{
    title = $title
    text  = $text
} | ConvertTo-Json -Depth 5

$utf8NoBom = New-Object System.Text.UTF8Encoding $false
$bodyBytes = $utf8NoBom.GetBytes($body)

try {
    Invoke-RestMethod -Uri $webhookUrl -Method Post -Body $bodyBytes -ContentType 'application/json; charset=utf-8'
    Write-Host "[OK] Message envoyé dans Teams."
} catch {
    Write-Error "Erreur lors de l'envoi vers Teams : $_"
    exit 1
}

# ╔══════════════════════════════════════════════════════════════╗
# ║  SUIVI CSV MENSUEL                                           ║
# ╚══════════════════════════════════════════════════════════════╝

if ($filtered.Count -eq 0) {
    Write-Host "[OK] Aucune erreur — pas d'écriture CSV."
    exit 0
}

if (Test-Path $csvStatsPath) {
    $stats = @(Import-Csv -Path $csvStatsPath)
} else {
    $stats = @()
}

foreach ($s in $filtered) {
    $errLabel = "$($s.device) - $($s.sensor)"
    $entry = $stats | Where-Object { $_.Erreur -eq $errLabel -and $_.Date -eq $dateDuJour }

    if ($entry) {
        $entry.Occurrences = [int]$entry.Occurrences + 1
    } else {
        $prioVal = if ($s.priority_raw) { [int]$s.priority_raw } else { 0 }
        $stats += [PSCustomObject]@{
            Date        = $dateDuJour
            Heure       = $heureDuJour
            Device      = $s.device
            Capteur     = $s.sensor
            Groupe      = $s.group
            Priorite    = $prioVal
            Erreur      = $errLabel
            Occurrences = 1
        }
    }
}

$stats | Sort-Object Date, Device, Capteur |
    Export-Csv -Path $csvStatsPath -NoTypeInformation -Encoding UTF8

Write-Host "[OK] Statistiques enregistrées : $csvStatsPath"
