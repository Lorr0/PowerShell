# === CONFIGURATION ===
$mois = (Get-Date -Format "yyyy-MM")
$moisAffichage = (Get-Culture).DateTimeFormat.GetMonthName((Get-Date).Month) + " " + (Get-Date).Year
$csvStatsPath = "C:\PRTG_STATS\prtg_errors_stats_$mois.csv"
$cheminHtmlRapport = "C:\HTML_RAPPORT\rapport_prtg_$mois.html"

Add-Type -AssemblyName System.Web

# === VERIFICATION ===
if (!(Test-Path $csvStatsPath)) {
    Write-Error "Fichier introuvable : $csvStatsPath"
    exit 1
}

$donnees = Import-Csv -Path $csvStatsPath
if ($donnees.Count -eq 0) {
    Write-Host "CSV vide." -ForegroundColor Yellow
    exit 0
}

# === PREPARATION DES DONNEES ===
function Normalize-Name ([string]$nom) {
    return $nom.Trim().ToLower() -replace '\.xxx\.xxx$', '' #domaine exemple : cesi.lan
}

$records = foreach ($row in $donnees) {
    $occ  = if ($row.Occurrences) { [int]$row.Occurrences } else { 1 }
    $prio = if ($row.Priorite)    { [int]$row.Priorite }    else { 0 }

    if ($row.Device -and $row.Capteur) {
        $device  = $row.Device.Trim()
        $capteur = $row.Capteur.Trim()
        $groupe  = if ($row.Groupe) { $row.Groupe.Trim() } else { "Non defini" }
    } else {
        $parts   = $row.Erreur -split ' - ', 2
        $device  = $parts[0].Trim()
        $capteur = if ($parts.Count -gt 1) { $parts[1].Trim() } else { "Inconnu" }
        $groupe  = "Non defini"
    }

    [PSCustomObject]@{
        Date        = [datetime]$row.Date
        Device      = $device
        DeviceNorm  = Normalize-Name $device
        Capteur     = $capteur
        Groupe      = $groupe
        Priorite    = $prio
        Occurrences = $occ
    }
}

# === DEDUPLICATION : 1 alerte max par device + capteur + jour ===
# Si la meme alerte apparait plusieurs fois dans la journee, on ne la compte qu'une fois
$records = $records |
    Group-Object @{E={"$($_.DeviceNorm)|$($_.Capteur)|$($_.Date.ToString('yyyy-MM-dd'))"}} |
    ForEach-Object {
        # On garde la pire priorite de la journee pour cette alerte
        $best = $_.Group | Sort-Object Priorite -Descending | Select-Object -First 1
        [PSCustomObject]@{
            Date        = $best.Date
            Device      = $best.Device
            DeviceNorm  = $best.DeviceNorm
            Capteur     = $best.Capteur
            Groupe      = $best.Groupe
            Priorite    = $best.Priorite
            Occurrences = 1
        }
    }

# === CALCUL DES INDICATEURS ===

$totalAlertes      = ($records | Measure-Object -Property Occurrences -Sum).Sum
$nbJoursMois       = [DateTime]::DaysInMonth((Get-Date).Year, (Get-Date).Month)
$joursAvecAlertes  = ($records | ForEach-Object { $_.Date.Date } | Sort-Object -Unique).Count
$joursSansAlerte   = $nbJoursMois - $joursAvecAlertes
$moyParJour        = if ($joursAvecAlertes -gt 0) { [math]::Round($totalAlertes / $joursAvecAlertes, 1) } else { 0 }
$nbDevices         = ($records | Select-Object -Property DeviceNorm -Unique).Count

# Severite
$sevCounts = @{
    Critique = 0; Haute = 0; Moyenne = 0; Basse = 0
}
foreach ($r in $records) {
    switch ([int]$r.Priorite) {
        { $_ -ge 5 } { $sevCounts.Critique += $r.Occurrences }
        4             { $sevCounts.Haute    += $r.Occurrences }
        3             { $sevCounts.Moyenne  += $r.Occurrences }
        default       { $sevCounts.Basse   += $r.Occurrences }
    }
}

# Top equipements
$topDevices = $records |
    Group-Object DeviceNorm |
    ForEach-Object {
        $nom = ($_.Group | Group-Object Device | Sort-Object Count -Descending | Select-Object -First 1).Name
        $worst = ($_.Group | Measure-Object -Property Priorite -Maximum).Maximum
        [PSCustomObject]@{
            Nom       = $nom
            Total     = ($_.Group | Measure-Object -Property Occurrences -Sum).Sum
            NbCapteurs = ($_.Group | Select-Object -Property Capteur -Unique).Count
            PrioMax   = $worst
        }
    } |
    Sort-Object Total -Descending |
    Select-Object -First 5

# Top capteurs
$topCapteurs = $records |
    Group-Object @{E={"$($_.DeviceNorm)|$($_.Capteur)"}} |
    ForEach-Object {
        $first = $_.Group[0]
        [PSCustomObject]@{
            Device  = $first.Device
            Capteur = $first.Capteur
            Total   = ($_.Group | Measure-Object -Property Occurrences -Sum).Sum
            NbJours = ($_.Group | ForEach-Object { $_.Date.Date } | Sort-Object -Unique).Count
            PrioMax = ($_.Group | Measure-Object -Property Priorite -Maximum).Maximum
        }
    } |
    Sort-Object Total -Descending |
    Select-Object -First 5

# Repartition par groupe
$parGroupe = $records |
    Group-Object Groupe |
    ForEach-Object {
        [PSCustomObject]@{
            Groupe = $_.Name
            Total  = ($_.Group | Measure-Object -Property Occurrences -Sum).Sum
        }
    } | Sort-Object Total -Descending

# === FONCTIONS HTML ===
function Get-PrioColor ([int]$p) {
    switch ($p) { 5 {"#c0392b"} 4 {"#e67e22"} 3 {"#f39c12"} default {"#7f8c8d"} }
}
function Get-PrioLabel ([int]$p) {
    switch ($p) { 5 {"Critique"} 4 {"Haute"} 3 {"Moyenne"} default {"Basse"} }
}
function Get-CountBg ([int]$c) {
    if ($c -ge 30) {"#c0392b"} elseif ($c -ge 15) {"#e67e22"} elseif ($c -ge 5) {"#f39c12"} else {"#27ae60"}
}
function Get-CountFg ([int]$c) {
    if ($c -ge 5 -and $c -lt 15) {"#000000"} else {"#ffffff"}
}
function Enc ([string]$t) { [System.Web.HttpUtility]::HtmlEncode($t) }

# === CONSTRUCTION HTML ===

# -- Lignes top equipements --
$devRows = ""
$alt = $false
foreach ($dv in $topDevices) {
    $bg     = if ($alt) {"#f8f9fa"} else {"#ffffff"}
    $pColor = Get-PrioColor $dv.PrioMax
    $pLabel = Get-PrioLabel $dv.PrioMax
    $cBg    = Get-CountBg $dv.Total
    $cFg    = Get-CountFg $dv.Total
    $pct    = if ($totalAlertes -gt 0) { [math]::Round(($dv.Total / $totalAlertes) * 100, 1) } else { 0 }
    $barW   = [math]::Min([math]::Round($pct * 2), 100)
    $nom    = Enc $dv.Nom

    $devRows += @"
<tr style="background-color:$bg;">
<td style="padding:10px 14px;border-bottom:1px solid #ecf0f1;font-weight:bold;color:#2c3e50;font-size:13px;">$nom</td>
<td style="padding:10px 14px;border-bottom:1px solid #ecf0f1;text-align:center;">
<table cellpadding="0" cellspacing="0" border="0" style="margin:0 auto;"><tr><td style="background-color:$cBg;color:$cFg;padding:4px 12px;font-weight:bold;font-size:12px;">$($dv.Total)</td></tr></table>
</td>
<td style="padding:10px 14px;border-bottom:1px solid #ecf0f1;text-align:center;color:#7f8c8d;font-size:13px;">$($dv.NbCapteurs)</td>
<td style="padding:10px 14px;border-bottom:1px solid #ecf0f1;text-align:center;color:$pColor;font-weight:bold;font-size:12px;">$pLabel</td>
<td style="padding:10px 14px;border-bottom:1px solid #ecf0f1;">
<table cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td style="background-color:#ecf0f1;"><table cellpadding="0" cellspacing="0" border="0" width="${barW}%"><tr><td style="background-color:$cBg;height:8px;font-size:1px;">&nbsp;</td></tr></table></td></tr></table>
<div style="font-size:10px;color:#7f8c8d;text-align:center;padding-top:3px;">${pct}%</div>
</td>
</tr>
"@
    $alt = -not $alt
}

# -- Lignes top capteurs --
$capRows = ""
$alt = $false
foreach ($cp in $topCapteurs) {
    $bg     = if ($alt) {"#f8f9fa"} else {"#ffffff"}
    $pColor = Get-PrioColor $cp.PrioMax
    $pLabel = Get-PrioLabel $cp.PrioMax
    $cBg    = Get-CountBg $cp.Total
    $cFg    = Get-CountFg $cp.Total

    $capRows += @"
<tr style="background-color:$bg;">
<td style="padding:10px 14px;border-bottom:1px solid #ecf0f1;">
<div style="font-weight:bold;color:#2c3e50;font-size:13px;">$(Enc $cp.Capteur)</div>
<div style="font-size:11px;color:#95a5a6;padding-top:2px;">$(Enc $cp.Device)</div>
</td>
<td style="padding:10px 14px;border-bottom:1px solid #ecf0f1;text-align:center;">
<table cellpadding="0" cellspacing="0" border="0" style="margin:0 auto;"><tr><td style="background-color:$cBg;color:$cFg;padding:4px 12px;font-weight:bold;font-size:12px;">$($cp.Total)</td></tr></table>
</td>
<td style="padding:10px 14px;border-bottom:1px solid #ecf0f1;text-align:center;color:#7f8c8d;font-size:13px;">$($cp.NbJours) j.</td>
<td style="padding:10px 14px;border-bottom:1px solid #ecf0f1;text-align:center;color:$pColor;font-weight:bold;font-size:12px;">$pLabel</td>
</tr>
"@
    $alt = -not $alt
}

# -- Lignes groupes --
$grpRows = ""
$alt = $false
foreach ($g in $parGroupe) {
    $bg   = if ($alt) {"#f8f9fa"} else {"#ffffff"}
    $pct  = if ($totalAlertes -gt 0) { [math]::Round(($g.Total / $totalAlertes) * 100, 1) } else { 0 }
    $barW = [math]::Min([math]::Round($pct * 2), 100)

    $grpRows += @"
<tr style="background-color:$bg;">
<td style="padding:8px 14px;border-bottom:1px solid #ecf0f1;font-weight:bold;color:#2c3e50;font-size:13px;">$(Enc $g.Groupe)</td>
<td style="padding:8px 14px;border-bottom:1px solid #ecf0f1;text-align:center;font-weight:bold;font-size:13px;">$($g.Total)</td>
<td style="padding:8px 14px;border-bottom:1px solid #ecf0f1;">
<table cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td style="background-color:#ecf0f1;"><table cellpadding="0" cellspacing="0" border="0" width="${barW}%"><tr><td style="background-color:#2980b9;height:8px;font-size:1px;">&nbsp;</td></tr></table></td></tr></table>
<div style="font-size:10px;color:#7f8c8d;text-align:center;padding-top:3px;">${pct}%</div>
</td>
</tr>
"@
    $alt = -not $alt
}

# ╔══════════════════════════════════════════════════════════════╗
# ║  HTML FINAL                                                  ║
# ╚══════════════════════════════════════════════════════════════╝

$html = @"
<!DOCTYPE html>
<html lang="fr">
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Rapport PRTG - $mois</title></head>
<body style="margin:0;padding:0;font-family:Arial,Helvetica,sans-serif;background-color:#ecf0f1;">
<table cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#ecf0f1;padding:20px 0;">
<tr><td align="center">
<table cellpadding="0" cellspacing="0" border="0" width="680" style="max-width:680px;">

<!-- EN-TETE -->
<tr><td style="background-color:#1a252f;padding:28px 30px;">
<table cellpadding="0" cellspacing="0" border="0" width="100%"><tr>
<td style="vertical-align:top;">
<div style="font-size:11px;color:#7f8c8d;text-transform:uppercase;letter-spacing:2px;">RAPPORT SUPERVISION</div>
<div style="font-size:24px;font-weight:bold;color:#ffffff;padding-top:8px;">Rapport de supervision PRTG</div>
<div style="font-size:13px;color:#95a5a6;padding-top:4px;">Bilan mensuel - $moisAffichage</div>
</td>
<td style="vertical-align:top;text-align:right;">
<table cellpadding="0" cellspacing="0" border="0"><tr><td style="background-color:#2c3e50;padding:10px 14px;">
<div style="font-size:11px;color:#95a5a6;">Infrastructure &amp; Systemes</div>
<div style="font-size:11px;color:#95a5a6;padding-top:2px;">DSI</div>
</td></tr></table>
</td>
</tr></table>
</td></tr>

<!-- CORPS -->
<tr><td style="background-color:#ffffff;padding:28px 30px;">

<!-- KPIs -->
<table cellpadding="0" cellspacing="0" border="0" width="100%"><tr>
<td width="25%" style="padding-right:6px;">
<table cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td style="background-color:#fdecea;border-top:3px solid #c0392b;padding:16px 10px;text-align:center;">
<div style="font-size:28px;font-weight:bold;color:#c0392b;">$totalAlertes</div>
<div style="font-size:11px;color:#7f8c8d;padding-top:4px;">Total alertes</div>
</td></tr></table>
</td>
<td width="25%" style="padding:0 3px;">
<table cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td style="background-color:#fef5e7;border-top:3px solid #e67e22;padding:16px 10px;text-align:center;">
<div style="font-size:28px;font-weight:bold;color:#e67e22;">$nbDevices</div>
<div style="font-size:11px;color:#7f8c8d;padding-top:4px;">Equipements</div>
</td></tr></table>
</td>
<td width="25%" style="padding:0 3px;">
<table cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td style="background-color:#eafaf1;border-top:3px solid #27ae60;padding:16px 10px;text-align:center;">
<div style="font-size:28px;font-weight:bold;color:#27ae60;">$joursSansAlerte<span style="font-size:14px;font-weight:normal;color:#7f8c8d;">/$nbJoursMois</span></div>
<div style="font-size:11px;color:#7f8c8d;padding-top:4px;">Jours sans alerte</div>
</td></tr></table>
</td>
<td width="25%" style="padding-left:6px;">
<table cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td style="background-color:#eaf2f8;border-top:3px solid #2980b9;padding:16px 10px;text-align:center;">
<div style="font-size:28px;font-weight:bold;color:#2980b9;">$moyParJour</div>
<div style="font-size:11px;color:#7f8c8d;padding-top:4px;">Moy./jour actif</div>
</td></tr></table>
</td>
</tr></table>

<!-- SEVERITE -->
<table cellpadding="0" cellspacing="0" border="0" width="100%" style="padding-top:24px;">
<tr><td style="font-size:15px;font-weight:bold;color:#2c3e50;padding-bottom:10px;">Repartition par severite</td></tr>
<tr><td>
<table cellpadding="0" cellspacing="6" border="0" width="100%"><tr>
<td width="25%" style="background-color:#c0392b;padding:12px 8px;text-align:center;">
<div style="font-size:24px;font-weight:bold;color:#ffffff;">$($sevCounts.Critique)</div>
<div style="font-size:11px;color:#ffffff;padding-top:2px;">Critique (P5)</div>
</td>
<td width="25%" style="background-color:#e67e22;padding:12px 8px;text-align:center;">
<div style="font-size:24px;font-weight:bold;color:#ffffff;">$($sevCounts.Haute)</div>
<div style="font-size:11px;color:#ffffff;padding-top:2px;">Haute (P4)</div>
</td>
<td width="25%" style="background-color:#f39c12;padding:12px 8px;text-align:center;">
<div style="font-size:24px;font-weight:bold;color:#000000;">$($sevCounts.Moyenne)</div>
<div style="font-size:11px;color:#000000;padding-top:2px;">Moyenne (P3)</div>
</td>
<td width="25%" style="background-color:#95a5a6;padding:12px 8px;text-align:center;">
<div style="font-size:24px;font-weight:bold;color:#ffffff;">$($sevCounts.Basse)</div>
<div style="font-size:11px;color:#ffffff;padding-top:2px;">Basse (P1-2)</div>
</td>
</tr></table>
</td></tr>
</table>

<!-- TOP EQUIPEMENTS -->
<table cellpadding="0" cellspacing="0" border="0" width="100%" style="padding-top:24px;">
<tr><td style="font-size:15px;font-weight:bold;color:#2c3e50;padding-bottom:10px;">Top 5 - Equipements les plus alertes</td></tr>
<tr><td>
<table cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;">
<tr style="background-color:#2c3e50;">
<td style="padding:10px 14px;color:#ffffff;font-size:11px;font-weight:bold;">EQUIPEMENT</td>
<td style="padding:10px 14px;color:#ffffff;font-size:11px;font-weight:bold;text-align:center;">ALERTES</td>
<td style="padding:10px 14px;color:#ffffff;font-size:11px;font-weight:bold;text-align:center;">CAPTEURS</td>
<td style="padding:10px 14px;color:#ffffff;font-size:11px;font-weight:bold;text-align:center;">SEV. MAX</td>
<td style="padding:10px 14px;color:#ffffff;font-size:11px;font-weight:bold;text-align:center;">PART</td>
</tr>
$devRows
</table>
</td></tr>
</table>

<!-- TOP CAPTEURS -->
<table cellpadding="0" cellspacing="0" border="0" width="100%" style="padding-top:24px;">
<tr><td style="font-size:15px;font-weight:bold;color:#2c3e50;padding-bottom:10px;">Top 5 - Capteurs les plus recurrents</td></tr>
<tr><td>
<table cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;">
<tr style="background-color:#2c3e50;">
<td style="padding:10px 14px;color:#ffffff;font-size:11px;font-weight:bold;">CAPTEUR / EQUIPEMENT</td>
<td style="padding:10px 14px;color:#ffffff;font-size:11px;font-weight:bold;text-align:center;">ALERTES</td>
<td style="padding:10px 14px;color:#ffffff;font-size:11px;font-weight:bold;text-align:center;">JOURS</td>
<td style="padding:10px 14px;color:#ffffff;font-size:11px;font-weight:bold;text-align:center;">SEV. MAX</td>
</tr>
$capRows
</table>
</td></tr>
</table>

<!-- GROUPES -->
<table cellpadding="0" cellspacing="0" border="0" width="100%" style="padding-top:24px;">
<tr><td style="font-size:15px;font-weight:bold;color:#2c3e50;padding-bottom:10px;">Repartition par groupe PRTG</td></tr>
<tr><td>
<table cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;">
<tr style="background-color:#f8f9fa;">
<td style="padding:10px 14px;font-size:11px;font-weight:bold;color:#7f8c8d;border-bottom:2px solid #ecf0f1;">GROUPE</td>
<td style="padding:10px 14px;font-size:11px;font-weight:bold;color:#7f8c8d;text-align:center;border-bottom:2px solid #ecf0f1;">ALERTES</td>
<td style="padding:10px 14px;font-size:11px;font-weight:bold;color:#7f8c8d;border-bottom:2px solid #ecf0f1;">PART</td>
</tr>
$grpRows
</table>
</td></tr>
</table>

</td></tr>

<!-- PIED DE PAGE -->
<tr><td style="background-color:#1a252f;padding:18px 30px;">
<table cellpadding="0" cellspacing="0" border="0" width="100%"><tr>
<td style="font-size:11px;color:#7f8c8d;">Genere le $(Get-Date -Format "dd/MM/yyyy") a $(Get-Date -Format "HH:mm")</td>
<td style="font-size:11px;color:#7f8c8d;text-align:right;">PRTG Network Monitor - DSI</td>
</tr></table>
</td></tr>

</table>
</td></tr>
</table>
</body>
</html>
"@

# === SAUVEGARDE ===
$repHtml = Split-Path $cheminHtmlRapport -Parent
if (!(Test-Path $repHtml)) { New-Item -ItemType Directory -Path $repHtml -Force | Out-Null }

try {
    $utf8BOM = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllText($cheminHtmlRapport, $html, $utf8BOM)
    Write-Host "Rapport HTML : $cheminHtmlRapport" -ForegroundColor Green
} catch {
    Write-Error "Erreur generation HTML : $_"
    exit 1
}

# === CONSOLE ===
Write-Host "`n=== RESUME ===" -ForegroundColor Cyan
Write-Host "Total alertes       : $totalAlertes"
Write-Host "Equipements touches : $nbDevices"
Write-Host "Jours sans alerte   : $joursSansAlerte / $nbJoursMois"
Write-Host "Critique: $($sevCounts.Critique) | Haute: $($sevCounts.Haute) | Moyenne: $($sevCounts.Moyenne) | Basse: $($sevCounts.Basse)"

Write-Host "`n=== TOP 5 ===" -ForegroundColor Cyan
$topDevices | ForEach-Object { Write-Host "$($_.Total) alertes - $($_.Nom)" -ForegroundColor Yellow }

# === ENVOI MAIL ===
$destinataires = @("destinataire@mail.fr")
Send-MailMessage -From "from@mail.fr" `
    -To $destinataires `
    -Subject "[PRTG] Rapport exploitation - $moisAffichage" `
    -BodyAsHtml $html `
    -SmtpServer "stmps.entreprise.fr" `
    -Port xx `
    -UseSsl

Write-Host "Rapport envoye." -ForegroundColor Green
