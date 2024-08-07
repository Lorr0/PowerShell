# Declaration des variables pour le serveur DNS et le domaine
$dnsServer = "SRV-DNS-AD"  # Remplacez par le nom de votre serveur DNS
$domain = "AD-DNS"  # Remplacez par le domaine que vous souhaitez interroger

# Liste des prefixes d'adresses IP a verifier
$ipPrefixes = @("172.168.2", "192.168.30") #Exemple d'ip

# Fonction pour extraire et afficher les enregistrements DNS
function Extract-DNSRecords {
    param (
        [string]$dnsServer,
        [string]$domain,
        [array]$ipPrefixes,
        [string]$outputFolder  # Nouveau paramètre pour le repertoire de sortie
    )

    # Assurez-vous que le repertoire de sortie existe
    if (-not (Test-Path $outputFolder)) {
        Write-Host "Le repertoire specifie n'existe pas. Creation du repertoire..."
        New-Item -Path $outputFolder -ItemType Directory | Out-Null
    }

    # Definir le nom du fichier de sortie avec la date actuelle
    $currentDate = Get-Date
    $dateString = $currentDate.ToString("yyyyMMdd")
    $outputFile = Join-Path -Path $outputFolder -ChildPath "DNSRecordsOutput_$dateString.txt"

    # Initialiser un dictionnaire pour stocker les enregistrements regroupes par prefixe
    $dnsRecordsGrouped = @{}

    # Initialiser les groupes dans le dictionnaire
    foreach ($prefix in $ipPrefixes) {
        $dnsRecordsGrouped[$prefix] = @()
    }

    try {
        # Recuperer tous les enregistrements DNS
        $allRecords = Get-DnsServerResourceRecord -ComputerName $dnsServer -ZoneName $domain -ErrorAction Stop
        # Initialiser la sortie
        $output = ""
    } catch {
        # En cas d'erreur, ecrire un message d'erreur dans le fichier de sortie
        $output = "Erreur lors de l'obtention des enregistrements DNS : $_`r`n"
        $output | Out-File -FilePath $outputFile -Encoding UTF8 -Append
        return
    }

    # Collecter tous les enregistrements DNS
    $allDnsRecords = @()
    
    foreach ($record in $allRecords) {
        if ($record.RecordType -eq "A") {
            $ip = $record.RecordData.IPv4Address.ToString()
            foreach ($prefix in $ipPrefixes) {
                if ($ip.StartsWith($prefix)) {
                    $allDnsRecords += [PSCustomObject]@{
                        Name = $record.HostName
                        IP = $ip
                    }
                    break
                }
            }
        }
    }

    # Supprimer les doublons bases sur l'adresse IP
    $uniqueDnsRecords = $allDnsRecords | Sort-Object -Property IP -Unique

    # Regrouper les enregistrements par prefixe
    foreach ($prefix in $ipPrefixes) {
        $dnsRecordsGrouped[$prefix] = $uniqueDnsRecords | Where-Object { $_.IP.StartsWith($prefix) }
    }

    # Afficher les enregistrements DNS groupes par prefixe
    foreach ($prefix in $ipPrefixes) {
        if ($dnsRecordsGrouped[$prefix].Count -gt 0) {
            $output += "Enregistrements DNS pour le prefixe $prefix :`r`n"
            $output += ($dnsRecordsGrouped[$prefix] | Format-Table -AutoSize | Out-String) + "`r`n"
        } else {
            $output += "Aucun enregistrement DNS trouve pour le prefixe $prefix.`r`n"
        }
    }

    # ecrire la sortie dans le fichier
    $output | Out-File -FilePath $outputFile -Encoding UTF8
}

# Specifiez le repertoire de sortie pour les fichiers de sortie
$outputFolder = "C:\Chemin_de_sortie"  # Remplacez par le chemin desire

# Extraire les enregistrements DNS actuels
Extract-DNSRecords -dnsServer $dnsServer -domain $domain -ipPrefixes $ipPrefixes -outputFolder $outputFolder

