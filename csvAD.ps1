$Users = Import-Csv ".\output.csv"
Import-Module ActiveDirectory

# Fonction pour obtenir le domaine basé sur le département
function Get-DomainFromDepartment {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Departement
    )
    
    # Journalisation du département reçu
    Write-Verbose "Traitement du département : $Departement"
    
    switch -Wildcard ($Departement) {
        "*Ressources humaines*" { return "rh.lan" }
        "*R&D*" { return "rd.lan" }
        "*Marketting*" { return "marketing.lan" }
        "*Marketing*" { return "marketing.lan" }
        "*Finances*" { return "finance.lan" }
        "*Technique*" { return "technique.lan" }
        "*Commerciaux*" { return "commercial.lan" }
        "*Informatique*" { return "it.lan" }
        "Direction" { return "direction.lan" }
        default { 
            Write-Warning "Département non reconnu: $Departement"
            return "belgique.lan" 
        }
    }
}

# Pour tester la fonction
try {
    $currentPath = Get-Location
    Write-Host "Chemin actuel : $currentPath"
    
    # Exemple d'utilisation
    $domain = Get-DomainFromDepartment -Departement "R&D"
    Write-Host "Domaine trouvé : $domain"
}
catch {
    Write-Error "Erreur : $_"
}
# Fonction pour obtenir un UPN valide
function Get-ValidUPN {
    param(
        [string]$FirstName,
        [string]$LastName
    )
    
    # D'abord essayer avec le prénom complet
    $baseUPN = "$FirstName.$LastName".ToLower()
    
    # Si plus long que 20 caractères, réduire au minimum (première lettre du prénom)
    if ($baseUPN.Length -gt 20) {
        Write-Warning "UPN trop long pour $FirstName $LastName : $baseUPN"
        $baseUPN = "$($FirstName.Substring(0,1)).$($LastName)".ToLower()
        
        # Si toujours trop long, tronquer à 20 caractères
        if ($baseUPN.Length -gt 20) {
            $baseUPN = $baseUPN.Substring(0, 20)
        }
    }
    
    return $baseUPN
}

# Initialisation des variables
$usedUPNs = @{}
$duplicates = @()

# Première passe : vérifier tous les UPN potentiels
foreach ($User in $Users) {
    $FirstName = $User.Prenom
    $LastName = $User.Nom
    $baseUPN = Get-ValidUPN -FirstName $FirstName -LastName $LastName
    $domain = Get-DomainFromDepartment -Departement $User.Departement
    
    if ($usedUPNs.ContainsKey($baseUPN)) {
        $usedUPNs[$baseUPN]++
        $duplicates += "Doublon trouvé: $FirstName $LastName -> $baseUPN$($usedUPNs[$baseUPN])@$domain"
    } else {
        $usedUPNs[$baseUPN] = 1
    }
}

# Écrire les doublons dans un fichier
if ($duplicates.Count -gt 0) {
    $duplicates | Out-File -FilePath ".\doublons_upn.txt"
    Write-Host "Des doublons ont été trouvés et enregistrés dans doublons_upn.txt"
}

# Réinitialiser le compteur
$usedUPNs.Clear()

# Deuxième passe : création des utilisateurs
foreach ($User in $Users) {
    $Displayname = "$($User.Prenom) $($User.Nom)"
    $UserFirstname = $User.Prenom
    $UserLastname = $User.Nom
    $OU = Get-OUPath -Departement $User.Departement
    
    # Générer l'UPN avec le bon domaine
    $baseUPN = Get-ValidUPN -FirstName $UserFirstname -LastName $UserLastname
    $domain = Get-DomainFromDepartment -Departement $User.Departement
    
    # Gérer les doublons d'UPN
    if ($usedUPNs.ContainsKey($baseUPN)) {
        $counter = $usedUPNs[$baseUPN] + 1
        $usedUPNs[$baseUPN] = $counter
        $UPN = "$baseUPN$counter@$domain"
        $SAM = "$baseUPN$counter"
    } else {
        $usedUPNs[$baseUPN] = 1
        $UPN = "$baseUPN@$domain"
        $SAM = $baseUPN
    }
    
    $Description = $User.Description
    $defaultPassword = "Changeme@123"

    # Création de l'utilisateur
    New-ADUser -Name "$Displayname" `
               -DisplayName "$Displayname" `
               -SamAccountName $SAM `
               -UserPrincipalName $UPN `
               -GivenName "$UserFirstname" `
               -Surname "$UserLastname" `
               -Description "$Description" `
               -AccountPassword (ConvertTo-SecureString $defaultPassword -AsPlainText -Force) `
               -Enabled $true `
               -Path "$OU" `
               -ChangePasswordAtLogon $true `
               -PasswordNeverExpires $false `
               -server $domain
    
    Write-Host "Créé utilisateur: $Displayname dans $OU avec UPN: $UPN"
}