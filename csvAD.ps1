# Fonction pour obtenir le domaine basé sur le département
function Get-DomainFromDepartment {
    param($Departement)
    
    switch -Wildcard ($Departement) {
        "*Ressources humaines*" { return "rh.lan" }
        "*R&D*" { return "rd.lan" }
        "*Marketting*" { return "marketing.lan" }
        "*Finances*" { return "finance.lan" }
        "*Technique*" { return "technique.lan" }
        "*Commerciaux*" { return "commercial.lan" }
        "*Informatique*" { return "it.lan" }
        "Direction" { return "direction.lan" }
        default { return "belgique.lan" }
    }
}

# Première passe : vérifier tous les UPN potentiels
foreach ($User in $Users) {
    $FirstName = $User.Prenom
    $LastName = $User.Nom
    $baseUPN = "$($FirstName.Substring(0,1)).$($LastName)".ToLower()
    $domain = Get-DomainFromDepartment -Departement $User.Departement
    
    if ($usedUPNs.ContainsKey($baseUPN)) {
        $usedUPNs[$baseUPN]++
        $duplicates += "Doublon trouvé: $FirstName $LastName -> $baseUPN$($usedUPNs[$baseUPN])@$domain"
    } else {
        $usedUPNs[$baseUPN] = 1
    }
}

foreach ($User in $Users) {
    $Displayname = "$($User.Prenom) $($User.Nom)"
    $UserFirstname = $User.Prenom
    $UserLastname = $User.Nom
    $OU = Get-OUPath -Departement $User.Departement
    
    # Générer l'UPN avec le bon domaine
    $baseUPN = "$($UserFirstname.Substring(0,1)).$($UserLastname)".ToLower()
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