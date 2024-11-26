# Import du module AD
Import-Module ActiveDirectory

# Fonction pour obtenir le chemin OU basé sur le département
function Get-DomainFromDepartment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Departement
    )
    
    try {
        # Vérification et nettoyage du département
        if ([string]::IsNullOrWhiteSpace($Departement)) {
            Write-Warning "Département vide ou null, utilisation de l'OU par défaut"
            return "OU=Users,DC=belgique,DC=lan"
        }

        $Departement = $Departement.Trim()
        Write-Verbose "Traitement du département : '$Departement'"
        
        # Base DN pour les OUs
        $baseDN = "DC=belgique,DC=lan"
        
        switch -Wildcard ($Departement) {
            "*Ressources humaines*" { return "OU=RH,$baseDN" }
            "*R&D*" { return "OU=RD,$baseDN" }
            "*Marketting*" { return "OU=Marketing,$baseDN" }
            "*Marketing*" { return "OU=Marketing,$baseDN" }
            "*Finances*" { return "OU=Finance,$baseDN" }
            "*Technique*" { return "OU=Technique,$baseDN" }
            "*Commerciaux*" { return "OU=Commercial,$baseDN" }
            "*Informatique*" { return "OU=IT,$baseDN" }
            "Direction" { return "OU=Direction,$baseDN" }
            default { 
                Write-Warning "Département non reconnu: '$Departement'"
                return "OU=Users,$baseDN" 
            }
        }
    }
    catch {
        Write-Error "Erreur lors de la détermination du chemin OU : $_"
        return "OU=Users,$baseDN"
    }
}

# Programme principal
try {
    Write-Host "Lecture du fichier CSV..."
    $Users = Import-Csv -Path ".\output.csv" -Delimiter ";" -Encoding UTF8
    
    foreach ($User in $Users) {
        Write-Host "`nTraitement de l'utilisateur : $($User.Prenom) $($User.Nom)"
        
        if ([string]::IsNullOrWhiteSpace($User.Departement)) {
            Write-Warning "Département manquant pour $($User.Prenom) $($User.Nom)"
            continue
        }

        try {
            Write-Host "Département trouvé : '$($User.Departement)'"
            $ouPath = Get-DomainFromDepartment -Departement $User.Departement -Verbose
            Write-Host "OU Path : $ouPath"
            
            # Création de l'utilisateur dans AD
            # ... code de création d'utilisateur ...
        }
        catch {
            Write-Error "Erreur lors de la création de l'utilisateur $($User.Login) : $_"
            continue
        }
    }
}
catch {
    Write-Error "Erreur générale : $_"
    exit 1
}