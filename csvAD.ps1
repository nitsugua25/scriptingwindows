# Import du module AD
Import-Module ActiveDirectory

# Fonction pour obtenir le domaine basé sur le département
function Get-DomainFromDepartment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Departement
    )
    
    try {
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
                Write-Warning "Département non reconnu: $Departement"
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
    # Import du fichier CSV
    Write-Host "Lecture du fichier CSV..."
    $Users = Import-Csv -Path ".\output.csv" -Delimiter ";" -Encoding UTF8
    
    foreach ($User in $Users) {
        Write-Host "Traitement de l'utilisateur : $($User.Prenom) $($User.Nom)"
        
        try {
            # Obtenir le chemin OU
            $ouPath = Get-DomainFromDepartment -Departement $User.Departement
            
            # Créer les paramètres pour le nouvel utilisateur
            $userParams = @{
                Name = "$($User.Prenom) $($User.Nom)"
                GivenName = $User.Prenom
                Surname = $User.Nom
                SamAccountName = $User.Login
                UserPrincipalName = "$($User.Login)@belgique.lan"
                Path = $ouPath
                AccountPassword = (ConvertTo-SecureString "P@ssw0rd" -AsPlainText -Force)
                Enabled = $true
                ChangePasswordAtLogon = $true
                Department = $User.Departement
                Description = $User.Description
                Office = $User.Bureau
                EmployeeID = $User.Numero_Interne
            }

            # Créer l'utilisateur
            New-ADUser @userParams
            Write-Host "Utilisateur créé avec succès : $($User.Login)"
        }
        catch {
            Write-Error "Erreur lors de la création de l'utilisateur $($User.Login) : $_"
            continue
        }
    }
    
    Write-Host "Traitement terminé avec succès."
}
catch {
    Write-Error "Erreur générale : $_"
    exit 1
}