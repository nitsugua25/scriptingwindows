# Import du module AD
Import-Module ActiveDirectory
function New-ADGG{
    param(
        [Parameter(Mandatory=$true)]
        [string]$GGName,
        [Parameter(Mandatory=$true)]
        [string]$BaseDN
    )

    try {
        $GG = Get-ADGroup -Filter { Name -eq $GGName } -SearchBase $BaseDN
        if($null -eq $GG) {
            Write-Host "Création du groupe global $GGName"
            New-ADGroup "GG-$GGName" -GroupCategory Security -GroupScope Global -Path "$BaseDN"
        }
        else {
            Write-Host "Le groupe global $GGName existe déja."
        }

    } catch {
        Write-Error "Erreur lors de la création du groupe global $GGName : $_"
    }
}
function New-ADOU {
    param(
        [Parameter(Mandatory=$true)]
        [string]$OUName,
        [Parameter(Mandatory=$true)]
        [string]$BaseDN
    )

    try {
        $OU = Get-ADOrganizationalUnit -Filter { Name -eq $OUName } -SearchBase $BaseDN
        if($null -eq $OU) {
            Write-Host "Création de l'OU $OUName"
            New-ADOrganizationalUnit -Name "$OUName" -Path "$BaseDN"
        }
        else {
            Write-Host "L'UO $OUName existe déja."
        }

    } catch {
        Write-Error "Erreur lors de la création de l'OU $OUName : $_"
    }
}

function New-RandomPassword {
    param (
        [Parameter(Mandatory=$true)][int]$Length,
        [Parameter(Mandatory=$false)][int]$Uppercase=0,
        [Parameter(Mandatory=$false)][int]$Digits=0,
        [Parameter(Mandatory=$false)][int]$SpecialCharacters=0
    )
    Begin {
        $Lowercase = $Length - $SpecialCharacters - $Uppercase - $Digits
        $ArrayLowerCharacters = @('a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z')
        $ArrayUpperCharacters = @('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z')
        $ArraySpecialCharacters = @('_','*','%','#','?','!','-')
        $ArrayDigits = @('0','1','2','3','4','5','6','7','8','9')
        $Password = @()
    }
    Process {
        for ($i = 0; $i -lt $Uppercase; $i++) {
            $Password += $ArrayUpperCharacters | Get-Random
        }
        for ($i = 0; $i -lt $Digits; $i++) {
            $Password += $ArrayDigits | Get-Random
        }
        for ($i = 0; $i -lt $SpecialCharacters; $i++) {
            $Password += $ArraySpecialCharacters | Get-Random
        }
        for ($i = 0; $i -lt $Lowercase; $i++) {
            $Password += $ArrayLowerCharacters | Get-Random
        }
        $Password = $Password | Sort-Object {Get-Random}
        $Password -join ''
    }
}

# Programme principal
try {
    New-ADOU -OUName "Groupes" -BaseDN "DC=astral,DC=lan"
    New-ADOU -OUName "Groupes Globaux" -BaseDN "OU=Groupes,DC=astral,DC=lan"
    New-ADOU -OUName "Groupes Locaux" -BaseDN "OU=Groupes,DC=astral,DC=lan"


    Write-Host "Lecture du fichier CSV..."
    $Users = Import-Csv -Path ".\output.csv" -Encoding UTF8
    $DepList = @()
    $AvailableUPN = @()
    $userList = @()
    $Blacklist = @()

    foreach ($User in $Users) {
        if ("$($User.Nom) $($User.Prenom)" -notin $userList) {
            $userList += "$($User.Nom) $($User.Prenom)"
        } else {
            $Blacklist += "$($User.Nom) $($User.Prenom)"
            Write-Host "Doublon détecté pour $($User.Prenom) $($User.Nom)"
        }
    }

    foreach ($User in $Users) {
        $parsedDN = ""
        $UserUPNSuffix = ""
        $GGName = ""

        Write-Host "`nTraitement de l'utilisateur : $($User.Prenom) $($User.Nom)"
        Write-Host "Département : $($User.Departement)"
        if (![string]::IsNullOrWhiteSpace($User.Departement)) {
            $Departement = $User.Departement
            $Departement = $Departement.Trim()
            $Departement = $Departement.ToLower()
            $Departement = $Departement -split "/"
            if ($Departement.Count -gt 1) {
                if ($DepList -notcontains $Departement[1]) {
                    $DepList += $Departement[1]
                    Write-Host ("Création de l'OU " + $Departement[1])
                    New-ADOU -OUName $Departement[1] -BaseDN "DC=astral,DC=lan"
                    New-ADGG -GGName $Departement[1] -BaseDN "OU=Groupes Globaux,OU=Groupes,DC=astral,DC=lan"
                }

                if ($DepList -notcontains $Departement[0]) {
                    $DepList += $Departement[0]
                    $baseDN = "OU=" + $Departement[1] + ",DC=astral,DC=lan".Trim()
                    New-ADOU -OUName $Departement[0].Trim() -BaseDN $BaseDN
                }
                $ParsedDN = ("OU=" + $Departement[0] + ",OU=" + $Departement[1] + ",DC=astral,DC=lan").Trim()
            } else {
                if ($DepList -notcontains $Departement[0]) {
                    $DepList += $Departement[0]
                    New-ADOU -OUName $Departement[0] -BaseDN "DC=astral,DC=lan"
                    New-ADGG -GGName $Departement[0] -BaseDN "OU=Groupes Globaux,OU=Groupes,DC=astral,DC=lan"
                }
                $parsedDN = ("OU=" + $Departement[0] + ",DC=astral,DC=lan").Trim()
            }

            }switch -Wildcard ($User.Departement) {
            "*Ressources humaines*" { $GGName = "Ressources-Humaines"; $UserUPNSuffix = "rh.lan" }
            "*R&D*" { $GGName = "R&D" ; $UserUPNSuffix = "r&d.lan" }
            "*Marketing*" { $GGName = "Marketing"; $UserUPNSuffix = "marketing.lan" }
            "*Finances*" { $GGName = "Finances" ; $UserUPNSuffix = "finance.lan" }
            "*Technique*" { $GGName = "Technique" ; $UserUPNSuffix = "technique.lan" }
            "*Commerciaux*" { $GGName = "Commerciaux" ; $UserUPNSuffix = "commercial.lan" }
            "*Informatique*" { $GGName = "Informatique" ; $UserUPNSuffix = "it.lan" }
            "Direction" { $GGName = "Direction" ; $UserUPNSuffix = "direction.lan" }
            default { 
                Write-Warning "Département non reconnu: '$Departement'"
                $UserUPNSuffix = "belgique.lan"
            }
        }

        if ($AvailableUPN -notcontains $UserUPNSuffix) {
            try {
                Get-ADForest | Set-ADForest -UPNSuffixes @{add=$UserUPNSuffix}
                $AvailableUPN += $UserUPNSuffix
                Write-Host "UPN '$UserUPNSuffix' ajouté à la forêt"
            }
            catch {
                Write-Error "Erreur lors de la création de l'UPN : $_"
                continue
            }   
        }
        } else {
            Write-Warning "Département manquant pour $($User.Prenom) $($User.Nom)"
            continue
        }

        $firstname = $User.Prenom
        $lastname = $User.Nom
        $baseUPN = $lastname.ToLower() + "." + $firstname.ToLower()
        if ($baseUPN.Length -gt 20) {
            Write-Warning "UPN trop long pour $FirstName $LastName : $baseUPN"
            $baseUPN = "$($FirstName.Substring(0,1)).$($LastName)".ToLower()
            
            # Si toujours trop long, tronquer à 20 caractères
            if ($baseUPN.Length -gt 20) {
                $baseUPN = $baseUPN.Substring(0, 20)
            }
        }
        if($UserUPNSuffix -eq "direction.lan") {
            $password = New-RandomPassword -Length 16 -Uppercase 2 -Digits 2 -SpecialCharacters 2

        } else {
            $password = New-RandomPassword -Length 7 -Uppercase 1 -Digits 2 -SpecialCharacters 1
        }
        if($Blacklist -contains "$($User.Nom) $($User.Prenom)") {
            Write-Warning "Doublon détecté pour $($User.Prenom) $($User.Nom)"
            Add-Content -Path ".\duplicate.txt" -Value "Nom Prénom: $($User.Nom) $($User.Prenom)\nDN: $parsedDN\nCouple Généré: $baseUPN@$UserUPNSuffix : $password"
        }
        try {
            New-ADUser -UserPrincipalName "$baseUPN@$UserUPNSuffix" -Name "$firstname $lastname" -GivenName $firstname -Surname $lastname -SamAccountName $baseUPN -DisplayName "$firstname $lastname" -Path $parsedDN -AccountPassword (ConvertTo-SecureString $password -AsPlainText -Force) -Enabled $true -OtherAttributes @{'ipPhone' = $User.NInterne} -Description $User.Description -Office $User.Bureau
            Write-Host "Utilisateur $($User.Prenom) $($User.Nom) créé" 
            Add-Content -Path ".\passwords.txt" -Value "$baseUPN@$UserUPNSuffix : $password"

            Add-AdGroupMember -Identity "GG-$GGName" -Members "$baseUPN"
        } catch {
            Write-Error "Erreur lors de la création de l'utilisateur $baseUPN : $_"
        }

} catch {
    Write-Error "Erreur générale : $_"
    exit 1
}