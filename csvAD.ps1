# Import du module AD
Import-Module ActiveDirectory

function New-ADGG {
    param (
        [Parameter(Mandatory=$true)]
        [string]$GName,
        [Parameter(Mandatory=$true)]
        [string]$BaseDN
    )
    try {
        $GN = Get-ADGroup -Filter { Name -eq $GName } -SearchBase $BaseDN
        if($null -eq $GN) {
            Write-Host "Création du groupe $GName"
            New-ADGroup -Name "$GName" -Path "$BaseDN" -GroupCategory Security -GroupScope Global
        }
        else {
            Write-Host "Le groupe $GName existe déja."
        }
    }
    catch {
        Write-Error "Erreur lors de la création du groupe $GName : $_"
    }
}

function New-ADGL {
    param (
        [Parameter(Mandatory=$true)]
        [string]$GName,
        [Parameter(Mandatory=$true)]
        [string]$BaseDN
    )
    try {
        $GN = Get-ADGroup -Filter { Name -eq $GName } -SearchBase $BaseDN
        if($null -eq $GN) {
            Write-Host "Création du groupe $GName"
            New-ADGroup -Name "$GName" -Path "$BaseDN" -GroupCategory Security -GroupScope Local
        }
        else {
            Write-Host "Le groupe $GName existe déja."
        }
    }
    catch {
        Write-Error "Erreur lors de la création du groupe $GName : $_"
    }
}

function New-GPO-Share {
    param (
        [Parameter(Mandatory=$true)]
        [string]$basePath
    )
    New-Item -Path "$basePath" -Name "gpo$" -ItemType "directory"
    New-Item -Path "$basePath" -Name "user-doccuments$" -ItemType "directory"
    New-SmbShare -Name "gpo$" -Path "$basePath\gpo$" -ReadAccess "BELGIQUE\Domain Users" -FullAccess "BELGIQUE\Domain Admins"
    New-SmbShare -Name "user-doccuments$" -Path "$basePath\user-doccuments$" -ChangeAccess "BELGIQUE\Domain Users" -FullAccess "BELGIQUE\Domain Admins"
}

function Set-AGDLP-and-Main-Share {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$repList,
        [Parameter(Mandatory=$true)]
        [string]$BaseDN,
        [Parameter(Mandatory=$true)]
        [string]$basePath
    )
    $permTable = @{
        "GL-direction-rw" = @("GG-direction")
        "GL-ressources humaine-r" = @()
        "GL-ressources humaine-rw" = @()
        "GL-gestion du personnel-r" = @()
        "GL-gestion du personnel-rw" = @()
        "GL-recrutement-r" = @()
        "GL-recrutement-rw" = @()
        "GL-r&d-r" = @()
        "GL-r&d-rw" = @()
        "GL-recherche-r" = @()
        "GL-recherche-rw" = @()
        "GL-testing-r" = @()
        "GL-testing-rw" = @()
        "GL-marketing-r" = @()
        "GL-marketing-rw" = @()
        "GL-site 1-r" = @()
        "GL-site 1-rw" = @()
        "GL-site 2-r" = @()
        "GL-site 2-rw" = @()
        "GL-site 3-r" = @()
        "GL-site 3-rw" = @()
        "GL-site 4-r" = @()
        "GL-site 4-rw" = @()
        "GL-informatique-r" = @()
        "GL-informatique-rw" = @()
        "GL-systemes-r" = @()
        "GL-systemes-rw" = @()
        "GL-developpement-r" = @()
        "GL-developpement-rw" = @()
        "GL-hotline-r" = @()
        "GL-hotline-rw" = @()
        "GL-finances-r" = @()
        "GL-finances-rw" = @()
        "GL-comptabilite-r" = @()
        "GL-comptabilite-rw" = @()
        "GL-investissements-r" = @()
        "GL-investissements-rw" = @()
        "GL-technique-r" = @()
        "GL-technique-rw" = @()
        "GL-techniciens-r" = @()
        "GL-techniciens-rw" = @()
        "GL-achat-r" = @()
        "GL-achat-rw" = @()
        "GL-commerciaux-r" = @()
        "GL-commerciaux-rw" = @()
        "GL-sedentaires-r" = @()
        "GL-sedentaires-rw" = @()
        "GL-technico-r" = @()
        "GL-technico-rw" = @()
    }

    New-Item -Path "$basePath" -Name "share$" -ItemType "directory"
    $basePath += "\share$\"
    New-SmbShare -Name "share$" -Path "$basePath" -ChangeAccess "BELGIQUE\Domain Users" -FullAccess "BELGIQUE\Domain Admins"

    foreach($key in $repList.Keys) {
        New-Item -Path "$basePath" -Name "$key" -ItemType "directory"
        foreach($depPath in $key){
            New-Item -Path "$basePath$key\" -Name "$depPath" -ItemType "directory"
        }
    }

    foreach($key in $permTable.Keys) {
        New-ADGL -GName "$key" -BaseDN "$baseDN"
        foreach($perm in $permTable["$key"]) {
            Add-ADGroupMember -Identity "$key" -Members "$perm"
        }
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

    $rootdc = "DC=belgique,DC=lan"

    New-ADOU -OUName "Groupes" -BaseDN "$rootdc"
    New-ADOU -OUName "Groupes Globaux" -BaseDN "OU=Groupes,$rootdc"
    New-ADOU -OUName "Groupes Locaux" -BaseDN "OU=Groupes,$rootdc"
    New-ADOU -OUName "ordinateurs" -BaseDN "$rootdc"

    redircmp "OU=ordinateurs,$rootdc" #Redirect Old Computers to New target

    Write-Host "Lecture du fichier CSV..."
    $Users = Import-Csv -Path ".\output.csv" -Encoding UTF8
    $DepList = @()
    $AvailableUPN = @()
    $userList = @()
    $Blacklist = @()
    $repList = @{}


    foreach ($User in $Users) {
        if ("$($User.Nom) $($User.Prenom)" -notin $userList) {
            $userList += "$($User.Nom) $($User.Prenom)"
        } else {
            $Blacklist += "$($User.Nom) $($User.Prenom)"
            Write-Host "Doublon détecté pour $($User.Prenom) $($User.Nom)"
        }
    }

    foreach ($User in $Users) {
        $depHead = $false
        $parsedDN = ""
        $UserUPNSuffix = ""

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
                    New-ADOU -OUName $Departement[1] -BaseDN "$rootdc"
                    New-ADGG -gname "GG-responsable-$($Departement[1])" -BaseDN "OU=Groupes Globaux,OU=Groupes,$rootdc"
                    New-ADGL -gname "GL-$($Departement[1])-r" -BaseDN "OU=Groupes Locaux,OU=Groupes,$rootdc"
                    New-ADGL -gname "GL-$($Departement[1])-rw" -BaseDN "OU=Groupes Locaux,OU=Groupes,$rootdc"
                    $repList.Add($Departement[1], @() )
                }

                if ($DepList -notcontains $Departement[0]) {
                    $DepList += $Departement[0]
                    $baseDN = "OU=" + $Departement[1] + ",$rootdc".Trim()
                    New-ADOU -OUName $Departement[0].Trim() -BaseDN $BaseDN
                    New-ADGG -gname "GG-$($Departement[0])" -BaseDN "OU=Groupes Globaux,OU=Groupes,$rootdc"
                    New-ADGL -gname "GL-$($Departement[0])-r" -BaseDN "OU=Groupes Locaux,OU=Groupes,$rootdc"
                    New-ADGL -gname "GL-$($Departement[0])-rw" -BaseDN "OU=Groupes Locaux,OU=Groupes,$rootdc"
                    $depHead = $true
                    $repList[$Departement[1]] += $Departement[0]
                }
                $ParsedDN = ("OU=" + $Departement[0] + ",OU=" + $Departement[1] + ",$rootdc").Trim()
            } else {
                if ($DepList -notcontains $Departement[0]) {
                    $DepList += $Departement[0]
                    New-ADOU -OUName $Departement[0] -BaseDN "$rootdc"
                    New-ADGG -gname "GG-$($Departement[0])" -BaseDN "OU=Groupes Globaux,OU=Groupes,$rootdc"
                    New-ADGL -gname "GL-$($Departement[0])-r" -BaseDN "OU=Groupes Locaux,OU=Groupes,$rootdc"
                    New-ADGL -gname "GL-$($Departement[0])-rw" -BaseDN "OU=Groupes Locaux,OU=Groupes,$rootdc"
                }
                $parsedDN = ("OU=" + $Departement[0] + ",$rootdc").Trim()
            }

            if($depHead -eq $false) {
                $GGName = "GG-$($Departement[0])"
            } else {
                $GGName = "GG-responsable-$($Departement[1])"
            }
            switch -Wildcard ($User.Departement) {
                "*Ressources humaines*" { 
                    $UserUPNSuffix = "rh.lan"
                }
                "*R&D*" { 
                    $UserUPNSuffix = "r&d.lan"
                }
                "*Marketing*" { 
                    $UserUPNSuffix = "marketing.lan" 
                }
                "*Finances*" { 
                    $UserUPNSuffix = "finance.lan" 
                }
                "*Technique*" { 
                    $UserUPNSuffix = "technique.lan" 
                }
                "*Commerciaux*" { 
                    $UserUPNSuffix = "commercial.lan" 
                }
                "*Informatique*" { 
                    $UserUPNSuffix = "it.lan" 
                }
                "Direction" { 
                    $GGName="GG-direction"
                    $UserUPNSuffix = "direction.lan" 
                }
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
            Add-ADGroupMember -Identity "$GGName" -Members $baseUPN
            Write-Host "Utilisateur $($User.Prenom) $($User.Nom) créé" 
            Add-Content -Path ".\passwords.txt" -Value "$baseUPN@$UserUPNSuffix : $password"
        } catch {
            Write-Error "Erreur lors de la création de l'utilisateur $baseUPN : $_"
        }
    }

} catch {
    Write-Error "Erreur générale : $_"
    exit 1
}