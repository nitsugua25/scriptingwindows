import-module ActiveDirectory

$outputFile = "./Doublons.xlsx"

$Users = Import-Csv ./Employes.csv

foreach ($Employes in $Users) {
    $User = $Employes.Login

    if (Get-ADUser -Filter {SamAccountName -eq $User}) {
        Write-Host "Exporting data to csv"
        $Data | Export-Excel -Path $outputFile -WorksheetName 'Doublons' -NoClobber:$false
    }
    else {
        userProps = @{
            Name = $User.Nom
            Surname = $User.Prenom
            EmailAddress = $User.Login
            Description = $User.description
            Departement = $User.Departement
            Office = $User.Bureau
            OfficeNumber = $User.Numero_interne
            AccountPassword = $User.Password
        }
        New-ADUser @userProps
        Write-Host "User $User created"
    }
}