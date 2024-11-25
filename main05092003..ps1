Import-Module ImportExcel
$Path = "./Employes.csv"
$outputFile = "./Doublons.xlsx"

$Users = $Path

foreach ($User in $Users) {
    $User = $User.SamAccountName

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