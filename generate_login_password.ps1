Import-Module ImportExcel
$inputFile = "C:\Users\Utilisateur\Downloads\Employes.xlsx"
$outputFile = "C:\Users\Utilisateur\Downloads\Employes.xlsx"

# on check tout le fichier et pas seulement la liste qu'il faut : à changer
# on ne limite pas longeur login : à changer 
# on ne verifie pas les doublons : à changer    

Write-Host "Starting script execution..."

# Function to remove diacritics from a string
Write-Host "Defining function Remove-Diacritics..."
function Remove-Diacritics {
    param (
        [string]$text
    )
    if (-not $text) { return $null }
    $normalized = $text.Normalize([System.Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($c in $normalized.ToCharArray()) {
        $unicodeCategory = [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($c)
        if ($unicodeCategory -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            $sb.Append($c) | Out-Null
        }
    }
    return $sb.ToString().Normalize([System.Text.NormalizationForm]::FormC)
}

# Function to remove non-alphanumeric characters
Write-Host "Defining function Remove-NonAlphanumeric..."
function Remove-NonAlphanumeric {
    param (
        [string]$text
    )
    if (-not $text) { return $null }
    return -join ($text -split '[^a-zA-Z0-9]')
}

# New function to generate a random password
Write-Host "Defining function New-RandomPassword..."
function New-RandomPassword {
    param (
        [Parameter(Mandatory=$true)][int]$Length,
        [Parameter(Mandatory=$false)][int]$Uppercase=1,
        [Parameter(Mandatory=$false)][int]$Digits=1,
        [Parameter(Mandatory=$false)][int]$SpecialCharacters=1
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

Write-Host "Importing Excel file..."
$Data = Import-Excel -Path $inputFile

# Display available properties
Write-Host "Available properties: $($Data[0].PSObject.Properties.Name -join ', ')"

Write-Host "Processing data..."
foreach ($row in $Data) {
    Write-Host "Processing row for employee: $($row.Nom) $($row.Prénom)"
    # Remove diacritics
    $cleanNom = Remove-Diacritics -text $row.Nom
    $cleanPrenom = Remove-Diacritics -text $row.Prénom

    # Remove non-alphanumeric characters
    $cleanNom = Remove-NonAlphanumeric -text $cleanNom
    $cleanPrenom = Remove-NonAlphanumeric -text $cleanPrenom

    # Generate login
    $login = "$cleanNom.$cleanPrenom".ToLower()
    # Add 'Login' property
    $row | Add-Member -NotePropertyName 'Login' -NotePropertyValue $login

    # Generate password
    $Password = New-RandomPassword -Length 7 -Uppercase 1 -Digits 2 -SpecialCharacters 2
    # Add 'Password' property
    $row | Add-Member -NotePropertyName 'Password' -NotePropertyValue $Password
}

Write-Host "Exporting data to Excel..."
$Data | Export-Excel -Path $outputFile -WorksheetName 'Employes' -NoClobber:$false

Write-Host "Script execution completed."