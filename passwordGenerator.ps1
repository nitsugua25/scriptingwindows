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
        $ArraySpecialCharacters = @('_','*','%','\#','?','!','-')
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

New-RandomPassword -Length 7 -Uppercase 1 -Digits 2 -SpecialCharacters 2