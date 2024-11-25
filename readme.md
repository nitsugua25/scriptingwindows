## Password generator

*Password Generator is a fonction that return a random password*:

```sh
New-RandomPassword -Length $LenghtValue -Uppercase $UppercaseValue -Digits $DigitsValue -SpecialCharacters $SpecialValue
```

length is the only parameter that is required every other is optional and will use the value 0 by default.

```sh
    [Parameter(Mandatory=$true)][int]$Length,
    [Parameter(Mandatory=$false)][int]$Uppercase=0,
    [Parameter(Mandatory=$false)][int]$Digits=0,
    [Parameter(Mandatory=$false)][int]$SpecialCharacters=0
```
