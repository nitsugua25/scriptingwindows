# Importer le module Active Directory
Import-Module ActiveDirectory

# Définir les paramètres de la stratégie de mot de passe
$passwordPolicyName = "DirectionPasswordPolicy"
$passwordPolicyDisplayName = "Direction Password Policy"
$passwordPolicyDescription = "Password policy for members of the direction group"
$minPasswordLength = 15
$complexityEnabled = $true
$passwordHistoryCount = 24
$maxPasswordAge = 90
$minPasswordAge = 1
$lockoutDuration = 30
$lockoutObservationWindow = 30
$lockoutThreshold = 5

# Créer la stratégie de mot de passe granulaire
New-ADFineGrainedPasswordPolicy -Name $passwordPolicyName -DisplayName $passwordPolicyDisplayName -Description $passwordPolicyDescription -MinPasswordLength $minPasswordLength -ComplexityEnabled $complexityEnabled -PasswordHistoryCount $passwordHistoryCount -MaxPasswordAge $maxPasswordAge -MinPasswordAge $minPasswordAge -LockoutDuration $lockoutDuration -LockoutObservationWindow $lockoutObservationWindow -LockoutThreshold $lockoutThreshold

# Appliquer la stratégie au groupe de la direction
$directionGroup = "GG-direction"
Add-ADFineGrainedPasswordPolicySubject -Identity $directionGroup -Subjects $passwordPolicyName

# Vérifier l'application de la stratégie
Get-ADFineGrainedPasswordPolicySubject -Identity $directionGroup
