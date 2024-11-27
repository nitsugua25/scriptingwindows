# Importer les modules nécessaires
Import-Module ActiveDirectory
Import-Module GroupPolicy

# Définir les variables
$gpoName = "Mount Y and Z Drives"
$ouPath = "OU=Users,DC=domain,DC=local"
$scriptContent = @"
# Définir les chemins des partitions à monter
`$yDrivePath = "\\server\share\Y"
`$zDrivePath = "\\server\share\Z"

# Monter la partition Y
if (-not (Test-Path Y:)) {
    New-PSDrive -Name Y -PSProvider FileSystem -Root `$yDrivePath -Persist
}

# Monter la partition Z
if (-not (Test-Path Z:)) {
    New-PSDrive -Name Z -PSProvider FileSystem -Root `$zDrivePath -Persist
}
"@

# Créer une nouvelle GPO
$gpo = New-GPO -Name $gpoName

# Définir le chemin du script de connexion dans la GPO
$gpoScriptPath = "\\domain.local\SysVol\domain.local\Policies\{$($gpo.Id)}\User\Scripts\Logon"
New-Item -Path $gpoScriptPath -ItemType Directory -Force

# Créer le fichier de script de connexion
$scriptFilePath = "$gpoScriptPath\MountDrives.ps1"
Set-Content -Path $scriptFilePath -Value $scriptContent

# Configurer la GPO pour exécuter le script à la connexion
$gpoScriptsPath = "$gpoScriptPath\Logon.ini"
$gpoScriptsContent = @"
[Scripts]
0CmdLine=powershell.exe -ExecutionPolicy Bypass -File "$scriptFilePath"
0Parameters=
"@
Set-Content -Path $gpoScriptsPath -Value $gpoScriptsContent

# Lier la GPO à une OU
New-GPLink -Name $gpoName -Target $ouPath

# Vérifier la création et la liaison de la GPO
Get-GPO -Name $gpoName
Get-GPLink -Target $ouPath
