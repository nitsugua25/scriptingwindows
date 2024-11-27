Import-Module GroupPolicy

# Créer une nouvelle GPO
$gpoName = "Deploy Office"
$gpo = New-GPO -Name $gpoName

# Configurer le script de démarrage
$scriptPath = "\\server\share\install-office.bat"
$gpo | Set-GPRegistryValue -Key "HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\System" -ValueName "Run" -Type String -Value $scriptPath

# Lier la GPO au niveau du domaine
$domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
$domainPath = $domain.Name
$gpo | New-GPLink -Target $domainPath

