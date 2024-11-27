# Définir le nom de la GPO
$gpoName = "Customize Start Menu"

# Créer une nouvelle GPO
$gpo = New-GPO -Name $gpoName

# Définir le chemin du fichier XML de configuration du menu Démarrer
$startMenuXmlPath = "C:\Path\To\Your\StartMenuLayout.xml"

# Créer le répertoire pour le fichier XML dans la GPO
$gpoPath = "\\domain.local\SysVol\domain.local\Policies\{$($gpo.Id)}\User\Preferences\StartMenu"
New-Item -Path $gpoPath -ItemType Directory -Force

# Copier le fichier XML dans le répertoire de la GPO
Copy-Item -Path $startMenuXmlPath -Destination "$gpoPath\StartMenuLayout.xml"

# Configurer la GPO pour utiliser le fichier XML
$gpoXmlPath = "$gpoPath\Preferences.xml"
$gpoXmlContent = @"
<Preferences>
  <StartMenuLayout>
    <LayoutFile>$gpoPath\StartMenuLayout.xml</LayoutFile>
  </StartMenuLayout>
</Preferences>
"@
Set-Content -Path $gpoXmlPath -Value $gpoXmlContent

# Lier la GPO à une OU
$ouPath = "OU=Users,DC=domain,DC=local"
New-GPLink -Name $gpoName -Target $ouPath
