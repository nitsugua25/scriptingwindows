# Définir le nom de la GPO
$gpoName = "Set Wallpaper"

# Créer une nouvelle GPO
$gpo = New-GPO -Name $gpoName

# Définir le chemin du fond d'écran
$wallpaperPath = "C:\Path\To\Your\Wallpaper.jpg"

# Créer le répertoire pour le fond d'écran dans la GPO
$gpoPath = "\\domain.local\SysVol\domain.local\Policies\{$($gpo.Id)}\User\Preferences\Wallpaper"
New-Item -Path $gpoPath -ItemType Directory -Force

# Copier le fond d'écran dans le répertoire de la GPO
Copy-Item -Path $wallpaperPath -Destination "$gpoPath\Wallpaper.jpg"

# Créer le fichier XML pour les préférences de la GPO
$gpoXmlPath = "$gpoPath\Preferences.xml"
$gpoXmlContent = @"
<Preferences>
  <DesktopWallpaper>
    <Wallpaper>$gpoPath\Wallpaper.jpg</Wallpaper>
    <Style>Fill</Style>
  </DesktopWallpaper>
</Preferences>
"@
Set-Content -Path $gpoXmlPath -Value $gpoXmlContent

# Lier la GPO à une OU
$ouPath = "OU=Users,DC=domain,DC=local"
New-GPLink -Name $gpoName -Target $ouPath
