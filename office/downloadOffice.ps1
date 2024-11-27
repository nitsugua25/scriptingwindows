# Chemin vers le fichier d'installation d'office
$odtPath = "C:\Path\To\ODT"
# Chemin vers le fichier de configuration XML 
$configPath = ".\configOffice.xml"

Start-Process -FilePath "$odtPath\setup.exe" -ArgumentList "/download $configPath" -Wait
