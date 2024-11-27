#
#A CHANGER
$SharedFolderPath = "\\chemin\redirection"

# Nom de la GPO
$GpoName = "Redirect My Documents"

# Créer une nouvelle GPO
$Gpo = New-GPO -Name $GpoName

# Lier la GPO au niveau du domaine
$Domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
$DomainDN = $Domain.GetDirectoryEntry().distinguishedName
New-GPLink -Name $GpoName -Target $DomainDN

# Configurer la redirection des dossiers "Mes documents"
$GpoPath = "C:\Windows\SYSVOL\domain\Policies\{$($Gpo.Id)}\User\Preferences\Folders\Folder"

# Créer le fichier XML pour la redirection des dossiers
$XmlContent = @"
<Folder>
  <Name>Documents</Name>
  <Action>R</Action>
  <Path>$SharedFolderPath\%username%\Documents</Path>
  <Options>
    <GrantExclusiveRights>true</GrantExclusiveRights>
    <MoveContents>true</MoveContents>
    <PolicyRemoval>Redirect</PolicyRemoval>
  </Options>
</Folder>
"@

# Enregistrer le fichier XML dans le répertoire de la GPO
$XmlFilePath = "$GpoPath\Documents.xml"
New-Item -Path $GpoPath -ItemType Directory -Force
Set-Content -Path $XmlFilePath -Value $XmlContent

# Afficher un message de confirmation
Write-Output "La GPO '$GpoName' a été créée et configurée pour rediriger le dossier 'Mes documents' vers '$SharedFolderPath'."
