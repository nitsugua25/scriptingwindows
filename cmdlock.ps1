# Import the Group Policy module
Import-Module GroupPolicy

# Variables - Replace these with your actual environment details
$domain = 'forest.com'  # Your domain name
$ouInformatique = 'OU=informatique,DC=forest,DC=com'  # OU where CMD and Control Panel will be enabled

# Define GPO names
$gpoName_Disable = 'Disable CMD and Control Panel Globally'
$gpoName_Enable = 'Enable CMD and Control Panel for Informatique'

# Step 1: Create or Retrieve the Global Disable GPO
$gpoDisable = Get-GPO -Name $gpoName_Disable -ErrorAction SilentlyContinue
if (-not $gpoDisable) {
    $gpoDisable = New-GPO -Name $gpoName_Disable -Domain $domain
    Write-Host "Created GPO: $gpoName_Disable"
} else {
    Write-Host "GPO already exists: $gpoName_Disable"
}

# Configure settings in the GPO to disable CMD and Control Panel
Set-GPRegistryValue -Name $gpoName_Disable -Key 'HKCU\Software\Policies\Microsoft\Windows\System' `
    -ValueName 'DisableCMD' -Type DWord -Value 1
Set-GPRegistryValue -Name $gpoName_Disable -Key 'HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer' `
    -ValueName 'NoControlPanel' -Type DWord -Value 1

# Link the Global Disable GPO to the domain root
New-GPLink -Name $gpoName_Disable -Target "DC=forest,DC=com" -LinkEnabled Yes
Write-Host "Linked GPO '$gpoName_Disable' to the domain root"

# Step 2: Create or Retrieve the Enable GPO for Informatique
$gpoEnable = Get-GPO -Name $gpoName_Enable -ErrorAction SilentlyContinue
if (-not $gpoEnable) {
    $gpoEnable = New-GPO -Name $gpoName_Enable -Domain $domain
    Write-Host "Created GPO: $gpoName_Enable"
} else {
    Write-Host "GPO already exists: $gpoName_Enable"
}

# Configure settings in the GPO to enable CMD and Control Panel
Set-GPRegistryValue -Name $gpoName_Enable -Key 'HKCU\Software\Policies\Microsoft\Windows\System' `
    -ValueName 'DisableCMD' -Type DWord -Value 0
Set-GPRegistryValue -Name $gpoName_Enable -Key 'HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer' `
    -ValueName 'NoControlPanel' -Type DWord -Value 0

# Link the Enable GPO to the "informatique" OU
New-GPLink -Name $gpoName_Enable -Target $ouInformatique -LinkEnabled Yes
Write-Host "Linked GPO '$gpoName_Enable' to OU '$ouInformatique'"

Write-Host "GPO configuration completed successfully."
