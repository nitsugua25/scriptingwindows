# Define the OU Distinguished Name for "informatique"
$informatiqueOU = "OU=informatique,DC=forest,DC=com"

# Get all users in the domain
$allUsers = Get-ADUser -Filter * -Properties DistinguishedName | Select-Object SamAccountName, DistinguishedName

foreach ($user in $allUsers) {
    # Check if the user is in the "informatique" OU
    $isInInformatique = $user.DistinguishedName -like "*$informatiqueOU*"

    Write-Host "Checking user: $($user.SamAccountName)" -ForegroundColor Cyan

    try {
        # Define registry paths
        $cmdPath = "HKCU:\Software\Policies\Microsoft\Windows\System"
        $controlPanelPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"

        # Check and apply settings based on OU
        if ($isInInformatique) {
            # Ensure CMD and Control Panel are enabled for "informatique" users
            if (Test-Path $cmdPath) {
                Remove-ItemProperty -Path $cmdPath -Name "DisableCMD" -ErrorAction SilentlyContinue
            }
            if (Test-Path $controlPanelPath) {
                Remove-ItemProperty -Path $controlPanelPath -Name "NoControlPanel" -ErrorAction SilentlyContinue
            }
            Write-Host "User $($user.SamAccountName) (Informatique): CMD and Control Panel are ENABLED." -ForegroundColor Green
        } else {
            # Disable CMD for non-"informatique" users
            if (-not (Test-Path $cmdPath)) {
                New-Item -Path $cmdPath -Force | Out-Null
            }
            Set-ItemProperty -Path $cmdPath -Name "DisableCMD" -Value 1 -Type DWord -Force

            # Disable Control Panel for non-"informatique" users
            if (-not (Test-Path $controlPanelPath)) {
                New-Item -Path $controlPanelPath -Force | Out-Null
            }
            Set-ItemProperty -Path $controlPanelPath -Name "NoControlPanel" -Value 1 -Type DWord -Force

            Write-Host "User $($user.SamAccountName): CMD and Control Panel are DISABLED." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "An error occurred while processing $($user.SamAccountName): $_" -ForegroundColor Red
    }
}

Write-Host "Script execution completed." -ForegroundColor Yellow
