# Define the OU
$ouDN = "OU=Informatique,DC=forest,DC=com"

# Get all users not in the Informatique OU
$users = Get-ADUser -Filter * | Where-Object {
    $_.DistinguishedName -notlike "*$ouDN*"
}

foreach ($user in $users) {
    Write-Host "Processing user: $($user.SamAccountName)" -ForegroundColor Cyan

    try {
        # Ensure registry path for CMD exists
        $cmdPath = "HKCU:\Software\Policies\Microsoft\Windows\System"
        if (-not (Test-Path $cmdPath)) {
            New-Item -Path $cmdPath -Force | Out-Null
            Write-Host "Created registry path: $cmdPath" -ForegroundColor Yellow
        }

        # Disable CMD
        Set-ItemProperty -Path $cmdPath -Name "DisableCMD" -Value 1 -Type DWord -Force
        $cmdDisabled = Get-ItemProperty -Path $cmdPath -Name "DisableCMD"
        if ($cmdDisabled.DisableCMD -eq 1) {
            Write-Host "Successfully disabled CMD for $($user.SamAccountName)." -ForegroundColor Green
        } else {
            Write-Host "Failed to disable CMD for $($user.SamAccountName)." -ForegroundColor Red
        }

        # Ensure registry path for Control Panel exists
        $controlPanelPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
        if (-not (Test-Path $controlPanelPath)) {
            New-Item -Path $controlPanelPath -Force | Out-Null
            Write-Host "Created registry path: $controlPanelPath" -ForegroundColor Yellow
        }

        # Disable Control Panel
        Set-ItemProperty -Path $controlPanelPath -Name "NoControlPanel" -Value 1 -Type DWord -Force
        $controlPanelDisabled = Get-ItemProperty -Path $controlPanelPath -Name "NoControlPanel"
        if ($controlPanelDisabled.NoControlPanel -eq 1) {
            Write-Host "Successfully disabled Control Panel for $($user.SamAccountName)." -ForegroundColor Green
        } else {
            Write-Host "Failed to disable Control Panel for $($user.SamAccountName)." -ForegroundColor Red
        }
    } catch {
        Write-Host "An error occurred while processing $($user.SamAccountName): $_" -ForegroundColor Red
    }
}

Write-Host "Script execution completed." -ForegroundColor Yellow
