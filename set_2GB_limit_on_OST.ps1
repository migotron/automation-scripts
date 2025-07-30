# PowerShell script to set 2GB limit on PST and OST files for Outlook
$officeVersions = @("14.0", "15.0", "16.0")
$maxFileSize = 2097152     # 2 GB
$warnFileSize = 2048000    # ~1.95 GB

foreach ($version in $officeVersions) {
    $regPath = "HKCU:\Software\Policies\Microsoft\Office\$version\Outlook\PST"

    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }

    Set-ItemProperty -Path $regPath -Name "MaxFileSize" -Value $maxFileSize -Type DWord
    Set-ItemProperty -Path $regPath -Name "WarnFileSize" -Value $warnFileSize -Type DWord
}

# Apply to Default User (requires admin)
$defaultHive = "HKU\Default"
$defaultUserDat = "C:\Users\Default\NTUSER.DAT"

if (Test-Path $defaultUserDat) {
    reg load "$defaultHive" "$defaultUserDat" | Out-Null

    foreach ($version in $officeVersions) {
        $defaultPath = "$defaultHive\Software\Policies\Microsoft\Office\$version\Outlook\PST"
        if (-not (Test-Path "Registry::$defaultPath")) {
            New-Item -Path "Registry::$defaultPath" -Force | Out-Null
        }

        Set-ItemProperty -Path "Registry::$defaultPath" -Name "MaxFileSize" -Value $maxFileSize -Type DWord
        Set-ItemProperty -Path "Registry::$defaultPath" -Name "WarnFileSize" -Value $warnFileSize -Type DWord
    }

    reg unload "$defaultHive" | Out-Null
}

Write-Host "✅ PST/OST size limits applied for current user and Default User (if admin)."
