# Set Office version (adjust if needed)
$officeVersion = "16.0"
$logPath = "C:\ProgramData\Barcom\Logs"
$logFile = "$logPath\OutlookOnlineMode_Deployment.log"
# Ensure log directory exists
if (-not (Test-Path $logPath)) {
    New-Item -Path $logPath -ItemType Directory -Force | Out-Null
}
# Start logging
Add-Content -Path $logFile -Value "=== Deployment started: $(Get-Date) ==="
# Loop through all loaded user profiles
Get-ChildItem "HKU:\" | ForEach-Object {
    $sid = $_.PSChildName
    $cachedModePath = "Registry::HKU\$sid\Software\Microsoft\Office\$officeVersion\Outlook\Cached Mode"
    try {
        if (-not (Test-Path $cachedModePath)) {
            New-Item -Path $cachedModePath -Force | Out-Null
        }
        Set-ItemProperty -Path $cachedModePath -Name "Enable" -Value 0 -Force
        Add-Content -Path $logFile -Value "[$(Get-Date)] Cached Mode disabled for SID: $sid"
    }
    catch {
        Add-Content -Path $logFile -Value "[$(Get-Date)] ERROR for SID: $sid - $_"
    }
}
Add-Content -Path $logFile -Value "=== Deployment completed: $(Get-Date) ===`n"
