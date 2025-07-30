# PowerShell script to enforce 'Empty Deleted Items on Exit' policy for Outlook
# and disable the confirmation prompt

$officeVersions = @("14.0", "15.0", "16.0")

foreach ($version in $officeVersions) {
    $policyPath = "HKCU:\Software\Policies\Microsoft\Office\$version\Outlook\Preferences"

    if (-not (Test-Path $policyPath)) {
        New-Item -Path $policyPath -Force | Out-Null
    }

    # Set EmptyTrash to 1 (enable empty deleted items on exit)
    Set-ItemProperty -Path $policyPath -Name "EmptyTrash" -Value 1 -Type DWord

    # Set PromptForDelete to 0 (disable confirmation prompt)
    Set-ItemProperty -Path $policyPath -Name "PromptForDelete" -Value 0 -Type DWord
}

# Apply to Default User profile for new users
$defaultUserHive = "HKU\Default"
$defaultUserDat = "C:\Users\Default\NTUSER.DAT"

if (Test-Path $defaultUserDat) {
    reg load $defaultUserHive $defaultUserDat | Out-Null

    foreach ($version in $officeVersions) {
        $defaultRegPath = "$defaultUserHive\Software\Policies\Microsoft\Office\$version\Outlook\Preferences"
        reg add $defaultRegPath /v EmptyTrash /t REG_DWORD /d 1 /f | Out-Null
        reg add $defaultRegPath /v PromptForDelete /t REG_DWORD /d 0 /f | Out-Null
    }

    reg unload $defaultUserHive | Out-Null
}

Write-Host "✅ Outlook policy settings applied for current user and Default User profile."
