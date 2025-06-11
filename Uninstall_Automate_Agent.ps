$zip='C:\\Windows\\Temp\\Agent_Uninstaller.zip';
$dest='C:\\Windows\\Temp\\AutomateUninstall';
Invoke-WebRequest -Uri 'https://s3.amazonaws.com/assets-cp/assets/Agent_Uninstaller.zip' -OutFile $zip -UseBasicParsing;
Expand-Archive -Path $zip -DestinationPath $dest -Force;
$exe=Join-Path $dest 'Agent_Uninstall.exe';
if(Test-Path $exe) {
  Start-Process -FilePath $exe -ArgumentList '/S' -Wait;
  Write-Host 'Uninstall complete.'
} else {
  Write-Host 'Uninstaller not found.'
}
