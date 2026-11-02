# Sideloader MIAI-manifestet i Word, Excel og PowerPoint på Windows (aktuel bruger).
# Bruger samme registreringsnøgle som Microsofts office-addin-dev-settings.
# Kør:  powershell -ExecutionPolicy Bypass -File scripts\sideload-windows.ps1
$manifest = (Resolve-Path "$PSScriptRoot\..\packages\addin\manifest.xml").Path
$key = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
New-Item -Path $key -Force | Out-Null
Set-ItemProperty -Path $key -Name "fdac46a3-f2b8-4e84-b704-78a669531789" -Value $manifest
Write-Host "Manifest registreret: $manifest"
Write-Host "Genstart Word/Excel/PowerPoint. Knappen 'AI-assistent' findes under fanen Hjem."
# Fjern igen:  Remove-ItemProperty -Path $key -Name "fdac46a3-f2b8-4e84-b704-78a669531789"
