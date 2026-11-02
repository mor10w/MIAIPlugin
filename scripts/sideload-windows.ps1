# Sideloader et MIAI-manifest i Word, Excel og PowerPoint på Windows (aktuel bruger).
# Bruger samme registreringsnøgle som Microsofts office-addin-dev-settings.
#   npm run sideload:windows                                  → manifest.dev.xml (direkte, localhost:3000)
#   npm run sideload:windows -- packages\addin\manifest.xml   → bridge-varianten
#   npm run sideload:windows -- C:\sti\til\manifest.xml       → hostet variant (fx fra GitHub Pages)
param([string]$Manifest = "$PSScriptRoot\..\packages\addin\manifest.dev.xml")
$path = (Resolve-Path $Manifest).Path
$id = ([xml](Get-Content $path)).OfficeApp.Id
$key = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
New-Item -Path $key -Force | Out-Null
Set-ItemProperty -Path $key -Name $id -Value $path
Write-Host "Manifest registreret ($id): $path"
Write-Host "Genstart Word/Excel/PowerPoint. Knappen 'AI-assistent' findes under fanen Hjem."
# Fjern igen:  Remove-ItemProperty -Path $key -Name $id
