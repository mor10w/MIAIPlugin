# winget-manifest (skabelon)

Til at indsende MIAI til [winget-pkgs](https://github.com/microsoft/winget-pkgs), når Windows-installeren er **signeret** og ligger på en **offentlig** adresse (ADR-0016).

1. Udgiv en release (tag `vX.Y.Z`). Pages lægger installeren på `https://mor10w.github.io/MIAIPlugin/updates/MIAI-Setup-X.Y.Z-x64.exe`.
2. Udfyld `{{VERSION}}` og `{{SHA256}}` i de tre filer. SHA-256 står i `latest.yml` (som base64; omregn med `[System.Convert]::FromBase64String(...)`), eller beregn den med `Get-FileHash`.
3. Valider og indsend:
   ```powershell
   winget validate --manifest .
   ```
   ```powershell
   wingetcreate submit --token <GitHub-token> .
   ```

Installeren er pr. bruger (ingen admin) og opdaterer sig selv; winget bruges til den første installation og til central udrulning (`winget install MIAI.MIAI`).
