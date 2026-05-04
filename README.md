# Anywhere365 PnP setup script

Gebruik `run-anywhere365-pnp-setup.ps1` om alle PnP PowerShell stappen in één keer uit te voeren via **Run with PowerShell** (rechtermuisknop op het `.ps1` bestand).

## Wat doet het script?

1. Registreert een Entra ID app met `Register-PnPEntraIDAppForInteractiveLogin`.
2. Gebruikt de gevonden `ClientId` voor `Connect-PnPOnline`.
3. Toont bestaande site permissies.
4. Geeft `Write` permissie aan de opgegeven certificate app.
5. Zet daarna permissie naar `FullControl`.

## Input

Als waarden niet als parameter zijn meegegeven, vraagt het script interactief om:

- Tenant
- SharePoint Site URL
- App ID met certificaat

## Optionele parameters

```powershell
.\run-anywhere365-pnp-setup.ps1 -Tenant "contoso.onmicrosoft.com" -SiteUrl "https://contoso.sharepoint.com/sites/Anywhere365" -CertificateAppId "00000000-0000-0000-0000-000000000000"
```

Optioneel FullControl stap overslaan:

```powershell
.\run-anywhere365-pnp-setup.ps1 -SkipFullControl
```
