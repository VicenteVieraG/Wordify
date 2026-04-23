# Wordify

Herramienta VBA para Microsoft Word Desktop (Windows) orientada a flujos notariales.

## Estructura de módulos

- `modSetup`: instalación/desinstalación de keybindings F2..F11.
- `modCommands`: entrada principal por comando y despacho por tecla.
- `modTokenDetection`: selección o token adyacente y aproximación de línea visible.
- `modValidation`: validadores de hora, número, terreno y deletreo.
- `modSpanishNumbers`: motor reusable número->texto en español.
- `modConversions`: conversiones de hora, moneda MXN, terreno y deletreo.
- `modDeletion`: F9/F10/F11 (borrado token/línea visible).
- `modUI`: mensajes de error amigables en español.

## Documentación operativa

- Instalación: `docs/INSTALL.md`
- Plan de pruebas manual: `docs/TEST_PLAN.md`
- Limitaciones y tradeoffs (línea visible): `docs/LIMITATIONS.md`
Word add-in to convert numbers to words.

## Install (Windows PowerShell)
Use the install script to sideload the add-in manifest into Word's local add-in folder.

```powershell
./scripts/install-wordify-addin.ps1 -ManifestPath .\manifest.xml
```

### Uninstall
```powershell
./scripts/install-wordify-addin.ps1 -ManifestPath .\manifest.xml -Uninstall
```
