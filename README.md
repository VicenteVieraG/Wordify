# Wordify
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
