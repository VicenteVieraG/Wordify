# Instalacion de Wordify

Wordify se instala como modulos VBA en Microsoft Word Desktop para Windows. El destino predeterminado es `Normal.dotm`, de modo que los comandos y atajos queden disponibles para el usuario actual.

## Requisitos

- Windows.
- Microsoft Word Desktop instalado.
- PowerShell.
- Word cerrado antes de instalar o desinstalar.
- Acceso confiable al proyecto VBA habilitado en Word:
  `File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model`.

## Instalacion automatica desde este checkout

Desde la raiz del repositorio:

```powershell
.\scripts\install-wordify-addin.ps1 -SkipSourceBootstrap
```

El parametro `-SkipSourceBootstrap` indica que el codigo fuente ya esta disponible localmente. El script usa `.\vba` como origen, importa los modulos en `Normal.dotm` e instala los atajos de teclado.

Si PowerShell bloquea la ejecucion del script:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\install-wordify-addin.ps1 -SkipSourceBootstrap
```

Tambien puede ejecutar `Install-Wordify.bat` desde un paquete que incluya `scripts\`, `vba\` y `manifest.xml`. El lanzador usa `-SkipSourceBootstrap` solamente cuando detecta esos archivos locales.

## Instalacion desde cero

Si el repositorio no existe localmente, el instalador puede preparar el codigo fuente en `%USERPROFILE%\Wordify`. En ese modo intenta usar Scoop y Git, instalando lo necesario para el usuario actual si falta.

```powershell
.\scripts\install-wordify-addin.ps1
```

Desde un paquete minimo que solo incluya `Install-Wordify.bat` y `install-wordify-addin.ps1` en la misma carpeta, el lanzador ejecuta el mismo flujo desde cero. En ese modo el script instala o localiza Scoop, instala o localiza Git, clona el codigo fuente en `%USERPROFILE%\Wordify` y continua la instalacion desde esa carpeta.
Si `%USERPROFILE%\Wordify` ya existe pero no contiene una fuente valida de Wordify, el script usa `%USERPROFILE%\Wordify\.wordify-source` como carpeta de codigo, salvo que se indique `-ClonePath` explicitamente.

Opciones relacionadas:

```powershell
.\scripts\install-wordify-addin.ps1 -RepositoryUrl https://github.com/VicenteVieraG/Wordify.git -ClonePath "$HOME\Wordify"
```

## Opciones utiles

Vista previa sin aplicar cambios:

```powershell
.\scripts\install-wordify-addin.ps1 -SkipSourceBootstrap -WhatIf
```

Mostrar Word durante la automatizacion:

```powershell
.\scripts\install-wordify-addin.ps1 -SkipSourceBootstrap -Visible
```

Instalar modulos sin atajos:

```powershell
.\scripts\install-wordify-addin.ps1 -SkipSourceBootstrap -SkipKeybindings
```

Usar otra carpeta de modulos VBA:

```powershell
.\scripts\install-wordify-addin.ps1 -VbaPath C:\ruta\a\Wordify\vba -SkipSourceBootstrap
```

## Desinstalacion

```powershell
.\scripts\install-wordify-addin.ps1 -Uninstall
```

La desinstalacion elimina los modulos VBA de Wordify y, salvo que se use `-SkipKeybindings`, tambien elimina sus atajos. Tambien elimina la carpeta de codigo fuente de Wordify indicada por `-ClonePath` o, de forma predeterminada, `%USERPROFILE%\Wordify`; si se uso la carpeta alternativa automatica, elimina `%USERPROFILE%\Wordify\.wordify-source`. Por seguridad, el script solo borra carpetas que todavia parecen una fuente valida de Wordify.

## Atajos instalados

- `F2`: hora a texto.
- `F3`: deletreo.
- `Shift+F3`: deletreo en mayusculas.
- `F4`: hectarea-area-centiarea.
- `F5`: cantidad a texto.
- `F6`: moneda nacional MXN.
- `F7`: metros lineales.
- `F8`: metros cuadrados.
- `F9`: borrar seleccion o token adyacente.
- `F10`: borrar desde cursor al fin de la linea visible.
- `F11`: borrar linea visible actual.

Los atajos se guardan en `Normal.dotm` mediante `CustomizationContext = NormalTemplate`.

## Instalacion manual alternativa

Use este flujo solo si no puede ejecutar el instalador PowerShell.

1. Abra Word Desktop en Windows.
2. Presione `Alt + F11` para abrir el Editor de VBA.
3. En `Normal.dotm` o en una plantilla `.dotm` corporativa, use `File > Import File...` e importe:
   - `vba/modSpanishNumbers.bas`
   - `vba/modConversions.bas`
   - `vba/modValidation.bas`
   - `vba/modTokenDetection.bas`
   - `vba/modDeletion.bas`
   - `vba/modUI.bas`
   - `vba/modCommands.bas`
   - `vba/modSetup.bas`
4. Ejecute la macro `Wordify_InstallKeybindings` para crear los atajos.
5. Guarde la plantilla.

Para quitar solamente los atajos desde VBA, ejecute `Wordify_RemoveKeybindings`.

## Notas

- Cierre Word antes de ejecutar el script. Si existe un proceso `WINWORD`, el instalador se detiene para evitar conflictos con `Normal.dotm`.
- Si Word bloquea la automatizacion VBA, revise la opcion de Trust Center indicada en requisitos.
- `manifest.xml` no participa en esta instalacion; es un artefacto de add-in Office y no representa el flujo operativo actual.
