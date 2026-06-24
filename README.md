# Wordify

Herramienta VBA para Microsoft Word Desktop en Windows, orientada a flujos notariales. Wordify convierte tokens cercanos al cursor o texto seleccionado a texto en espanol y agrega atajos de teclado para operaciones frecuentes dentro de Word.

## Estado actual

El proyecto funciona como un conjunto de modulos VBA instalados en `Normal.dotm`. El instalador actual es `scripts/install-wordify-addin.ps1`: importa los modulos desde `vba/`, elimina versiones anteriores de los modulos Wordify y, salvo que se indique lo contrario, instala los atajos F2..F11.

`manifest.xml` existe como artefacto de add-in Office/manifest, pero no es el flujo operativo actual. No se usa con el instalador de VBA.

## Instalacion rapida

Requisitos:

- Windows.
- Microsoft Word Desktop instalado.
- Word cerrado antes de ejecutar el instalador.
- En Word, habilitar `File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model`.

Desde la raiz del repositorio:

```powershell
.\scripts\install-wordify-addin.ps1 -SkipSourceBootstrap
```

Para desinstalar:

```powershell
.\scripts\install-wordify-addin.ps1 -Uninstall
```

Mas detalles: `docs/INSTALL.md`.

## Atajos

- `F2`: hora a texto.
- `F3`: deletreo.
- `Shift+F3`: deletreo en mayusculas.
- `F4`: medida de terreno en hectarea-area-centiarea.
- `F5`: numero a texto.
- `F6`: moneda nacional MXN.
- `F7`: metros lineales.
- `F8`: metros cuadrados.
- `F9`: borrar seleccion o token adyacente.
- `F10`: borrar desde cursor al final de la linea visible.
- `F11`: borrar linea visible actual.

## Estructura de modulos

- `modSetup`: instalacion y desinstalacion de keybindings F2..F11.
- `modCommands`: entrada principal por comando y despacho por tecla.
- `modTokenDetection`: seleccion o token adyacente y aproximacion de linea visible.
- `modValidation`: validadores de hora, numero, terreno y deletreo.
- `modSpanishNumbers`: motor reusable numero a texto en espanol.
- `modConversions`: conversiones de hora, moneda MXN, terreno y deletreo.
- `modDeletion`: F9/F10/F11 para borrado de token o linea visible.
- `modUI`: mensajes de error amigables en espanol.

## Documentacion operativa

- Instalacion: `docs/INSTALL.md`
- Plan de pruebas manual: `docs/TEST_PLAN.md`
- Limitaciones y tradeoffs: `docs/LIMITATIONS.md`
