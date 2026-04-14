# Wordify (VBA para Microsoft Word Desktop en Windows)

## 1) Importar módulos VBA

1. Abra Word Desktop (Windows).
2. Presione `Alt + F11` para abrir el Editor de VBA.
3. En su plantilla objetivo (`Normal.dotm` o una plantilla `.dotm` corporativa), use **File > Import File...** e importe estos archivos:
   - `vba/modSetup.bas`
   - `vba/modCommands.bas`
   - `vba/modTokenDetection.bas`
   - `vba/modValidation.bas`
   - `vba/modSpanishNumbers.bas`
   - `vba/modConversions.bas`
   - `vba/modDeletion.bas`
   - `vba/modUI.bas`
4. Guarde la plantilla como `.dotm`.

## 2) Instalar atajos de teclado

1. En VBA, ejecute la macro `Wordify_InstallKeybindings`.
2. Esto asigna:
   - `F2` hora a texto
   - `F3` deletreo
   - `Shift+F3` deletreo en mayúsculas
   - `F4` hectárea-área-centiárea
   - `F5` cantidad a texto
   - `F6` moneda nacional (MXN)
   - `F7` metros lineales
   - `F8` metros cuadrados
   - `F9` borrar selección/token adyacente
   - `F10` borrar desde cursor al fin de línea visible
   - `F11` borrar línea visible actual

## 3) Dónde se guardan los atajos

La macro instala los atajos en `Normal.dotm` (CustomizationContext = `NormalTemplate`) para disponibilidad general del usuario.

## 4) Desinstalación de atajos

Ejecute `Wordify_RemoveKeybindings`.
