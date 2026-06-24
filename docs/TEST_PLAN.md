# Plan manual de pruebas (Wordify)

## Preparacion

- Documento nuevo en Word Desktop (Windows).
- Cursor en parrafo de texto normal, fuente proporcional y con ajuste de linea habilitado.
- Wordify instalado con `.\scripts\install-wordify-addin.ps1 -SkipSourceBootstrap`.
- Como alternativa manual, modulos importados en VBA y atajos instalados con `Wordify_InstallKeybindings`.

## Casos por tecla

### F2

- `12:03` + F2 => `12:03 doce horas tres minutos`
- `7:05` + F2 => valido
- `24:00` + F2 => error
- `12:60` + F2 => error

### F3

- `VIGv010` + F3 => `letra "UVE", ...`
- `abc def` seleccionado + F3 => error
- `A-1` + F3 => valido (`guion` soportado)
- simbolo no mapeado, por ejemplo `~`, + F3 => error

### Shift+F3

- `VIGv010` + Shift+F3 => salida totalmente en MAYUSCULAS

### F4

- `01-02-03` + F4 => valido
- `1-2-3` + F4 => error
- `01:02:03` + F4 => error

### F5

- `123` + F5 => valido
- `00123` + F5 => valido
- `123.25` + F5 => valido
- `123.` + F5 => error
- `-1` + F5 => error
- `1,234` + F5 => error

### F6

- `123` + F6 => valido
- `123.01` + F6 => `un centavo`
- `123.23` + F6 => `centavos`
- `-12` + F6 => error

### F7

- `123` + F7 => valido
- `0` + F7 => valido
- `-2` + F7 => error

### F8

- `123` + F8 => valido
- `0.5` + F8 => valido
- `-2` + F8 => error

### F9

- Con seleccion: elimina exactamente la seleccion.
- Sin seleccion y con cursor adyacente a token: elimina solo ese token.

### F10

- En medio de una linea visual envuelta: elimina desde cursor hasta final de esa linea visual.

### F11

- En una linea visual envuelta: elimina solo la linea visual actual.

## Validacion de no modificacion en fallo

- Para cada error esperado, confirmar que el documento no cambia.
