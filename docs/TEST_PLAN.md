# Plan manual de pruebas (Wordify)

## Preparación
- Documento nuevo en Word Desktop (Windows).
- Cursor en párrafo de texto normal, fuente proporcional y con ajuste de línea habilitado.
- Atajos instalados con `Wordify_InstallKeybindings`.

## Casos por tecla

### F2
- `12:03` + F2 => `12:03 doce horas tres minutos`
- `7:05` + F2 => válido
- `24:00` + F2 => error
- `12:60` + F2 => error

### F3
- `VIGv010` + F3 => `letra "UVE", ...`
- `abc def` seleccionado + F3 => error
- `A-1` + F3 => válido (`guion` soportado)
- símbolo no mapeado (ej. `~`) + F3 => error

### Shift+F3
- `VIGv010` + Shift+F3 => salida totalmente en MAYÚSCULAS

### F4
- `01-02-03` + F4 => válido
- `1-2-3` + F4 => error
- `01:02:03` + F4 => error

### F5
- `123` + F5 => válido
- `00123` + F5 => válido
- `123.25` + F5 => válido
- `123.` + F5 => error
- `-1` + F5 => error
- `1,234` + F5 => error

### F6
- `123` + F6 => válido
- `123.01` + F6 => `un centavo`
- `123.23` + F6 => `centavos`
- `-12` + F6 => error

### F7
- `123` + F7 => válido
- `0` + F7 => válido
- `-2` + F7 => error

### F8
- `123` + F8 => válido
- `0.5` + F8 => válido
- `-2` + F8 => error

### F9
- Con selección: elimina exactamente la selección.
- Sin selección y con cursor adyacente a token: elimina solo ese token.

### F10
- En medio de una línea visual envuelta: elimina desde cursor hasta final de esa línea visual.

### F11
- En una línea visual envuelta: elimina solo la línea visual actual.

## Validación de no modificación en fallo
- Para cada error esperado, confirmar que el documento no cambia.
