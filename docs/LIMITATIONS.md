# Limitaciones conocidas y tradeoffs

## Detección de línea visible (F10/F11)
Word VBA no expone una API perfecta para obtener segmentos visuales envueltos idénticos al motor de render en todos los escenarios (zoom extremo, columnas complejas, tablas anidadas, objetos flotantes).

Wordify usa `Range.MoveStart wdLine` + `Range.MoveEnd wdLine` como aproximación conservadora de la línea visible actual.

### Implicaciones
- En texto normal de párrafos con ajuste de línea, el comportamiento coincide de forma estable con la línea visual esperada.
- En estructuras complejas (tablas, cuadros de texto, saltos manuales especiales), el resultado puede variar levemente respecto a la percepción visual exacta.
- Se evita intencionalmente borrar párrafos completos para reducir riesgo.

## Rango numérico
La conversión completa de grupos está implementada de forma robusta para unidades, miles y millones. Grupos superiores se etiquetan con marcador técnico (`*10^N`) para evitar resultados incorrectos silenciosos.

## Deletreo
Se soportan letras, dígitos y un conjunto amplio de puntuación común en teclado español. Símbolos fuera del mapeo generan error explícito para evitar suposiciones.
