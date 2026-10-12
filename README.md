# MACRO_NETEOS_BMSC

Macro de Excel (VBA) para **netear y clasificar partidas de mayores contables** en auditoría:
cruza, por ejemplo, el mayor de gasto (45902) con el mayor de pasivo (24399) en una sola hoja,
del criterio más sólido al más débil, y clasifica cada partida como Provisiones/Reversiones, Neto,
Pagos, Ajuste u Otras Provisiones/Reversiones.

| Archivo | Contenido |
|---|---|
| `Motor_Neteo_AM_v7.bas` | Macro actual (v7). Macro a ejecutar: `Account_Offset_Engine_AM_v7` |
| `historico/Account_Offset_Engine_AM_v6.bas` | Versión anterior (v6), como referencia |

## Instalación

1. Abrir el libro de Excel y presionar `Alt + F11`.
2. Menú **Archivo → Importar archivo...** y elegir `Motor_Neteo_AM_v7.bas`.
3. Guardar el libro como `.xlsm` (o importar la macro en el libro personal `PERSONAL.XLSB` para usarla en cualquier archivo).

Requiere Excel para Windows (usa `Scripting.Dictionary`, igual que la v6).

## Uso

Antes de ejecutar: pegar los mayores a cruzar **en una sola hoja**, uno debajo del otro, con los mismos
encabezados (por ejemplo gasto 45902 y pasivo 24399). Guardar una copia del archivo, porque los
cambios de una macro no se pueden deshacer con Ctrl+Z.

1. **Importe**: clic en el encabezado del saldo. Si el mayor trae Debe y Haber por separado, seleccionar
   los dos encabezados con Ctrl+clic (aunque haya columnas entre ellos); la macro calcula Debe − Haber.
2. **Cuenta contable**: clic en el encabezado de la cuenta. El grupo de cuenta es el primer dígito por
   defecto (2 = Pasivo, 4 = Gasto).
3. **Base detectada**: la macro muestra las filas de datos que encontró. Se detiene en la primera fila
   vacía, así que no toma totales, leyendas ni tablas dinámicas que estén debajo.
4. **Filas ocultas o filtradas**: si las hay, se puede procesar solo lo visible (útil para segmentar).
5. **Niveles de cruce**: la macro reconoce los encabezados del mayor y propone los niveles. Se pueden
   aceptar, elegir uno por uno o agregar combinaciones propias con Ctrl+clic en los encabezados.
6. **Corte preliminar**: resumen de cruces por nivel y de las partidas pendientes por cuenta y signo.
7. **Opción final** (opcional): Fuerza bruta (FB) o Absorción asimétrica (ASI).
8. **Partidas sin cruce**: nombres por defecto (Pagos / Ajuste / Otras Provisiones/Reversiones),
   editables.
9. **Ubicación de resultados**: al final de la tabla o a la derecha de la columna que se elija.
10. **Dictamen final**: totales por clasificación y por grupo de cuenta, con control de cuadre.

La macro recuerda en el equipo los niveles, la materialidad y los nombres usados la última vez.

## Columnas de resultado

| Columna | Ejemplo | Uso |
|---|---|---|
| Clasificación | Provisiones/Reversiones | Categoría general para filtrar |
| Detalle Clasificación | Provisiones/Reversiones 45902 | Categoría con la cuenta de resultado |
| ID Cruce | PREV-0001 | Las partidas de un mismo cruce comparten el ID |
| Criterio | NRO_TRANSACCION + GLOSA | Sustento del cruce (columnas usadas) |

Prefijos del ID: `PREV` (cruce entre cuentas distintas), `NET` (dentro de la misma cuenta),
`FB` (fuerza bruta) y `ASI` (absorción asimétrica).

## Cómo cruza

Cada cruce es siempre Debe (+) contra Haber (−) por el mismo importe; no se necesita columna ABS.

| Debe (+) | Haber (−) | Resultado |
|---|---|---|
| Gasto | Pasivo | Provisiones/Reversiones (constitución) |
| Pasivo | Gasto | Provisiones/Reversiones (reversión) |
| Gasto | Gasto | Neto |
| Pasivo | Pasivo | Neto |

Niveles, del más sólido al más débil (cada partida queda en el primer nivel que la cruza):

1. **Referencia de comprobante en la glosa**: la glosa cita el número de otro comprobante con importe
   exactamente opuesto (por ejemplo `REV CPTE 250725340`, `REG CPTE ...`). Funciona con cualquier formato
   de glosa: busca números de 5 o más dígitos, sin depender de palabras como CPTE o CBT.
2. **NRO_TRANSACCION + GLOSA** (el CONCATENAR del año anterior, sin columnas auxiliares).
3. **MODULO_ORIGEN + TIPO_DOCUMENTO + NRO_DOCUMENTO**.
4. **NOMBRE_USUARIO + FECHA_DOCUMENTO** (el más débil; conviene revisar sus cruces).
5. Combinaciones adicionales que elija el usuario.

En cada nivel de columnas se hacen tres pasos: cruce 1 a 1 (llave + importe), grupos N a N con la misma
llave que suman 0 (un solo ID para todo el grupo) y, si el grupo no cuadra completo, búsqueda del
subconjunto que sí suma 0 dentro del grupo. Las llaves se normalizan (mayúsculas, sin tildes ni signos,
`00123` = `123`, fechas por número de serie), así que diferencias de formato entre mayores no impiden el cruce.

**Opción final** sobre lo pendiente:

- **Fuerza bruta (FB)**: combinaciones que suman 0 solo por importe, con límite de materialidad,
  de partidas por combinación (2 = pares 1 a 1, recomendado) y de tiempo (120 s). Se puede restringir
  a partidas que compartan una columna. Ante importes repetidos, elige la contrapartida con la glosa
  más parecida.
- **Absorción asimétrica (ASI)**: como la fase 4 de la v6, por grupo de cuenta. Pide autorización
  mostrando la varianza.

**Partidas sin cruce**: Gasto (+) → Pagos, Gasto (−) → Ajuste, Pasivo (+) → Pagos,
Pasivo (−) → Otras Provisiones/Reversiones. Si en el pasivo la glosa cita otra cuenta de gasto que no
está en la base (ej. `45299101`), se clasifica como `Otras Provisiones/Reversiones 452`.

Los números guardados como texto en las columnas de importe (`1.234,56`, `(1.234,56)`, `1234-`) se
convierten a número y se aplica formato contable; las fórmulas no se tocan.

## Parámetros ajustables

Al inicio del módulo hay constantes que se pueden cambiar: prefijos (`PFX_PREV`...), nombres de
categoría (`CAT_PREV`...), encabezados de salida, tamaños máximos de grupos, dígitos mínimos de una
referencia en la glosa y tiempo máximo de la fuerza bruta.

## Validación

Antes de entregar la v7 se replicó su lógica en Python y se comparó con la clasificación manual del
año anterior: coincidió en el 98% de las partidas. Las diferencias corresponden a criterios que el
papel de trabajo aplicó de forma distinta en casos parecidos. Además, el motor VBA se ejecutó sobre
los mismos datos en LibreOffice (modo compatible con VBA) y dio resultados idénticos fila por fila
al prototipo. Lo que no se pudo probar fuera de Excel son los cuadros de diálogo y la escritura en
la hoja.
