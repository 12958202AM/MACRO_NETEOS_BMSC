# Neteo de mayores con fórmulas de Excel (sin macro)

Paso a paso para cruzar el mayor de **gasto (45902)** con el mayor de **pasivo (24399)** y clasificar cada
partida (Provisiones/Reversiones, Neto, Pagos, Ajuste, Otras) usando solo fórmulas. Es el mismo criterio
del papel de trabajo del año anterior, con las correcciones a los errores que tuvo.

> Las fórmulas están en **Excel en español** con separador `;` (configuración de Bolivia). Al final hay una
> tabla con los nombres en inglés.

---

## 1. La idea en 1 minuto

Cada partida tiene **cuenta** (2 = Pasivo, 4 = Gasto) y **lado** (Debe + / Haber −). Eso da 4 casilleros:

| Casillero | Qué suele ser | Se cruza con | Resultado |
|---|---|---|---|
| **4D** Gasto Debe (+) | Provisión del mes o pago directo | **2H** (mismo asiento) | Provisiones/Reversiones |
| **2H** Pasivo Haber (−) | Constitución de la provisión | **4D** | Provisiones/Reversiones |
| **4H** Gasto Haber (−) | Reversión de provisión o anulación | **2D** o **4D** | Provisiones/Reversiones o Neto |
| **2D** Pasivo Debe (+) | Pago de la provisión o reversión | **4H** | Provisiones/Reversiones (si no, es Pago) |

Reglas de oro:
1. Un cruce siempre es **Debe contra Haber** por el **mismo importe**: nunca dos positivos.
   El CONCAT con ABS del año pasado no revisaba esto y cruzó 11 pagos contra pagos (Bs 30.850,79).
2. Primero los criterios **con sustento** (mismo comprobante, reverso que cita el comprobante) y recién al
   final los criterios débiles (solo importe, o fecha + importe).
3. Cuando un importe se repite, hay que cruzar **1 a 1**, no "existe / no existe". Para eso se usa el
   **número de ocurrencia** (paso 5).

---

## 2. Preparar la hoja

1. Pega **los dos mayores en una sola hoja**, uno debajo del otro, con los encabezados en la fila 1.
   Las hojas "Mayor G" y "Mayor P" de 2026 ya tienen las mismas columnas.
2. No dejes filas vacías ni totales en medio.
3. Pon el cálculo en **manual** mientras armas las fórmulas: Fórmulas > Opciones para el cálculo > Manual (F9 calcula).
4. En esta guía las columnas del mayor son:

| Col | Campo | Col | Campo |
|---|---|---|---|
| A | NRO_TRANSACCION | H | GLOSA |
| B | TIPO_TRANSACCION (MAN/AUT) | I | DEBE_BS |
| D | TIPO_DOCUMENTO | J | HABER_BS |
| E | NRO_DOCUMENTO | P | NOMBRE_USUARIO |
| F | MODULO_ORIGEN | T | FECHA_DOCUMENTO |
| G | CUENTA_CONTABLE | | |

Las columnas de ayuda van de **AE en adelante**; si tu archivo es distinto, ajusta las letras. Los datos
van de la fila 2 a la 6000 (rangos fijos `$2:$6000`, más rápidos que columnas completas).

---

## 3. Columnas base (identificar cuenta, lado e importe)

| Col | Nombre | Fórmula (fila 2, copiar hacia abajo) | Para qué |
|---|---|---|---|
| AE | Saldo | `=I2-J2` | Importe con signo |
| AF | Grupo | `=IZQUIERDA(G2;1)` | 2 = Pasivo, 4 = Gasto |
| AG | Lado | `=SI(AE2>0;"D";"H")` | Debe / Haber |
| AH | Casillero | `=AF2&AG2` | 4D, 4H, 2D, 2H |
| AI | Centavos | `=REDONDEAR(ABS(AE2)*100;0)` | Importe exacto sin signo, sin errores de decimales |
| AJ | Glosa limpia | ver fórmula abajo | Evita fallos por mayúsculas, espacios y barras |

Fórmula de **AJ**:

```
=ESPACIOS(MAYUSC(SUSTITUIR(H2;"|";" ")))
```

> Tus hojas 2026 ya traen columnas **Grupo I a V** (1 a 5 dígitos de la cuenta); puedes usarlas en lugar de AF.

**Por qué centavos:** `2800647,07` puede guardarse internamente como `2800647,0699999`. Comparando
`REDONDEAR(...*100;0)` la igualdad es exacta.

---

## 4. Cómo leer la cuenta y la glosa

- **Cuenta de 8 dígitos:** `45902106` = grupo 4 (gasto), cuenta 45902 (alquileres), subcuenta 106.
  `24399103` = grupo 2 (pasivo), cuenta 24399 (otras provisiones).
- **Glosa de provisión:** empieza con las dos cuentas del asiento: `45902106|24399103|3.10.2.24.2|AGENCIA CALACOTO`.
  Si empieza con `45299101` o `45601101`, es la provisión de **otra cuenta de gasto** (452 / 456).
- **Glosa de pago:** `3.10.7.24.2| F6973| LAS BRISAS SA| PAGO ALQUILER...` (F = factura, SF = sin factura).
- **Glosa de reverso:** cita el comprobante original:
  `REV. ... CPTE 250725340`, `REG CPTE 258758454`, `REV CBT 20661131`, `REV. COMP.113174609`, `FV_...`.
- **2026 automático (AUT):** `FT26090..._Movs Contables Externos > CTE.2127...`. El gasto y el pasivo
  tienen **distinto** número de transacción y distinta glosa (ver paso 8).

Columna para saber si la glosa nombra otra cuenta de gasto:

| Col | Nombre | Fórmula |
|---|---|---|
| AS | Cuenta citada | `=SI(Y(IZQUIERDA(H2;1)="4";ESNUMERO(--IZQUIERDA(H2;8)));IZQUIERDA(H2;3);"")` |

Devuelve `459`, `452`, `456`... o vacío.

---

## 5. Nivel 1 — Provisiones: mismo comprobante + misma glosa + mismo importe (CONCAT mejorado)

Es lo que hizo el equipo el año pasado, pero **exigiendo signo contrario** y **cruzando 1 a 1**.

| Col | Nombre | Fórmula |
|---|---|---|
| AK | Llave | `=A2&"/"&AJ2&"/"&AI2` |
| AL | Ocurrencia | `=CONTAR.SI.CONJUNTO($AK$2:AK2;AK2;$AH$2:AH2;AH2)` |
| AM | Casillero contrario | `=ELEGIR(COINCIDIR(AH2;{"4D";"2H";"4H";"2D"};0);"2H";"4D";"2D";"4H")` |
| AN | Provisión | `=SI(CONTAR.SI.CONJUNTO($AK$2:$AK$6000;AK2;$AH$2:$AH$6000;AM2)>=AL2;"PREV";"")` |

**Cómo funciona la ocurrencia:** si una llave aparece 2 veces en 4D y 1 vez en 2H, la primera del 4D es la
ocurrencia 1 (hay 1 en 2H → cruza) y la segunda es la ocurrencia 2 (solo hay 1 en 2H → no cruza).
Así nunca se cruza dos veces la misma contrapartida.

Para ver con **qué fila** se cruzó (trazabilidad, Excel 365):

```
=SI(AN2="";"";INDICE(FILTRAR(FILA($AK$2:$AK$6000);($AK$2:$AK$6000=AK2)*($AH$2:$AH$6000=AM2));AL2))
```

---

## 6. Nivel 2 — Reversos: la glosa cita el comprobante original

### 6.1 Sacar el número citado en la glosa

**Excel 365** (toma el primer número de 5 o más dígitos de la glosa):

```
=SI.ERROR(LET(c;EXTRAE(H2;SECUENCIA(LARGO(H2));1);
              d;SI(ESNUMERO(--c);c;" ");
              t;DIVIDIRTEXTO(ESPACIOS(CONCAT(d));" ");
              f;FILTRAR(t;LARGO(t)>=5);
              INDICE(f;1));"")
```

**Excel 2016/2019** (solo glosas con "CPTE"):

```
=SI.ERROR(ESPACIOS(EXTRAE(H2;HALLAR("CPTE";H2)+5;9));"")
```

Repite la fórmula cambiando "CPTE" por "CBT" o "COMP." si las usa tu mayor.

Ponlo en **AO (Ref. citada)**. En glosas de provisión el primer número puede ser la cuenta (`45902106`);
no pasa nada, porque solo cuenta si existe un comprobante con ese número y con importe opuesto.

### 6.2 Marcar las dos partidas del reverso

| Col | Nombre | Fórmula |
|---|---|---|
| AP | Es reverso | `=SI(AO2="";"";SI(CONTAR.SI.CONJUNTO($A$2:$A$6000;--AO2;$AI$2:$AI$6000;AI2;$AG$2:$AG$6000;SI(AG2="D";"H";"D"))>0;"REV";""))` |
| AQ | Fue reversada | `=SI(CONTAR.SI.CONJUNTO($AO$2:$AO$6000;A2&"";$AI$2:$AI$6000;AI2;$AG$2:$AG$6000;SI(AG2="D";"H";"D"))>0;"REV";"")` |

Ejemplo 2025: `REV. 3.10.7.24.2 CPTE 250725340` (−188.256) cita la transacción 250725340 (+188.256)
→ las dos quedan **Neto**.

> **Decisión de criterio:** el año pasado una provisión + su reversión quedó a veces como Neto (Biscayne)
> y a veces como Provisiones (Villa Montenegro). Elige una regla y aplícala siempre. Esta guía usa
> **Neto** (el efecto se anula), por eso el nivel 2 tiene prioridad en la clasificación (paso 9).

---

## 7. Nivel 3 — Neto dentro de la misma cuenta y mismo comprobante

Para reclasificaciones dentro del gasto (ej. Terminal de buses 227642541: +4.070,38 y −4.070,38 en el
mismo comprobante):

| Col | Nombre | Fórmula |
|---|---|---|
| AU | Llave2 | `=A2&"/"&AI2` |
| AV | Ocurrencia2 | `=CONTAR.SI.CONJUNTO($AU$2:AU2;AU2;$AH$2:AH2;AH2)` |
| AR | Neto mismo cpte | `=SI(CONTAR.SI.CONJUNTO($AU$2:$AU$6000;AU2;$AH$2:$AH$6000;AF2&SI(AG2="D";"H";"D"))>=AV2;"NET";"")` |

Es la misma lógica del paso 5, pero la contrapartida se busca en la **misma cuenta** con el lado opuesto
(4D contra 4H, 2D contra 2H).

---

## 8. Nivel 4 (2026) — Movimientos automáticos: fecha + importe

En 2026 el 87% de los asientos son **AUT** (OPERADOR DEL BATCH). La transferencia del gasto (`CTE...`) y la
"TRANSFERENCIA ENTRANTE" del pasivo (`STE...`) tienen distinto número y distinta glosa, pero **misma
fecha y mismo importe**.

| Col | Nombre | Fórmula |
|---|---|---|
| AW | Llave3 | `=T2&"/"&AI2` |
| AX | Ocurrencia3 | `=CONTAR.SI.CONJUNTO($AW$2:AW2;AW2;$AH$2:AH2;AH2)` |
| AY | Candidato batch | `=SI(Y(B2="AUT";CONTAR.SI.CONJUNTO($AW$2:$AW$6000;AW2;$AH$2:$AH$6000;AM2;$B$2:$B$6000;"AUT")>=AX2);"REVISAR";"")` |

> No lo clasifiques automáticamente: es un criterio **débil**. Revisa una muestra de los "REVISAR" con el
> cliente (¿la transferencia entrante del pasivo es la contrapartida del gasto?) y recién entonces decide.

---

## 9. Clasificación final (una sola fórmula con prioridades)

Columna **AZ (Clasificación)**:

```
=SI(O(AP2="REV";AQ2="REV");"Neto";
 SI(AN2="PREV";"Provisiones/Reversiones";
 SI(AR2="NET";"Neto";
 SI(AF2="4";SI(AE2>0;"Pagos";"Ajuste");
 SI(AE2>0;"Pagos";
 SI(Y(AS2<>"";AS2<>"459");"Otras Provisiones/Reversiones "&AS2;"Otras Provisiones/Reversiones"))))))
```

| Prioridad | Si... | Queda |
|---|---|---|
| 1 | Reverso que cita el comprobante | Neto |
| 2 | Mismo comprobante + glosa + importe, gasto contra pasivo | Provisiones/Reversiones |
| 3 | Mismo comprobante, misma cuenta, signo opuesto | Neto |
| 4 | Gasto (+) sin cruce | Pagos (universo de la muestra) |
| 5 | Gasto (−) sin cruce | Ajuste |
| 6 | Pasivo (+) sin cruce | Pagos (pago de provisiones) |
| 7 | Pasivo (−) cuya glosa cita 452 / 456 | Otras Provisiones/Reversiones 452 / 456 |
| 8 | Pasivo (−) restante | Otras Provisiones/Reversiones |

> **Ojo con las superposiciones:** cada nivel se calcula sobre todas las filas, no solo las pendientes.
> Si una fila es "REV" y además "PREV", la prioridad la deja como Neto y su pareja de provisión puede quedar
> sola. Lo detectas porque el control "Provisiones gasto = − pasivo" del paso 10 no da cero. En ese caso,
> filtra las filas con `AP` o `AQ` = "REV" y `AN` = "PREV" y revísalas a mano.

Si después de revisar el nivel 4 decides aceptar los "REVISAR", agrega `SI(AY2="REVISAR";"Provisiones/Reversiones";...)`
antes de la prioridad 4.

---

## 10. Controles (no cierres sin esto)

| Control | Fórmula | Debe dar |
|---|---|---|
| Neto suma cero | `=SUMAR.SI.CONJUNTO($AE$2:$AE$6000;$AZ$2:$AZ$6000;"Neto")` | 0,00 |
| Provisiones gasto = − pasivo | `=SUMAR.SI.CONJUNTO($AE$2:$AE$6000;$AZ$2:$AZ$6000;"Provisiones/Reversiones")` | 0,00 |
| Total gasto = S&S | `=SUMAR.SI.CONJUNTO($AE$2:$AE$6000;$AF$2:$AF$6000;"4")` | Saldo 45902 según EEFF |
| Total pasivo = S&S | `=SUMAR.SI.CONJUNTO($AE$2:$AE$6000;$AF$2:$AF$6000;"2")` | Saldo 24399 según EEFF |
| Pares sin cruce doble | `=CONTAR.SI.CONJUNTO($AN$2:$AN$6000;"PREV";$AF$2:$AF$6000;"4")-CONTAR.SI.CONJUNTO($AN$2:$AN$6000;"PREV";$AF$2:$AF$6000;"2")` | 0 |
| Universo de la muestra | `=SUMAR.SI.CONJUNTO($AE$2:$AE$6000;$AZ$2:$AZ$6000;"Pagos";$AF$2:$AF$6000;"4")` | Igual en Det. muestra y en el resumen |

Luego arma una **tabla dinámica**: Filas = Clasificación, Columnas = Grupo, Valores = Suma de Saldo.
**Actualízala siempre** (clic derecho > Actualizar). El año pasado el resumen salió de tablas dinámicas
desactualizadas y difería en Bs 151.036,79 del universo de la muestra.

---

## 11. Determinación de la muestra (como en 2025)

Sobre las filas **Gasto + Pagos**, ordenadas de mayor a menor:

| Col | Fórmula | Nota |
|---|---|---|
| Alcance % | `=AE2/SUMAR.SI.CONJUNTO($AE$2:$AE$6000;$AZ$2:$AZ$6000;"Pagos";$AF$2:$AF$6000;"4")` | |
| Acumulado % | fila 2: `=Alcance2` ; fila 3: `=Acumulado2+Alcance3` | |
| Muestra | `TT` para las partidas clave (escrito a mano), `NSS` para la selección no estadística | Probablemente *Targeted Testing* / *Non-Statistical Sampling* |
| Resumen | `=SUMAR.SI(Muestra;"TT";Saldo)`, `=CONTAR.SI(Muestra;"TT")` (igual para NSS) | |
| Error tolerable | `=(Universo − TT − Inmateriales) × 20%` | Así lo calculó el equipo en 2025 |

---

## 12. Consejos para que no se cuelgue

- Rangos fijos (`$A$2:$A$6000`), **no** columnas completas (`A:A`).
- Cálculo manual mientras armas; F9 para recalcular.
- Cuando termines un nivel, **copia y pega como valores** las columnas de ese nivel.
- Más de ~50.000 filas: usa **Power Query** (Combinar consultas de la tabla consigo misma por la llave) o la macro.
- Convierte los importes que vengan como texto: Datos > Texto en columnas > Finalizar.

---

## 13. Nombres de funciones español / inglés

| Español | Inglés | Español | Inglés |
|---|---|---|---|
| SI | IF | EXTRAE | MID |
| SI.ERROR | IFERROR | HALLAR | SEARCH |
| Y / O | AND / OR | IZQUIERDA | LEFT |
| CONTAR.SI.CONJUNTO | COUNTIFS | LARGO | LEN |
| SUMAR.SI.CONJUNTO | SUMIFS | ESPACIOS | TRIM |
| CONTAR.SI / SUMAR.SI | COUNTIF / SUMIF | MAYUSC | UPPER |
| BUSCARV / BUSCARX | VLOOKUP / XLOOKUP | SUSTITUIR | SUBSTITUTE |
| INDICE / COINCIDIR | INDEX / MATCH | REDONDEAR | ROUND |
| ELEGIR | CHOOSE | ESNUMERO | ISNUMBER |
| FILTRAR | FILTER | SECUENCIA | SEQUENCE |
| DIVIDIRTEXTO | TEXTSPLIT | LET / CONCAT | LET / CONCAT |
| FILA | ROW | | |

> En la versión en inglés de Excel el separador de argumentos suele ser `,` en lugar de `;`.
