# Nelson Nurseries: diagnóstico financiero y proyección 2020 al 30%

Cifras en miles de USD salvo indicación. Todas las cifras salen de `nelson_nurseries_modelo.xlsx` (fórmulas vivas) y se verificaron contra un cálculo independiente en Python (`calculos.py`). Ver la sección 8.

---

## 1. Resumen ejecutivo

- **El negocio es rentable y está mejorando, pero se está quedando sin caja.** Entre 2017 y 2019 ganó USD 138,1 mil y aun así la caja cayó de 120,1 a 9,4 (0,9% de las ventas, contra una meta de 8%). El 96% de lo que generó la operación (utilidad + depreciación = 253,7) se fue a inventario (188,6) y cartera (55,8).
- **El problema no es de rentabilidad sino de capital de trabajo y de financiamiento.** El ciclo de conversión de caja subió de 451 a 517 días (benchmark: 381). Cada dólar de venta adicional inmoviliza unos 0,74 dólares de capital de trabajo.
- **Con 100% de reinversión y sin deuda, el negocio solo puede crecer 5,7% al año.** Roger plantea 30%.
- **Crecer 30% en 2020 exige unos USD 290 mil de financiamiento externo** (25% del patrimonio) si se repone la caja a 8% de las ventas, o 189 mil si la caja se mantiene en el nivel actual. Incluso sin crecer harían falta 47,7 mil para reponer la caja y pagar el terreno.
- **Sin deuda, el plan de crecimiento no se puede financiar.** La política de Christine de no endeudarse ya no es compatible con el crecimiento de la empresa.
- **Recomendación:** (1) contratar una línea de crédito comprometida de unos USD 300 mil más una hipoteca para el terreno; (2) bajar la cartera a 45 días y el inventario a 450 días; (3) con la línea comprometida, bajar la caja objetivo a 3% de las ventas. Con eso la deuda a fin de 2020 baja a unos **USD 147 mil**, con deuda/patrimonio de 0,12x y cobertura de intereses de 14,4x.
- **Mantener el descuento de proveedores.** Dejar de tomarlo equivale a financiarse al 37% anual.

---

## 2. Verificación de los datos del caso

Antes de analizar, se verificó la consistencia interna de los Exhibits:

| Prueba | Resultado |
|---|---|
| Subtotales del Exhibit 1 (utilidad bruta, EBIT, utilidad neta, activo corriente, activo total, pasivo corriente) | Cuadran |
| Patrimonio = patrimonio anterior + utilidad neta (2017-2019) | Cuadra exacto: **no hubo dividendos ni aportes**. Todo el ingreso familiar es el salario de USD 50 mil registrado en SG&A |
| Activo fijo = AF anterior + capex − depreciación | Cuadra exacto |
| Flujo de caja indirecto vs. variación de caja del balance | Diferencia 0 los tres años |
| 39 ratios del Exhibit 2 recalculados | Todos coinciden. Las diferencias son de redondeo (≤ 0,05 días o ≤ 0,05 pp) |
| Crecimiento 2016 (2,9%) | No se puede recalcular porque falta la venta de 2015; se toma tal como lo da el caso |

**Observación:** la tasa efectiva de impuestos saltó de 33,5% en 2018 a **39,2% en 2019**. El caso no lo explica. En la proyección se usa 39,2% (supuesto conservador), pero conviene revisarlo con el contador.

---

## 3. Indicadores principales y comentarios

| Indicador | 2016 | 2017 | 2018 | 2019 | Benchmark |
|---|---:|---:|---:|---:|---:|
| Crecimiento de ventas | 2,9% | 2,4% | 12,5% | 15,5% | −1,8% |
| Margen bruto | 48,9% | 46,9% | 51,8% | 52,0% | 48,9% |
| SG&A / ventas | 38,2% | 37,4% | 39,2% | 38,6% | – |
| Margen operativo | 6,4% | 4,8% | 8,6% | 9,5% | 7,6% |
| Margen neto | 4,1% | 3,1% | 5,7% | 5,8% | 2,8% |
| ROA | 3,2% | 2,4% | 4,7% | 5,1% | 2,9% |
| ROE | 3,3% | 2,5% | 4,8% | 5,4% | 4,0% |
| Días de cartera (DSO) | 41,9 | 45,0 | 48,0 | 50,9 | 21,8 |
| Días de inventario (DIO) | 424,2 | 432,1 | 436,5 | 476,3 | 386,3 |
| Días de proveedores (DPO) | 15,6 | 13,3 | 10,2 | 9,9 | 26,9 |
| Ciclo de conversión de caja | 450,6 | 463,7 | 474,3 | 517,4 | 381,2 |
| Rotación de activo fijo | 2,37x | 2,43x | 2,36x | 3,01x | 2,7x |
| Caja / ventas (meta 8%) | 15,2% | 13,0% | 7,4% | **0,9%** | – |
| Razón corriente | 19,5x | 17,1x | 17,0x | 17,6x | – |
| Prueba ácida (caja + CxC) / PC | 5,9x | 4,8x | 4,3x | 3,3x | – |
| Capital de trabajo operativo / ventas | 69,0% | 72,3% | 68,5% | 74,1% | – |

**Comentarios por indicador**

1. **Crecimiento de ventas.** Se aceleró de 2-3% a 15,5% mientras el sector cayó 1,8%. Es un logro comercial. Dos causas: el giro hacia plantas maduras ("instant landscape") y el crédito más flexible a viveros pequeños. La segunda causa es la que presiona la cartera.
2. **Margen bruto (52,0%).** Supera al benchmark en 3,1 pp y viene subiendo. Es coherente con el sobreprecio de las plantas maduras. Riesgo: depende del salario H-2A de USD 10,32/h, que el Congreso podría subir.
3. **SG&A / ventas.** Estable en 37-39%, así que no hay descontrol de gastos. Incluye el salario de los dueños (USD 50 mil, 4,8% de las ventas), que es bajo para una familia de seis y no sube con las ventas.
4. **Margen operativo (9,5%).** Duplica al de 2017 y supera al benchmark (7,6%). Las mejoras vienen del margen bruto y de diluir la depreciación.
5. **Margen neto (5,8%).** Duplica al benchmark, pero en 2019 casi no subió (5,7% → 5,8%) por el salto de la tasa de impuestos a 39,2%.
6. **ROA (5,1%) y ROE (5,4%).** Superan al benchmark, pero son bajos en términos absolutos: el ROE es menor que la tasa hipotecaria (6,5%). La causa es un balance muy pesado: se necesitan 1,13 dólares de activos por cada dólar de venta. Como la empresa no tiene deuda, ROE ≈ ROA.
7. **Días de cartera (50,9).** Subieron 9 días en tres años y son más del doble del benchmark (21,8). Es el costo de las mejores condiciones de crédito a viveros pequeños, y también el riesgo: esos clientes son los más frágiles financieramente.
8. **Días de inventario (476).** Es el problema central. Subieron 40 días solo en 2019 (inventario +25,5% contra un costo de ventas +15%) y están 90 días sobre el benchmark. Las causas son el giro a plantas más maduras (más tiempo en el vivero), el aumento de especies en más de 40% y el inventario sembrado para el crecimiento de 2020. Parte del aumento es inversión deliberada, pero cada 10 días de inventario equivalen a unos USD 18 mil de caja con el costo de ventas de 2020.
9. **Días de proveedores (9,9).** Están por debajo del benchmark (26,9), pero **no es un problema**. Christine paga dentro de los 10 días para obtener el 2% de descuento, que equivale a ganar 37% anual (sección 7, medida 5). Lo único que hay que cuidar es pagar el día 10 y no antes.
10. **Ciclo de conversión de caja (517 días).** La empresa tarda casi un año y medio en recuperar cada dólar invertido en la operación, 136 días más que el sector. Por eso crecer consume caja.
11. **Rotación de activo fijo (3,0x).** Sube porque el capex de 2019 fue casi nulo (4,5 contra una depreciación de 40,9). Con un 30% más de ventas llegaría a 3,6x, muy sobre el benchmark (2,7x): la capacidad física se está agotando. La compra de los 12 acres está justificada.
12. **Caja / ventas (0,9%).** Está en el nivel más bajo y muy por debajo de la meta de 8% (83,9 sobre las ventas de 2019). La empresa no tiene colchón para una mala temporada.
13. **Razón corriente (17,6x).** Es engañosa: el 79% del activo corriente es inventario que tarda de 2 a 5 años en venderse. La prueba ácida cae año a año (5,9x → 3,3x). La liquidez real se deteriora.

---

## 4. Desempeño histórico: utilidades sin caja

| Flujo de caja (método indirecto) | 2017 | 2018 | 2019 | Total |
|---|---:|---:|---:|---:|
| Utilidad neta | 25,3 | 52,0 | 60,8 | 138,1 |
| + Depreciación | 38,4 | 36,3 | 40,9 | 115,6 |
| − Aumento de cartera | −8,9 | −20,0 | −26,9 | −55,8 |
| − Aumento de inventario | −39,3 | −15,8 | −133,5 | −188,6 |
| − Aumento de otros activos corrientes | 1,6 | −3,3 | 1,7 | 0,0 |
| + Aumento de pasivo corriente | 6,8 | 0,5 | 4,1 | 11,4 |
| **Flujo de caja operativo** | 23,9 | 49,7 | **−52,9** | 20,7 |
| − Capex | −38,8 | −88,1 | −4,5 | −131,4 |
| **Flujo de caja libre = variación de caja** | −14,9 | −38,4 | **−57,4** | **−110,7** |

- 2019 fue el año de mayor utilidad y también el de mayor quema de caja. El flujo operativo fue negativo (−52,9) por los 133,5 que se invirtieron en inventario.
- En 2018 el flujo operativo fue positivo, pero el capex de 88,1 se lo llevó.
- Solo se pudo crecer sin deuda porque la caja inicial (120,1) absorbió el déficit. Ese colchón ya se agotó.

**Crecimiento autosostenible.** Con un margen neto de 5,8%, activos/ventas de 1,127 y pasivos espontáneos/ventas de 0,045, el crecimiento máximo sin financiamiento externo es m / (A/V − P/V − m) = **5,7%**. Coincide con el ROE sobre el patrimonio inicial, porque no se reparten dividendos. La empresa ya creció 12,5% y 15,5% (por eso se consumió la caja) y ahora apunta a 30%.

---

## 5. Problemas identificados

1. **Brecha de financiamiento estructural.** El crecimiento deseado (30%) es más de cinco veces el autosostenible (5,7%), y la empresa no tiene deuda ni más caja.
2. **Inventario creciente y cada vez más lento** (476 días). El giro a plantas maduras y el aumento de especies alargan el ciclo. No hay evidencia de que el sobreprecio compense el costo de mantener ese inventario.
3. **Cartera en deterioro** (50,9 días, más del doble del sector), concentrada en viveros pequeños, que son los clientes más vulnerables a una subida de tasas.
4. **Liquidez crítica.** Con 9,4 mil de caja, cualquier imprevisto (helada, retraso de un cliente, temporada floja) obliga a dejar de pagar a tiempo a proveedores o empleados.
5. **Política financiera inconsistente con la estrategia.** "Crecer rápido" y "no endeudarse nunca" no son compatibles con este nivel de capital de trabajo. La preocupación de Christine por el riesgo climático es válida, pero hoy ese riesgo se cubre con una caja que ya no existe.
6. **Riesgos externos concentrados.** El salario H-2A afecta el margen, las tasas de interés afectan la demanda de los clientes y el clima afecta el inventario, que es el principal activo.
7. **Rentabilidad sobre el capital baja.** El ROE de 5,4% está por debajo de la tasa hipotecaria. Los Barton tienen USD 1,13 millones invertidos que rinden menos que un préstamo, y viven con USD 50 mil al año.

---

## 6. Proyección 2020 con crecimiento de 30%

### 6.1 Supuestos (escenario base)

| Supuesto | Valor | Fuente |
|---|---|---|
| Crecimiento de ventas | 30% | Caso (expectativa de Roger) |
| Costo de ventas, SG&A, otros activos corrientes, salarios y otras cuentas por pagar | Mismo % de ventas que 2019 | Ratios 2019 |
| DSO / DIO / DPO | 50,9 / 476,3 / 9,9 días | Ratios 2019 |
| Compras | +30% | Proporcional al crecimiento |
| Capex / depreciación | 75 / 46 | Caso |
| Caja objetivo | 8% de las ventas | Meta de Christine |
| Tasa de impuestos | 39,2% | Tasa efectiva 2019 |
| Tasa de la deuda nueva | 6,5% sobre el saldo de cierre | Tasa hipotecaria del caso |
| Dividendos | 0 | Igual que en la historia |

La deuda es la variable de cierre. Para evitar la referencia circular entre deuda e intereses se usa la solución cerrada D = necesidad antes de intereses / (1 − i·(1 − t)).

### 6.2 Estados proforma (base)

| | 2019 real | 2020 proyectado | Variación |
|---|---:|---:|---:|
| Ventas | 1.048,8 | 1.363,4 | +314,6 |
| Costo de ventas | 503,4 | 654,4 | |
| Utilidad bruta | 545,4 | 709,0 | |
| SG&A | 404,5 | 525,9 | |
| Depreciación | 40,9 | 46,0 | |
| **EBIT** | 100,0 | **137,2** | +37% |
| Intereses | – | 18,8 | |
| Impuestos | 39,2 | 46,4 | |
| **Utilidad neta** | 60,8 | **71,9** | +18% |
| Caja | 9,4 | 109,1 | +99,7 |
| Cuentas por cobrar | 146,4 | 190,3 | +43,9 |
| Inventario | 656,9 | 854,0 | +197,1 |
| Otros activos corrientes | 20,9 | 27,2 | +6,3 |
| Activo fijo neto | 347,9 | 376,9 | +29,0 |
| **Activo total** | 1.181,5 | **1.557,4** | **+375,9** |
| Pasivo corriente | 47,3 | 61,5 | +14,2 |
| **Deuda bancaria requerida** | 0,0 | **289,8** | +289,8 |
| Patrimonio | 1.134,2 | 1.206,1 | +71,9 |
| **Pasivo + patrimonio** | 1.181,5 | **1.557,4** | cuadra |

**Usos y fuentes 2020:** se necesitan 375,9 de activos adicionales (inventario 197,1; caja 99,7; cartera 43,9; activo fijo 29,0; otros 6,3). Los pasivos espontáneos aportan 14,2 y la utilidad 71,9. **Faltan 289,8**, que deben venir de deuda (o de aportes de capital).

Flujo de caja 2020: operativo **−115,1** y libre **−190,1**. La utilidad sube, pero el flujo de caja empeora frente a 2019 (−52,9 y −57,4).

### 6.3 Cómo cambian los indicadores

| Indicador | 2019 | 2020 base | Comentario |
|---|---:|---:|---|
| Crecimiento | 15,5% | 30,0% | Supuesto |
| Margen bruto | 52,0% | 52,0% | Se mantiene por supuesto |
| Margen operativo | 9,5% | 10,1% | Sube porque la depreciación crece menos que las ventas |
| Margen neto | 5,8% | 5,3% | **Baja** por los intereses de la nueva deuda (18,8) |
| ROA | 5,1% | 4,6% | Baja: el activo crece 32% y la utilidad 18% |
| ROE | 5,4% | 6,0% | Sube por el apalancamiento |
| DSO / DIO / DPO | 50,9 / 476,3 / 9,9 | iguales | Por supuesto. Son las palancas de la sección 7 |
| Ciclo de caja | 517 | 517 | |
| Rotación de activo fijo | 3,0x | 3,6x | Capacidad física muy exigida |
| Caja / ventas | 0,9% | 8,0% | Se repone la meta |
| Razón corriente (incluye deuda) | 17,6x | 3,4x | Baja por la deuda, pero sigue holgada |
| Deuda / patrimonio | 0,00x | 0,24x | Apalancamiento moderado |
| Cobertura de intereses (EBIT / intereses) | – | 7,3x | Cómoda para un banco |
| Deuda / EBITDA | – | 1,6x | Cómoda (los bancos suelen aceptar hasta ~3x) |

**¿La deuda destruye valor?** No. El retorno operativo después de impuestos sobre el activo en 2020 es 5,4% (EBIT·(1 − t)/activo), mayor que el costo de la deuda después de impuestos: 6,5% × (1 − 0,392) = 4,0%. Por eso el ROE sube de 5,4% a 6,0% aunque el margen neto baje.

### 6.4 Sensibilidad de la deuda requerida al crecimiento

| Crecimiento 2020 | 0% | 5% | 10% | 15% | 20% | 25% | 30% | 35% |
|---|---:|---:|---:|---:|---:|---:|---:|---:|
| Deuda requerida, caja a 8% | 47,7 | 88,0 | 128,4 | 168,7 | 209,1 | 249,4 | 289,8 | 330,2 |
| Deuda requerida, caja a 0,9% | excedente 28,7 | 6,6 | 43,1 | 79,5 | 116,0 | 152,5 | 189,0 | 225,4 |

Cada 5 puntos de crecimiento cuestan unos USD 40 mil (36 mil sin reponer la caja). Reponer la caja a 8% cuesta unos 100 mil adicionales. Sin deuda y con la caja actual, solo se puede crecer un poco menos de 5%.

### 6.5 Escenarios de riesgo

| Escenario | Utilidad neta | FCL | Deuda requerida | D/P | Cobertura |
|---|---:|---:|---:|---:|---:|
| Base | 71,9 | −190,1 | 289,8 | 0,24x | 7,3x |
| Alza del salario H-2A (+2 pp de costo de ventas) | 53,2 | −244,4 | **344,1** | 0,29x | 4,9x |
| Ventas +15% con inventario armado para +30% | 59,9 | −183,5 | 270,6 | 0,23x | 6,6x |

- **Alza del salario H-2A:** el margen neto cae a 3,9% y la necesidad sube 54 mil, porque el mayor costo también encarece el inventario.
- **Ventas por debajo del plan:** la deuda baja (hay menos cartera y menos caja que financiar), pero son **102 mil más** de lo que se habría necesitado con un plan de 15% (168,7). El inventario ya está sembrado y los días de inventario suben a 538. El riesgo no es pedir el préstamo, sino haber sembrado para una demanda que no llega.

---

## 7. Medidas recomendadas (cuantificadas y con viabilidad)

Impacto de cada medida sobre la deuda requerida a fin de 2020 (escenario de 30%):

| Medida | Deuda requerida | Ahorro vs. base | FCL 2020 | Cobertura |
|---|---:|---:|---:|---:|
| Base | 289,8 | – | −190,1 | 7,3x |
| Cobranza: DSO de 50,9 a 45 días | 266,7 | −23,1 | −167,0 | 7,9x |
| Inventario: DIO de 476 a 450 días | 240,7 | −49,1 | −141,0 | 8,8x |
| Combinado (DSO 45 + DIO 450) | 217,6 | −72,2 | −117,9 | 9,7x |
| **Combinado + caja objetivo de 3% (con línea comprometida)** | **146,6** | **−143,2** | −115,1 | **14,4x** |

**1. Financiar el crecimiento con deuda bancaria estructurada (imprescindible; viabilidad alta).**
- *Qué:* una hipoteca a largo plazo para el terreno de 12 acres (activo de largo plazo con financiamiento de largo plazo, a ~6,5%) y una línea de crédito revolvente comprometida, garantizada con cartera e inventario, para el capital de trabajo.
- *Tamaño:* línea de unos **USD 300 mil**, alrededor de la necesidad base sin medidas. La necesidad se mide a fin de año, y en la temporada alta de primavera el pico intra-año probablemente sea mayor. Además, el escenario H-2A llega a 344 mil. Con el plan recomendado se usarían unos 147 mil a fin de año y el resto queda de holgura.
- *Por qué es viable:* deuda/patrimonio de 0,12x a 0,24x, cobertura de 7x a 14x y deuda/EBITDA de 0,8x a 1,6x. Son métricas cómodas para un banco, y la empresa tiene USD 377 mil en activo fijo y 190 mil en cartera como garantía.
- *El temor de Christine:* que una helada destruya el inventario. Eso se cubre con un seguro agrícola para viveros (en EE. UU. existe cobertura federal para viveros) y con la diversificación en 52 invernaderos y más especies. Así se separa el riesgo climático del riesgo financiero.
- *Alternativa:* aporte de capital (por ejemplo, del padre de Christine). Diluye y es más caro, porque el capital rinde 5,4%. Solo tiene sentido si el banco no presta.

**2. Volver a 45 días de cartera (viabilidad alta).**
- La empresa ya operó con 45 días en 2017, así que no hace falta llegar a los 21,8 del sector.
- *Cómo:* reportes de antigüedad de cartera mensuales, cupos de crédito por cliente pequeño según su historial, cobro activo al vencimiento y suspensión de despachos a cuentas vencidas.
- *Libera:* 22,2 mil de cartera (23,1 de deuda, contando intereses).
- *Riesgo:* perder ventas a viveros pequeños. Hay que aplicarlo de forma selectiva: mantener las condiciones a los buenos pagadores y endurecerlas con los morosos.

**3. Bajar el inventario a 450 días (viabilidad media).**
- Es un punto intermedio entre 2018 (436) y 2019 (476).
- *Cómo:* (a) racionalizar las especies, que subieron más de 40%, eliminando las de baja rotación; (b) liquidar el inventario lento o vencido, aunque sea con descuento; (c) exigir que cada línea de plantas maduras tenga un sobreprecio que cubra el costo de mantenerla un año más en el vivero (financiamiento a ~4% después de impuestos, más riesgo climático y mano de obra).
- *Libera:* 47,2 mil de inventario (49,1 de deuda).
- *Limitación:* el inventario de 2020 ya está sembrado por los ciclos de 2 a 5 años, así que la medida es más efectiva en 2021-2022. Para 2020 es una meta exigente.

**4. Con la línea comprometida, bajar la caja objetivo de 8% a ~3% de las ventas (viabilidad alta, condicionada a la línea).**
- El 8% es razonable sin acceso a crédito. Si hay una línea comprometida disponible, la línea cumple la función de colchón y mantener USD 109 mil ociosos cuesta intereses.
- Con 3% (~41 mil) la deuda baja unos 71 mil más.
- *Condición:* **no** bajar la caja si no hay una línea comprometida.

**5. Mantener el descuento por pronto pago (no es una palanca de caja).**
- No tomar el 2/10 neto 30 cuesta **37,2% anual** (44,6% compuesto).
- Pagar el día 30 liberaría solo 13,2 mil, pero se perderían 4,8 mil de descuento al año, contra 0,9 mil que costaría esa misma suma con la línea bancaria.
- Hay que seguir pagando el día 10, y no antes.

**6. Proteger el margen (viabilidad media-alta).**
- Incluir en los contratos o listas de precios una cláusula de ajuste si sube el salario H-2A. El escenario de +2 pp de costo de ventas le quita 18,7 mil a la utilidad neta.
- Seguir moviendo la mezcla hacia productos premium, pero solo donde el sobreprecio cubra el mayor tiempo de inventario (medida 3).

**7. Escalonar el crecimiento si el banco no aprueba la línea completa (plan de contingencia).**
- Según la sensibilidad, sin deuda y con la caja actual solo se puede crecer un poco menos de 5%. Con unos 80 mil se puede crecer 15% sin reponer la caja.
- Si el financiamiento es limitado, conviene reducir la meta de crecimiento antes que dejar la empresa sin caja.
- El terreno se puede posponer o financiar 100% con la hipoteca. El capex entra casi uno a uno en la necesidad de fondos: posponer todo el capex de 2020 (75, que incluye el terreno) bajaría la deuda en unos 78 mil, y posponer solo el terreno, en la parte que le corresponde. Se comprueba poniendo capex = 0 en la hoja `Supuestos`.

**Otras recomendaciones de gestión**
- Pasar a un presupuesto de caja mensual: el negocio es estacional y la foto de fin de año subestima el pico de necesidad.
- Revisar con el contador el salto de la tasa de impuestos (39,2%).
- Revisar la remuneración de los dueños: USD 50 mil para una familia de seis, sin dividendos, no es sostenible. El escenario con el salario como costo fijo muestra que la utilidad sería 9,5 mil mayor si ese salario no subiera con las ventas.

---

## 8. Verificación de los cálculos

1. **Doble cálculo independiente.** Todo se calcula dos veces: en Python (`calculos.py`) y en Excel con fórmulas (`construir_excel.py` → `nelson_nurseries_modelo.xlsx`). Luego se recalcula el Excel con LibreOffice y `verificar_excel.py` compara **982 celdas** (histórico, ratios, flujos, 9 escenarios y 16 columnas de sensibilidad). Resultado: **0 diferencias** con tolerancia de 1e-6 y **0 errores de fórmula** en las 1.901 fórmulas.
2. **Cuadre contable.** En las 25 columnas proyectadas, activo − (pasivo + deuda + patrimonio) = 0.
3. **Conciliaciones históricas.** Patrimonio, activo fijo y flujo de caja contra la variación de caja: diferencia 0 los tres años.
4. **Exhibit 2.** Los 39 ratios recalculables se reproducen (las diferencias son de redondeo).
5. **Comprobaciones manuales del escenario base:**
   - Ventas: 1.048,8 × 1,3 = 1.363,4.
   - Inventario: 656,9 × 1,3 = 854,0.
   - Necesidad antes de intereses: 1.557,4 − 61,5 − (1.134,2 + 137,2 × 0,608) = 278,3.
   - Deuda con intereses: 278,3 / (1 − 0,065 × 0,608) = 289,8.
   - Utilidad neta: (137,2 − 18,8) × 0,608 = 71,9.
   - Crecimiento autosostenible: 0,058 / (1,127 − 0,045 − 0,058) = 5,7%, igual al ROE sobre el patrimonio inicial (60,8 / 1.073,4).
   - Costo de no tomar el descuento: 0,02/0,98 × 365/20 = 37,2%.

---

## 9. Limitaciones

- La proyección es de fin de año. El negocio es estacional y el pico de necesidad dentro del año puede ser mayor.
- Se supone que el inventario crece con las ventas del año. En la realidad, el inventario de 2020 se sembró en años anteriores y el inventario de fin de 2020 respalda las ventas de 2021-2022.
- El costo de la deuda nueva (6,5%) usa la tasa hipotecaria del caso. Una línea de capital de trabajo podría costar más. Ese supuesto se cambia en la hoja `Supuestos`.
- El benchmark corresponde a productores que cotizan en bolsa, más grandes y con otra mezcla de productos. Es una referencia, no una meta exacta.
- El impacto de las medidas sobre las ventas (por ejemplo, perder clientes por endurecer el crédito) no está modelado.

---

## 10. Archivos

| Archivo | Contenido |
|---|---|
| `nelson_nurseries_modelo.xlsx` | Modelo con fórmulas: Resumen, Supuestos, Histórico, Indicadores, Flujo de caja, Proyección 2020 (9 escenarios), Sensibilidad (16 columnas) y Descuento de proveedores |
| `calculos.py` | Cálculo independiente en Python (histórico, ratios, flujos, proyección y escenarios) |
| `construir_excel.py` | Genera el Excel con fórmulas vivas (azul = dato, negro = fórmula, verde = vínculo, amarillo = supuesto clave) |
| `verificar_excel.py` | Compara el Excel recalculado contra Python |

Para regenerar el modelo:

```bash
cd nelson_nurseries
python calculos.py                 # cálculos y chequeos en consola
python construir_excel.py          # escribe el Excel
# recalcular con LibreOffice (o abrir y guardar en Excel), luego:
python verificar_excel.py          # debe mostrar "diferencias: 0"
```
