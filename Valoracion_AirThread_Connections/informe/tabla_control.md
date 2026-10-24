# Tabla de control: cifra del informe → celda del Excel

Cada cifra del informe sale de una celda de `modelo/modelo_airthread.xlsx`. La misma lista está en la hoja **Control** del Excel (columna D vinculada a la celda).

| Clave | Descripción en el Excel | Celda | Valor en el Excel |
|---|---|---|---:|
| C001 | Precio máximo (retiro): valor independiente + 100% backhaul (sin pagar sinergias de ingresos) | `Valoracion!C60` | 7.409,4978 |
| C002 | Apertura: valor independiente con el descuento por iliquidez de una salida a bolsa (10%, nota 6) | `Valoracion!C58` | 6.082,5093 |
| C003 | Objetivo: valor independiente + 50% de las sinergias de backhaul | `Valoracion!C59` | 7.083,9208 |
| C004 | Valor de las operaciones (APV) | `Valoracion!C11` | 6.103,6439 |
| C005 | Total activos no operativos | `Valoracion!C20` | 1.944,4871 |
| C006 | Total deuda e interés minoritario | `Valoracion!C26` | 1.289,7873 |
| C007 | Valor del patrimonio, 100%, sin sinergias (valor intrínseco independiente) | `Valoracion!C33` | 6.758,3437 |
| C008 | + VP sinergias de backhaul | `Valoracion!C36` | 651,1541 |
| C009 | + VP sinergias de ingresos empresariales | `Valoracion!C37` | 1.553,8483 |
| C010 | Valor del patrimonio con todas las sinergias | `Valoracion!C38` | 8.963,3461 |
| C011 | VE/EBITDA 2007 implícito en el valor de las operaciones | `Valoracion!C51` | 5,9068 |
| C012 | Costo de capital no apalancado Ku = rf + βu × PRM | `Costo_Capital!C10` | 0,0823 |
| C013 | CPPC (WACC) objetivo = E/V × Ke + D/V × Kd × (1 − t) | `Costo_Capital!C16` | 0,0762 |
| C014 | Crecimiento a perpetuidad (g) | `Valoracion!C7` | 0,0278 |
| C015 | Valor terminal / valor de las operaciones | `Valoracion!C12` | 0,7472 |
| C016 | Mediana [VE/EBITDA] | `Comparables!P13` | 10,3334 |
| C017 | Patrimonio con valor terminal por múltiplo (sin sinergias) | `Valoracion!C46` | 13.866,9095 |
| C018 | Crecimiento implícito en el múltiplo (resuelve VT = NOPAT₂₀₁₂(1+g)(1−g/ROIC)/(CPPC−g)) | `Valoracion!C47` | 0,0665 |
| C019 | Prima del precio máximo sobre el valor independiente | `Valoracion!C66` | 0,0963 |
| C020 | Sinergias de ingresos / total de sinergias | `Valoracion!C65` | 0,7047 |
| C021 | Patrimonio | `Sensibilidad!I21` | 7.860,8449 |
| C022 | Deuda inicial = intereses máximos / Kd | `Deuda!C5` | 3.757,5455 |
| C023 | Deuda de adquisición / valor de las operaciones (límite 50%) | `Valoracion!C55` | 0,6156 |
| C024 | Deuda / valor máximo admitido [Valor] | `Supuestos!C12` | 0,5000 |
| C025 | Deuda de adquisición / valor total de la empresa con activos no operativos | `Valoracion!C56` | 0,4669 |
| C026 | Línea de crédito (variable de cierre) [2012] | `Proyeccion!H70` | 563,4335 |
| C027 | Costo de la deuda (Kd), BBB+ [Valor] | `Supuestos!C7` | 0,0550 |
| C028 | Diferencial sobre el bono del Tesoro a 10 años [Valor] | `Supuestos!C8` | 0,0125 |
| C029 | Prima de riesgo de mercado (PRM) [Valor] | `Supuestos!C10` | 0,0500 |
| C030 | Cobertura EBITDA / intereses para dimensionar la deuda [Valor] | `Supuestos!C11` | 5,0000 |
| C031 | Múltiplo P/U histórico de la industria [Valor] | `Supuestos!C15` | 19,0950 |
| C032 | Costo de salir a bolsa (% del patrimonio), tope razonable del descuento [Valor] | `Supuestos!C16` | 0,1000 |
| C033 | Descuento por iliquidez, regla práctica (20%-30%), punto medio [Valor] | `Supuestos!C18` | 0,2500 |
| C034 | Descuento por iliquidez, investigación académica [Valor] | `Supuestos!C19` | 0,3500 |
| C035 | Diferencia máxima de todos los subtotales del estado de resultados [Diferencia máx.] | `Historicos!G31` | 0,0490 |
| C036 | Control: activos − pasivos y patrimonio [Diferencia máx.] | `Historicos!G79` | 0,0000 |
| C037 | Diferencia máxima de los totales del balance [Diferencia máx.] | `Historicos!G87` | 0,0470 |
| C038 | Control: diferencia Tabla 3 vs. recalculado [Valor empresa (VE)] | `Comparables!O15` | 0,0005 |
| C039 | ROE (utilidad neta / patrimonio al cierre) [2007] | `Historicos!E110` | 0,0985 |
| C040 | Control: diferencia (fila 16 − fila 14) | `Deuda!Q15` | 0,3765 |
| C041 | Días de cartera (sobre ingresos totales) [Valor] | `Supuestos!C46` | 41,6673 |
| C042 | Días de inventario (sobre ingresos por equipos) [Valor] | `Supuestos!C47` | 154,3602 |
| C043 | Ingresos por equipos / ingresos por servicios [Valor] | `Supuestos!C45` | 0,0750 |
| C044 | Diferencia = dividendos u otros cargos implícitos [2006] | `Historicos!D93` | 0,0000 |
| C045 | Diferencia = dividendos u otros cargos implícitos [2007] | `Historicos!E93` | −40,1900 |
| C046 | Variación de utilidades retenidas publicadas − utilidad neta [2006] | `Historicos!D94` | 68,8860 |
| C047 | Variación de utilidades retenidas publicadas − utilidad neta [2007] | `Historicos!E94` | −137,0700 |
| C048 | Días de ingresos diferidos (sobre ingresos por servicios) [Valor] | `Supuestos!C50` | 14,0121 |
| C049 | Gastos prepagados / costos operativos en efectivo [Valor] | `Supuestos!C48` | 0,0138 |
| C050 | Activos operativos netos implícitos en la Tabla 2 (ingresos / rotación) [2007] | `Historicos!E150` | 3.816,5029 |
| C051 | Activo fijo neto (PP&E + licencias + listas de clientes) [2007] | `Historicos!E96` | 4.092,9170 |
| C052 | Promedio [D/P publicado] | `Comparables!E11` | 0,4009 |
| C053 | Promedio [D/V publicado] | `Comparables!D11` | 0,2812 |
| C054 | Inconsistencia: D/P promedio publicado vs. D/P implícito en el D/V promedio [D/V publicado] | `Comparables!D17` | 0,3912 |
| C055 | Agile Connections [VE/EBITDA] | `Comparables!P7` | 2,6386 |
| C056 | Neuberger Wireless [VE/EBITDA] | `Comparables!P6` | 19,0667 |
| C057 | EBIT (utilidad operativa) [2007] | `Historicos!E13` | 451,0560 |
| C058 | Capex implícito = Δ activo fijo neto + D&A [2006] | `Historicos!D97` | 741,9550 |
| C059 | Capex implícito = Δ activo fijo neto + D&A [2007] | `Historicos!E97` | 525,8150 |
| C060 | Descuento por iliquidez, Anexo 8 (empresa rentable, ingresos US$1.000 MM) [Valor] | `Supuestos!C17` | 0,1570 |
| C061 | Ingresos totales [2007] | `Historicos!E7` | 3.946,2640 |
| C062 | Crecimiento de ingresos totales [2006] | `Historicos!D102` | 0,1460 |
| C063 | Crecimiento de ingresos totales [2007] | `Historicos!E102` | 0,1362 |
| C064 | Variación de los minutos por cliente [2007] | `Historicos!E143` | 0,2202 |
| C065 | Margen EBITDA [2005] | `Historicos!C104` | 0,2300 |
| C066 | Margen EBITDA [2006] | `Historicos!D104` | 0,2491 |
| C067 | Margen EBITDA [2007] | `Historicos!E104` | 0,2618 |
| C068 | Margen EBITDA [Comparables] | `Historicos!G104` | 0,3160 |
| C069 | Margen EBIT [2005] | `Historicos!C105` | 0,0683 |
| C070 | Margen EBIT [2006] | `Historicos!D105` | 0,0891 |
| C071 | Margen EBIT [2007] | `Historicos!E105` | 0,1143 |
| C072 | Margen EBIT [Comparables] | `Historicos!G105` | 0,1560 |
| C073 | D&A / ingresos totales [2007] | `Historicos!E130` | 0,1475 |
| C074 | Margen neto [2005] | `Historicos!C106` | 0,0511 |
| C075 | Margen neto [2006] | `Historicos!D106` | 0,0517 |
| C076 | Margen neto [2007] | `Historicos!E106` | 0,0798 |
| C077 | Margen neto [Comparables] | `Historicos!G106` | 0,0660 |
| C078 | ROA (utilidad neta / activos totales) [2005] | `Historicos!C109` | 0,0286 |
| C079 | ROA (utilidad neta / activos totales) [2006] | `Historicos!D109` | 0,0316 |
| C080 | ROA (utilidad neta / activos totales) [2007] | `Historicos!E109` | 0,0561 |
| C081 | ROE (utilidad neta / patrimonio al cierre) [2005] | `Historicos!C110` | 0,0565 |
| C082 | ROE (utilidad neta / patrimonio al cierre) [2006] | `Historicos!D110` | 0,0600 |
| C083 | Costo del patrimonio Ke = rf + βe × PRM | `Costo_Capital!C15` | 0,0930 |
| C084 | ROIC (EBIT × (1 − t) / (patrimonio + deuda), definición de la nota 7) [2005] | `Historicos!C111` | 0,0320 |
| C085 | ROIC (EBIT × (1 − t) / (patrimonio + deuda), definición de la nota 7) [2006] | `Historicos!D111` | 0,0461 |
| C086 | ROIC (EBIT × (1 − t) / (patrimonio + deuda), definición de la nota 7) [2007] | `Historicos!E111` | 0,0645 |
| C087 | Días de cartera (sobre ingresos totales, 360 días) [2005] | `Historicos!C112` | 43,0417 |
| C088 | Días de cartera (sobre ingresos totales, 360 días) [2006] | `Historicos!D112` | 42,2318 |
| C089 | Días de cartera (sobre ingresos totales, 360 días) [2007] | `Historicos!E112` | 39,7284 |
| C090 | Días de inventario (sobre costo de equipos) [2005] | `Historicos!C114` | 65,2212 |
| C091 | Días de inventario (sobre costo de equipos) [2006] | `Historicos!D114` | 74,1568 |
| C092 | Días de inventario (sobre costo de equipos) [2007] | `Historicos!E114` | 56,7869 |
| C093 | Días de proveedores (sobre costos operativos en efectivo) [2005] | `Historicos!C117` | 39,1996 |
| C094 | Días de proveedores (sobre costos operativos en efectivo) [2006] | `Historicos!D117` | 35,1796 |
| C095 | Días de proveedores (sobre costos operativos en efectivo) [2007] | `Historicos!E117` | 32,2303 |
| C096 | Ciclo de caja = cartera + inventario (costo) − proveedores [2005] | `Historicos!C120` | 69,0633 |
| C097 | Ciclo de caja = cartera + inventario (costo) − proveedores [2006] | `Historicos!D120` | 81,2091 |
| C098 | Ciclo de caja = cartera + inventario (costo) − proveedores [2007] | `Historicos!E120` | 64,2851 |
| C099 | Capital de trabajo operativo neto (CxC + inventario + prepagados − CxP − diferidos − acumulados) [2007] | `Historicos!E192` | 114,6150 |
| C100 | Rotación de activos (ingresos / activos totales) [2005] | `Historicos!C122` | 0,5596 |
| C101 | Rotación de activos (ingresos / activos totales) [2006] | `Historicos!D122` | 0,6114 |
| C102 | Rotación de activos (ingresos / activos totales) [2007] | `Historicos!E122` | 0,7032 |
| C103 | Razón corriente [2005] | `Historicos!C123` | 0,8146 |
| C104 | Razón corriente [2006] | `Historicos!D123` | 0,9979 |
| C105 | Razón corriente [2007] | `Historicos!E123` | 1,3798 |
| C106 | Deuda financiera / patrimonio contable [2005] | `Historicos!C125` | 0,4146 |
| C107 | Deuda financiera / patrimonio contable [2006] | `Historicos!D125` | 0,3464 |
| C108 | Deuda financiera / patrimonio contable [2007] | `Historicos!E125` | 0,3136 |
| C109 | Deuda financiera / patrimonio contable [Comparables] | `Historicos!G125` | 0,4009 |
| C110 | Deuda financiera / EBITDA [2005] | `Historicos!C126` | 1,6303 |
| C111 | Deuda financiera / EBITDA [2006] | `Historicos!D126` | 1,1986 |
| C112 | Deuda financiera / EBITDA [2007] | `Historicos!E126` | 0,9700 |
| C113 | EBITDA / intereses [2005] | `Historicos!C127` | 8,2131 |
| C114 | EBITDA / intereses [2006] | `Historicos!D127` | 9,2342 |
| C115 | EBITDA / intereses [2007] | `Historicos!E127` | 12,2028 |
| C116 | (Caja + valores negociables) / ingresos totales [2005] | `Historicos!C128` | 0,0096 |
| C117 | (Caja + valores negociables) / ingresos totales [2006] | `Historicos!D128` | 0,0812 |
| C118 | (Caja + valores negociables) / ingresos totales [2007] | `Historicos!E128` | 0,0560 |
| C119 | Ingreso por minuto (US$) [2005] | `Historicos!C140` | 0,0724 |
| C120 | Ingreso por minuto (US$) [2006] | `Historicos!D140` | 0,0671 |
| C121 | Ingreso por minuto (US$) [2007] | `Historicos!E140` | 0,0595 |
| C122 | Variación del ingreso por minuto [2007] | `Historicos!E141` | −0,1128 |
| C123 | Costo por cliente nuevo (US$) [2005] | `Historicos!C135` | 372,0000 |
| C124 | Costo por cliente nuevo (US$) [2006] | `Historicos!D135` | 385,0000 |
| C125 | Costo por cliente nuevo (US$) [2007] | `Historicos!E135` | 487,0000 |
| C126 | Variación del costo por cliente nuevo [2007] | `Historicos!E142` | 0,2649 |
| C127 | Tasa mensual de abandono (churn) [2005] | `Historicos!C138` | 0,0210 |
| C128 | Tasa mensual de abandono (churn) [2006] | `Historicos!D138` | 0,0210 |
| C129 | Tasa mensual de abandono (churn) [2007] | `Historicos!E138` | 0,0170 |
| C130 | Utilidad neta [2006] | `Historicos!D170` | 179,4900 |
| C131 | Utilidad neta [2007] | `Historicos!E170` | 314,7340 |
| C132 | + Depreciación y amortización [2006] | `Historicos!D171` | 555,5250 |
| C133 | + Depreciación y amortización [2007] | `Historicos!E171` | 582,2690 |
| C134 | Flujo de caja operativo [2006] | `Historicos!D176` | 726,7670 |
| C135 | Flujo de caja operativo [2007] | `Historicos!E176` | 862,0290 |
| C136 | Flujo de caja de inversión [2006] | `Historicos!D181` | −753,9140 |
| C137 | Flujo de caja de inversión [2007] | `Historicos!E181` | −302,0050 |
| C138 | Flujo de caja de financiación [2006] | `Historicos!D187` | 31,0560 |
| C139 | Flujo de caja de financiación [2007] | `Historicos!E187` | −388,4030 |
| C140 | Variación de caja según flujo [2006] | `Historicos!D188` | 3,9090 |
| C141 | Variación de caja según flujo [2007] | `Historicos!E188` | 171,6210 |
| C142 | Flujo de caja libre operativo = EBIT×(1−t) + D&A − capex − Δ capital de trabajo operativo [2006] | `Historicos!D191` | −55,4782 |
| C143 | Flujo de caja libre operativo = EBIT×(1−t) + D&A − capex − Δ capital de trabajo operativo [2007] | `Historicos!E191` | 345,9976 |
| C144 | Capex implícito / ingresos totales [2007] | `Historicos!E98` | 0,1332 |
| C145 | − Compra neta de valores negociables y títulos de renta variable [2007] | `Historicos!E179` | 237,5600 |
| C146 | Tasa libre de riesgo (Tesoro 10 años) | `Costo_Capital!C4` | 0,0425 |
| C147 | PRM 4% (supuesto propio, sensibilidad) [Valor] | `Costo_Capital!C27` | 0,0400 |
| C148 | PRM 6% (supuesto propio, sensibilidad) [Valor] | `Costo_Capital!C28` | 0,0600 |
| C149 | PRM 4% (supuesto propio, sensibilidad) [Patrimonio] | `Costo_Capital!F27` | 7.773,0706 |
| C150 | Beta: 3 Harris-Pringle, βd = 0,25 (elegida) [Patrimonio] | `Costo_Capital!F26` | 6.758,3437 |
| C151 | PRM 6% (supuesto propio, sensibilidad) [Patrimonio] | `Costo_Capital!F28` | 6.008,3836 |
| C152 | Beta: 1 Hamada, βd = 0 [Valor] | `Costo_Capital!C24` | 0,8155 |
| C153 | Beta: 2 Harris-Pringle, βd = 0 [Valor] | `Costo_Capital!C25` | 0,7264 |
| C154 | Beta de la deuda implícita en el CAPM = (Kd − rf) / PRM | `Costo_Capital!C7` | 0,2500 |
| C155 | Beta: 3 Harris-Pringle, βd = 0,25 (elegida) [Valor] | `Costo_Capital!C26` | 0,7967 |
| C156 | Beta: 1 Hamada, βd = 0 [Patrimonio] | `Costo_Capital!F24` | 6.747,8895 |
| C157 | Beta: 2 Harris-Pringle, βd = 0 [Patrimonio] | `Costo_Capital!F25` | 7.208,3436 |
| C158 | D/V objetivo: promedio comparables (elegida) [Valor] | `Costo_Capital!C29` | 0,2812 |
| C159 | D/V objetivo: mediana comparables [Valor] | `Costo_Capital!C30` | 0,2952 |
| C160 | D/V objetivo: 50% (máximo de R&R) a perpetuidad [Valor] | `Costo_Capital!C31` | 0,5000 |
| C161 | D/V objetivo: promedio comparables (elegida) [Patrimonio] | `Costo_Capital!F29` | 6.758,3437 |
| C162 | D/V objetivo: mediana comparables [Patrimonio] | `Costo_Capital!F30` | 6.787,5187 |
| C163 | D/V objetivo: 50% (máximo de R&R) a perpetuidad [Patrimonio] | `Costo_Capital!F31` | 7.262,5416 |
| C164 | Costo de la deuda antes de impuestos (Kd) | `Costo_Capital!C6` | 0,0550 |
| C165 | Tasa de impuestos | `Costo_Capital!C8` | 0,4000 |
| C166 | Tasa efectiva de impuestos [2007] | `Historicos!E131` | 0,3965 |
| C167 | Primas por tamaño, regulación o país [Patrimonio] | `Costo_Capital!F32` | 6.758,3437 |
| C168 | Afiliadas y minoritario: libros vs. P/U [Bajo] | `Sensibilidad!B30` | 5.440,9549 |
| C169 | Descuento [Ninguno (base)] | `Sensibilidad!B20` | 0,0000 |
| C170 | Descuento [Costo de salir a bolsa] | `Sensibilidad!C20` | 0,1000 |
| C171 | Descuento [Anexo 8 (lectura)] | `Sensibilidad!D20` | 0,1570 |
| C172 | Descuento [Regla práctica] | `Sensibilidad!E20` | 0,2500 |
| C173 | Descuento [Académico] | `Sensibilidad!F20` | 0,3500 |
| C174 | Patrimonio [Ninguno (base)] | `Sensibilidad!B21` | 6.758,3437 |
| C175 | Patrimonio [Costo de salir a bolsa] | `Sensibilidad!C21` | 6.082,5093 |
| C176 | Patrimonio [Anexo 8 (lectura)] | `Sensibilidad!D21` | 5.697,2838 |
| C177 | Patrimonio [Regla práctica] | `Sensibilidad!E21` | 5.068,7578 |
| C178 | Patrimonio [Académico] | `Sensibilidad!F21` | 4.392,9234 |
| C179 | Descuento por iliquidez aplicado al patrimonio [Valor] | `Supuestos!C29` | 0,0000 |
| C180 | Beta desapalancada (activos) según selector | `Costo_Capital!C9` | 0,7967 |
| C181 | Prima de riesgo de mercado | `Costo_Capital!C5` | 0,0500 |
| C182 | Estructura objetivo D/V (post 2012, valor terminal) | `Costo_Capital!C11` | 0,2812 |
| C183 | D/P objetivo = L / (1 − L) | `Costo_Capital!C12` | 0,3912 |
| C184 | Beta del patrimonio reapalancada | `Costo_Capital!C14` | 1,0107 |
| C185 | Control (solo Harris-Pringle): Ku − t × Kd × D/V | `Costo_Capital!C17` | 0,0762 |
| C186 | Rendimiento anual exigido sobre el valor de las operaciones = Ku × valor (APV) | `Valoracion!C63` | 502,5584 |
| C187 | Utilidad operativa después de impuestos (NOPAT) [2008] | `Proyeccion!D17` | 243,5466 |
| C188 | (sin etiqueta) | `Sensibilidad!D4` | 0,0250 |
| C189 | (sin etiqueta) | `Sensibilidad!A7` | 0,0800 |
| C190 | (sin etiqueta) | `Sensibilidad!A8` | 0,0825 |
| C191 | (sin etiqueta) | `Sensibilidad!D7` | 6.926,4937 |
| C192 | (sin etiqueta) | `Sensibilidad!D8` | 6.636,5086 |
| C193 | VP de los flujos libres 2008-2012 | `Valoracion!C8` | 1.258,4170 |
| C194 | Flujo de caja libre no apalancado [2013 normal.] | `Proyeccion!I28` | 327,5023 |
| C195 | VP del valor terminal (perpetuidad) | `Valoracion!C9` | 4.560,4438 |
| C196 | VP de los escudos fiscales de la deuda de adquisición 2008-2012 | `Valoracion!C10` | 284,7831 |
| C197 | Caja y valores negociables 2007 | `Valoracion!C15` | 220,8850 |
| C198 | Participación en utilidades de afiliadas 2007 | `Valoracion!C16` | 90,0330 |
| C199 | Valor de las afiliadas | `Valoracion!C17` | 1.719,1801 |
| C200 | Nota por cobrar de largo plazo | `Valoracion!C18` | 4,4220 |
| C201 | Deuda existente (largo plazo + pagaré), se refinancia | `Valoracion!C22` | 1.002,2930 |
| C202 | Interés minoritario en resultados 2007 | `Valoracion!C23` | 15,0560 |
| C203 | Valor del interés minoritario | `Valoracion!C24` | 287,4943 |
| C204 | ROIC (NOPAT / capital invertido al inicio del año) [2012] | `Proyeccion!H30` | 0,1042 |
| C205 | Tasa de reinversión = (capex − D&A + Δ capital de trabajo) / NOPAT [2012] | `Proyeccion!H31` | 0,2668 |
| C206 | Patrimonio por CPPC-DCF (sin sinergias) | `Valoracion!C49` | 6.627,0547 |
| C207 | Diferencia APV − CPPC-DCF (valor del endeudamiento inicial por encima del objetivo, neto) | `Valoracion!C50` | 131,2890 |
| C208 | Valor del patrimonio por P/U de la industria × utilidad neta 2007 | `Valoracion!C54` | 6.009,8457 |
| C209 | Valor del patrimonio por múltiplo VE/EBITDA 2007 actual (mediana) | `Valoracion!C53` | 11.332,4944 |
| C210 | Múltiplo VE/EBITDA 2012 implícito en el valor terminal por perpetuidad | `Valoracion!C44` | 4,0385 |
| C211 | Múltiplo VE/EBITDA usado | `Valoracion!C43` | 10,3334 |
| C212 | Promedio [Margen EBITDA] | `Comparables!M11` | 0,3159 |
| C213 | Margen EBITDA [2008] | `Proyeccion!D13` | 0,2464 |
| C214 | Mediana [P/U] | `Comparables!S13` | 30,9656 |
| C215 | Valor terminal: perpetuidad vs. múltiplo [Bajo] | `Sensibilidad!B29` | 6.758,3437 |
| C216 | Valor terminal: perpetuidad vs. múltiplo [Alto] | `Sensibilidad!C29` | 13.866,9095 |
| C217 | Valor terminal: perpetuidad vs. múltiplo [Rango] | `Sensibilidad!D29` | 7.108,5657 |
| C218 | Escenario combinado de riesgo [Bajo] | `Sensibilidad!B32` | 4.640,8903 |
| C219 | Escenario combinado de riesgo [Alto] | `Sensibilidad!C32` | 6.758,3437 |
| C220 | Escenario combinado de riesgo [Rango] | `Sensibilidad!D32` | 2.117,4534 |
| C221 | PRM 4% - 6% [Bajo] | `Sensibilidad!B27` | 6.008,3836 |
| C222 | PRM 4% - 6% [Alto] | `Sensibilidad!C27` | 7.773,0706 |
| C223 | PRM 4% - 6% [Rango] | `Sensibilidad!D27` | 1.764,6871 |
| C224 | Descuento por iliquidez 0% - 25% [Bajo] | `Sensibilidad!B31` | 5.068,7578 |
| C225 | Descuento por iliquidez 0% - 25% [Alto] | `Sensibilidad!C31` | 6.758,3437 |
| C226 | Descuento por iliquidez 0% - 25% [Rango] | `Sensibilidad!D31` | 1.689,5859 |
| C227 | Afiliadas y minoritario: libros vs. P/U [Alto] | `Sensibilidad!C30` | 6.758,3437 |
| C228 | Afiliadas y minoritario: libros vs. P/U [Rango] | `Sensibilidad!D30` | 1.317,3888 |
| C229 | D/V objetivo (mediana a 50%) [Bajo] | `Sensibilidad!B28` | 6.758,3437 |
| C230 | D/V objetivo (mediana a 50%) [Alto] | `Sensibilidad!C28` | 7.262,5416 |
| C231 | D/V objetivo (mediana a 50%) [Rango] | `Sensibilidad!D28` | 504,1979 |
| C232 | Método de beta (opciones 1 a 3) [Bajo] | `Sensibilidad!B26` | 6.747,8895 |
| C233 | Método de beta (opciones 1 a 3) [Alto] | `Sensibilidad!C26` | 7.208,3436 |
| C234 | Método de beta (opciones 1 a 3) [Rango] | `Sensibilidad!D26` | 460,4541 |
| C235 | g ± 0,5 pp [Bajo] | `Sensibilidad!B25` | 6.580,0463 |
| C236 | g ± 0,5 pp [Alto] | `Sensibilidad!C25` | 6.974,5318 |
| C237 | g ± 0,5 pp [Rango] | `Sensibilidad!D25` | 394,4855 |
| C238 | Indicador: deuda total / EBITDA [2007 proforma] | `Proyeccion!C89` | 3,6364 |
| C239 | Afiliadas / valor del patrimonio sin sinergias | `Valoracion!C64` | 0,2544 |
| C240 | Ahorro de backhaul (antes de impuestos) [2012] | `Sinergias!H8` | 75,9762 |
| C241 | VE/EBITDA 2007 implícito en el precio máximo (con deuda existente y minoritario) | `Valoracion!C67` | 6,5370 |
| C242 | Crecimiento de ingresos por servicios [2008] | `Supuestos!D36` | 0,1400 |
| C243 | Crecimiento de ingresos por servicios [2009] | `Supuestos!E36` | 0,1400 |
| C244 | Crecimiento de ingresos por servicios [2010] | `Supuestos!F36` | 0,1250 |
| C245 | Crecimiento de ingresos por servicios [2011] | `Supuestos!G36` | 0,1000 |
| C246 | Crecimiento de ingresos por servicios [2012] | `Supuestos!H36` | 0,0700 |
| C247 | Costos de operación del sistema / ingresos por servicios [2008] | `Supuestos!D37` | 0,2000 |
| C248 | Costo de equipos vendidos / ingresos por equipos [2008] | `Supuestos!D38` | 2,4000 |
| C249 | SG&A / ingresos totales [2008] | `Supuestos!D39` | 0,4000 |
| C250 | Capex / ingresos totales [2008] | `Supuestos!D41` | 0,1400 |
| C251 | Capex / ingresos totales [2012] | `Supuestos!H41` | 0,1550 |
| C252 | Días de proveedores (sobre costos operativos en efectivo) [Valor] | `Supuestos!C49` | 35,5365 |
| C253 | Días de pasivos acumulados (sobre costos operativos en efectivo) [Valor] | `Supuestos!C51` | 6,8451 |
| C254 | EBITDA 2007 | `Deuda!C3` | 1.033,3250 |
| C255 | Cuota mensual (amortización a 120 meses) | `Deuda!C6` | 40,7792 |
| C256 | Ingresos totales [2007 real] | `Proyeccion!C8` | 3.946,2640 |
| C257 | Ingresos totales [2008] | `Proyeccion!D8` | 4.509,1033 |
| C258 | Ingresos totales [2010] | `Proyeccion!F8` | 5.782,9250 |
| C259 | Ingresos totales [2012] | `Proyeccion!H8` | 6.806,5027 |
| C260 | EBITDA [2007 real] | `Proyeccion!C12` | 1.033,3250 |
| C261 | EBITDA [2008] | `Proyeccion!D12` | 1.111,1404 |
| C262 | EBITDA [2010] | `Proyeccion!F12` | 1.425,0376 |
| C263 | EBITDA [2012] | `Proyeccion!H12` | 1.677,2692 |
| C264 | Margen EBITDA [2007 real] | `Proyeccion!C13` | 0,2618 |
| C265 | Margen EBITDA [2010] | `Proyeccion!F13` | 0,2464 |
| C266 | Margen EBITDA [2012] | `Proyeccion!H13` | 0,2464 |
| C267 | Utilidad operativa después de impuestos (NOPAT) [2010] | `Proyeccion!F17` | 334,5593 |
| C268 | Utilidad operativa después de impuestos (NOPAT) [2012] | `Proyeccion!H17` | 434,6153 |
| C269 | Capex [2008] | `Proyeccion!D27` | 631,2745 |
| C270 | Capex [2010] | `Proyeccion!F27` | 867,4388 |
| C271 | Capex [2012] | `Proyeccion!H27` | 1.055,0079 |
| C272 | Flujo de caja libre no apalancado [2008] | `Proyeccion!D28` | 291,6232 |
| C273 | Flujo de caja libre no apalancado [2010] | `Proyeccion!F28` | 314,5390 |
| C274 | Flujo de caja libre no apalancado [2012] | `Proyeccion!H28` | 318,6437 |
| C275 | Gasto de intereses [2007] | `Historicos!E14` | 84,6790 |
| C276 | Intereses de la deuda de adquisición [2008] | `Proyeccion!D50` | 199,4290 |
| C277 | Intereses de la línea de crédito (Kd sobre saldo inicial) [2008] | `Proyeccion!D51` | 0,0000 |
| C278 | Intereses de la deuda de adquisición [2010] | `Proyeccion!F50` | 165,7987 |
| C279 | Intereses de la línea de crédito (Kd sobre saldo inicial) [2010] | `Proyeccion!F51` | 10,7628 |
| C280 | Intereses de la deuda de adquisición [2012] | `Proyeccion!H50` | 128,2675 |
| C281 | Intereses de la línea de crédito (Kd sobre saldo inicial) [2012] | `Proyeccion!H51` | 23,6417 |
| C282 | Utilidad neta (solo operaciones) [2008] | `Proyeccion!D54` | 123,8893 |
| C283 | Utilidad neta (solo operaciones) [2010] | `Proyeccion!F54` | 228,6224 |
| C284 | Utilidad neta (solo operaciones) [2012] | `Proyeccion!H54` | 343,4698 |
| C285 | Deuda de adquisición (2012: saldo a refinanciar) [2007 proforma] | `Proyeccion!C71` | 3.757,5455 |
| C286 | Deuda de adquisición (2012: saldo a refinanciar) [2008] | `Proyeccion!D71` | 3.467,6235 |
| C287 | Deuda de adquisición (2012: saldo a refinanciar) [2010] | `Proyeccion!F71` | 2.837,7955 |
| C288 | Deuda de adquisición (2012: saldo a refinanciar) [2012] | `Proyeccion!H71` | 2.134,9090 |
| C289 | Línea de crédito (variable de cierre) [2008] | `Proyeccion!D70` | 117,9561 |
| C290 | Línea de crédito (variable de cierre) [2010] | `Proyeccion!F70` | 310,6369 |
| C291 | Indicador: deuda total / EBITDA [2008] | `Proyeccion!D89` | 3,2269 |
| C292 | Indicador: deuda total / EBITDA [2010] | `Proyeccion!F89` | 2,2094 |
| C293 | Indicador: deuda total / EBITDA [2012] | `Proyeccion!H89` | 1,6088 |
| C294 | Indicador: EBITDA / intereses totales [2008] | `Proyeccion!D90` | 5,5716 |
| C295 | Indicador: EBITDA / intereses totales [2010] | `Proyeccion!F90` | 8,0711 |
| C296 | Indicador: EBITDA / intereses totales [2012] | `Proyeccion!H90` | 11,0413 |
| C297 | Indicador: margen EBIT [2007 proforma] | `Proyeccion!C86` | 0,1143 |
| C298 | Indicador: margen EBIT [2008] | `Proyeccion!D86` | 0,0900 |
| C299 | Indicador: margen EBIT [2010] | `Proyeccion!F86` | 0,0964 |
| C300 | Indicador: margen EBIT [2012] | `Proyeccion!H86` | 0,1064 |
| C301 | Balance proforma cuadra | `Control!G13` | 0,0000 |
| C302 | Flujo proforma concilia | `Control!G14` | 0,0000 |
| C303 | (sin etiqueta) | `Sensibilidad!B4` | 0,0150 |
| C304 | (sin etiqueta) | `Sensibilidad!E4` | 0,0300 |
| C305 | (sin etiqueta) | `Sensibilidad!G4` | 0,0400 |
| C306 | (sin etiqueta) | `Sensibilidad!A5` | 0,0750 |
| C307 | (sin etiqueta) | `Sensibilidad!B5` | 7.110,6906 |
| C308 | (sin etiqueta) | `Sensibilidad!D5` | 7.605,6952 |
| C309 | (sin etiqueta) | `Sensibilidad!E5` | 7.943,2338 |
| C310 | (sin etiqueta) | `Sensibilidad!G5` | 8.954,6241 |
| C311 | (sin etiqueta) | `Sensibilidad!B7` | 6.575,5605 |
| C312 | (sin etiqueta) | `Sensibilidad!E7` | 7.157,1735 |
| C313 | (sin etiqueta) | `Sensibilidad!G7` | 7.810,6032 |
| C314 | (sin etiqueta) | `Sensibilidad!B8` | 6.340,8342 |
| C315 | (sin etiqueta) | `Sensibilidad!E8` | 6.827,6828 |
| C316 | (sin etiqueta) | `Sensibilidad!G8` | 7.356,3757 |
| C317 | (sin etiqueta) | `Sensibilidad!A9` | 0,0850 |
| C318 | (sin etiqueta) | `Sensibilidad!B9` | 6.124,5680 |
| C319 | (sin etiqueta) | `Sensibilidad!D9` | 6.373,4540 |
| C320 | (sin etiqueta) | `Sensibilidad!E9` | 6.531,8750 |
| C321 | (sin etiqueta) | `Sensibilidad!G9` | 6.960,4440 |
| C322 | (sin etiqueta) | `Sensibilidad!A11` | 0,0900 |
| C323 | (sin etiqueta) | `Sensibilidad!B11` | 5.739,3638 |
| C324 | (sin etiqueta) | `Sensibilidad!D11` | 5.914,3998 |
| C325 | (sin etiqueta) | `Sensibilidad!E11` | 6.022,5341 |
| C326 | (sin etiqueta) | `Sensibilidad!G11` | 6.303,5627 |
| C327 | Ajuste al crecimiento de ingresos por servicios (pp por año) [Precios] | `Supuestos!D63` | −0,0300 |
| C328 | Ajuste a SG&A / ingresos totales (pp) [Adquisición] | `Supuestos!E64` | 0,0200 |
| C329 | EBITDA 2012 [Base] | `Escenarios!C181` | 1.677,2692 |
| C330 | EBITDA 2012 [Precios] | `Escenarios!D181` | 1.463,3266 |
| C331 | EBITDA 2012 [Adquisición] | `Escenarios!E181` | 1.541,1392 |
| C332 | EBITDA 2012 [Combinado] | `Escenarios!F181` | 1.344,5605 |
| C333 | Crecimiento a perpetuidad (g) [Base] | `Escenarios!C184` | 0,0278 |
| C334 | Crecimiento a perpetuidad (g) [Precios] | `Escenarios!D184` | 0,0232 |
| C335 | Crecimiento a perpetuidad (g) [Adquisición] | `Escenarios!E184` | 0,0277 |
| C336 | Crecimiento a perpetuidad (g) [Combinado] | `Escenarios!F184` | 0,0231 |
| C337 | Valor del patrimonio sin sinergias [Base] | `Escenarios!C186` | 6.758,3437 |
| C338 | Valor del patrimonio sin sinergias [Precios] | `Escenarios!D186` | 5.799,8555 |
| C339 | Valor del patrimonio sin sinergias [Adquisición] | `Escenarios!E186` | 5.337,0777 |
| C340 | Valor del patrimonio sin sinergias [Combinado] | `Escenarios!F186` | 4.640,8903 |
| C341 | Δ patrimonio vs. base [Precios] | `Escenarios!D187` | −958,4882 |
| C342 | Δ patrimonio vs. base [Adquisición] | `Escenarios!E187` | −1.421,2660 |
| C343 | Δ patrimonio vs. base [Combinado] | `Escenarios!F187` | −2.117,4534 |
| C344 | Patrimonio | `Sensibilidad!J21` | 8.963,3461 |
| C345 | Obligación por retiro de activos [2007] | `Historicos!E66` | 126,8440 |
| C346 | Todas coinciden (diferencia máxima en la columna D) [Diferencia] | `Verificacion!D40` | 0,0000 |
| C347 | Todos los controles | `Control!H18` | OK |
| C348 | Ingresos por servicios [2008] | `Proyeccion!D5` | 4.194,3302 |
| C349 | Depreciación y amortización [2008] | `Proyeccion!D14` | 705,2293 |
| C350 | Aumento del capital de trabajo operativo neto [2008] | `Proyeccion!D26` | 25,8783 |
| C351 | Saldo al cierre antes del bullet [2012] | `Deuda!P12` | 2.134,9090 |
