"""Construye el modelo de valoración de AirThread Connections (HBS 4263) en Excel.

Todas las cifras del caso se escriben una sola vez como insumos (texto azul) en
'Supuestos' o 'Historicos'; todo lo demás son fórmulas. Después de construirlo hay
que recalcular con LibreOffice (ver README de la carpeta).

Uso: python construir_modelo.py salida.xlsx
"""
import sys

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

FONT = "Arial"
AZUL = Font(name=FONT, size=10, color="0000FF")
NEGRO = Font(name=FONT, size=10, color="000000")
VERDE = Font(name=FONT, size=10, color="008000")
NEGRITA = Font(name=FONT, size=10, bold=True)
TITULO = Font(name=FONT, size=13, bold=True)
SUBT = Font(name=FONT, size=10, bold=True, color="FFFFFF")
FONDO_SUBT = PatternFill("solid", fgColor="1F3864")
AMARILLO = PatternFill("solid", fgColor="FFFF00")
GRIS = PatternFill("solid", fgColor="F2F2F2")
LINEA = Border(top=Side(style="thin"))

F_MM = '#,##0.0;(#,##0.0);"-"'
F_MM0 = '#,##0;(#,##0);"-"'
F_PCT = '0.0%;(0.0%);"-"'
F_PCT2 = '0.00%;(0.00%);"-"'
F_X = '0.00"x"'
F_X1 = '0.0"x"'
F_NUM = '0.000'
F_DIAS = '0.00'

# Referencias absolutas a celdas clave: R["clave"] = "'Hoja'!$C$5"
R = {}


def absref(ws, addr):
    col = "".join(c for c in addr if c.isalpha())
    row = "".join(c for c in addr if c.isdigit())
    return f"'{ws.title}'!${col}${row}"


class Hoja:
    def __init__(self, wb, titulo, encabezado, anchos):
        self.ws = wb.create_sheet(titulo)
        self.ws["A1"] = encabezado
        self.ws["A1"].font = TITULO
        for i, w in enumerate(anchos, start=1):
            self.ws.column_dimensions[get_column_letter(i)].width = w
        self.ws.sheet_view.showGridLines = False

    def c(self, addr, valor, fmt=None, tipo="f", negrita=False, clave=None, nota=None, fondo=None):
        """tipo: 'i' insumo (azul), 'f' fórmula (negro), 'v' vínculo a otra hoja (verde), 't' texto."""
        cel = self.ws[addr]
        cel.value = valor
        if tipo == "i":
            cel.font = Font(name=FONT, size=10, color="0000FF", bold=negrita)
        elif tipo == "v":
            cel.font = Font(name=FONT, size=10, color="008000", bold=negrita)
        else:
            cel.font = Font(name=FONT, size=10, bold=negrita)
        if fmt:
            cel.number_format = fmt
        if fondo is not None:
            cel.fill = fondo
        if nota:
            cel.comment = Comment(nota, "Modelo")
        if clave:
            R[clave] = absref(self.ws, addr)
        return cel

    def seccion(self, fila, texto, hasta="J"):
        for col in range(1, self.ws[f"{hasta}1"].column + 1):
            self.ws.cell(row=fila, column=col).fill = FONDO_SUBT
        cel = self.ws[f"A{fila}"]
        cel.value = texto
        cel.font = SUBT

    def texto(self, addr, t, negrita=False, italica=False):
        cel = self.ws[addr]
        cel.value = t
        cel.font = Font(name=FONT, size=10, bold=negrita, italic=italica)
        return cel


ANIOS_P = [2008, 2009, 2010, 2011, 2012]
COLP = ["D", "E", "F", "G", "H"]  # columnas de proyección 2008-2012
COL7 = "C"  # 2007 real en hojas de proyección
COLH = {2005: "C", 2006: "D", 2007: "E"}  # columnas históricas


def construir(ruta):
    wb = Workbook()
    wb.remove(wb.active)

    # ------------------------------------------------------------------ Supuestos
    S = Hoja(wb, "Supuestos", "Supuestos del modelo (millones de US$, salvo indicación)",
             [52, 4, 13, 12, 12, 12, 12, 12, 70])
    ws = S.ws
    ws["A2"] = ("Convención: azul = dato o supuesto escrito a mano; negro = fórmula; verde = vínculo a otra hoja; "
                "fondo amarillo = selector o supuesto clave que se puede cambiar.")
    ws["A2"].font = Font(name=FONT, size=9, italic=True)
    S.seccion(4, "1. Mercado, deuda e impuestos")
    for col, t in zip(["C", "I"], ["Valor", "Fuente / tipo"]):
        S.c(f"{col}5", t, tipo="t", negrita=True)
    filas = [
        ("Tasa de impuestos", 0.40, F_PCT, "t", "Caso, Anexo 1 (40% en todos los años proyectados)."),
        ("Costo de la deuda (Kd), BBB+", 0.055, F_PCT2, "kd", "Caso, p. 8: 'tasa de interés de aproximadamente 5,50%'."),
        ("Diferencial sobre el bono del Tesoro a 10 años", 0.0125, F_PCT2, "spread", "Caso, p. 8: '125 pb sobre el rendimiento actual del Tesoro a 10 años'."),
        ("Tasa libre de riesgo (Tesoro a 10 años) = Kd − diferencial", None, F_PCT2, "rf", "Derivada del caso (5,50% − 1,25%). El caso no da el plazo de 20 o 30 años."),
        ("Prima de riesgo de mercado (PRM)", 0.05, F_PCT2, "erp", "Caso, p. 8: recomendación de Bianco (R&R). No dice período ni si es promedio aritmético o geométrico."),
        ("Cobertura EBITDA / intereses para dimensionar la deuda", 5.0, F_X1, "cob", "Caso, p. 2 y p. 8 (≥ 5,0x sobre EBITDA 2007)."),
        ("Deuda / valor máximo admitido", 0.50, F_PCT, "dv_max", "Caso, p. 2 y p. 8 (45%-50%)."),
        ("Plazo de amortización de la deuda (meses)", 120, "0", "plazo", "Caso, p. 8 y Anexo 6."),
        ("Mes del pago bullet (refinanciación)", 60, "0", "mes_bullet", "Caso, p. 8: pago bullet al final del año 5."),
        ("Múltiplo P/U histórico de la industria", 19.095, F_X, "pe", "Caso, nota 3."),
        ("Costo de salir a bolsa (% del patrimonio), tope razonable del descuento", 0.10, F_PCT, "ipo", "Caso, nota 6."),
        ("Descuento por iliquidez, Anexo 8 (empresa rentable, ingresos US$1.000 MM)", 0.157, F_PCT, "iliq_ex8",
         "Lectura aproximada de un gráfico (imagen); el valor exacto no está en el caso."),
        ("Descuento por iliquidez, regla práctica (20%-30%), punto medio", 0.25, F_PCT, "iliq_rule", "Caso, p. 8."),
        ("Descuento por iliquidez, investigación académica", 0.35, F_PCT, "iliq_acad", "Caso, p. 8 y nota 5 (Moroney 35%; Silber 33,75%)."),
    ]
    fila = 6
    for etiqueta, valor, fmt, clave, fuente in filas:
        S.texto(f"A{fila}", etiqueta)
        if valor is None:
            S.c(f"C{fila}", "=C7-C8", fmt, clave=clave)
        else:
            S.c(f"C{fila}", valor, fmt, tipo="i", clave=clave)
        S.texto(f"I{fila}", fuente)
        fila += 1

    fila += 1
    S.seccion(fila, "2. Selectores de decisiones metodológicas (cambiar el número para ver el efecto)")
    fila += 1
    selectores = [
        ("Beta desapalancada: 1 = Hamada (βd = 0); 2 = Harris-Pringle (βd = 0); 3 = Harris-Pringle con βd implícita", 3, "sel_beta",
         "Elección propia: la deuda BBB+ paga 125 pb sobre el Tesoro, así que su beta no es cero (ver hoja Costo_Capital)."),
        ("Estructura objetivo D/V: 1 = promedio comparables; 2 = mediana comparables; 3 = máximo de R&R (50%)", 1, "sel_dv",
         "Caso, p. 8: la deuda se paga hasta niveles de la industria; nota del Anexo 7: promedio de 3 años."),
        ("Crecimiento a perpetuidad: 1 = fundamental (ROIC × tasa de reinversión 2012); 2 = manual", 1, "sel_g",
         "Caso, p. 8 y notas 7 y 8: g depende del retorno sobre el capital y de la tasa de reinversión."),
        ("Crecimiento manual (solo si el selector anterior = 2)", 0.03, "sel_g_manual",
         "Supuesto propio para sensibilidad; no sale del caso."),
        ("Múltiplo terminal VE/EBITDA: 1 = promedio comparables; 2 = mediana", 2, "sel_mult",
         "Elección propia: la mediana es menos sensible a los extremos (2,6x a 19,1x)."),
        ("Valor de afiliadas e interés minoritario: 1 = P/U de la industria; 2 = valor en libros", 1, "sel_nooper",
         "Caso, p. 6 y nota 3 (Jennifer propone múltiplos de mercado)."),
        ("Margen de las sinergias de ingresos: 1 = margen EBITDA proyectado de ATC; 2 = margen sin SG&A ni equipos", 1, "sel_margen_sin",
         "El caso (Tabla 5) solo da ingresos, no la utilidad bruta que menciona el texto. Supuesto propio."),
        ("Descuento por iliquidez aplicado al patrimonio", 0.0, "iliq",
         "Elección propia (ver informe): comprador estratégico del 100%, empresa grande y rentable con opción de salir a bolsa."),
        ("Porcentaje de sinergias que se realiza", 1.0, "realiz_sin", "Escenario base del caso (Tablas 4 y 5)."),
        ("Parte de las sinergias de backhaul que se cede al vendedor en el precio objetivo", 0.5, "cesion_bh",
         "Supuesto propio de negociación (reparto en partes iguales de la sinergia más cierta)."),
        ("Variación de g para el análisis tornado (± pp)", 0.005, "torn_dg", "Supuesto propio para sensibilidad."),
    ]
    for etiqueta, valor, clave, fuente in selectores:
        S.texto(f"A{fila}", etiqueta)
        fmt = F_PCT if isinstance(valor, float) else "0"
        S.c(f"C{fila}", valor, fmt, tipo="i", clave=clave, fondo=AMARILLO)
        S.texto(f"I{fila}", fuente)
        fila += 1

    fila += 1
    S.seccion(fila, "3. Supuestos operativos 2008-2012 (Anexo 1)")
    fila += 1
    S.texto(f"A{fila}", "Año", negrita=True)
    for col, a in zip(COLP, ANIOS_P):
        S.c(f"{col}{fila}", str(a), tipo="t", negrita=True)
    fila += 1
    anuales = [
        ("Crecimiento de ingresos por servicios", [0.14, 0.14, 0.125, 0.10, 0.07], F_PCT, "g_serv", "Caso, Anexo 1."),
        ("Costos de operación del sistema / ingresos por servicios", [0.20] * 5, F_PCT, "sysop_pct", "Caso, Anexo 1."),
        ("Costo de equipos vendidos / ingresos por equipos", [2.40] * 5, F_PCT, "cogs_pct", "Caso, Anexo 1 (240%: los equipos se subsidian)."),
        ("SG&A / ingresos totales", [0.40] * 5, F_PCT, "sga_pct", "Caso, Anexo 1."),
        ("Depreciación y amortización", [705.229319748866, 803.961424513707, 867.438753154676, 922.376540854473, 952.910384965518],
         F_MM, "dya", "Caso, Anexo 1 (valores absolutos)."),
        ("Capex / ingresos totales", [0.14, 0.14, 0.15, 0.1525, 0.155], F_PCT2, "capex_pct", "Caso, Anexo 1 (incluye licencias y listas de clientes)."),
        ("D&A / ingresos totales (implícito en el Anexo 1)", [None] * 5, F_PCT2, "dya_pct",
         "Derivado: D&A del Anexo 1 / ingresos del escenario base. En los escenarios la D&A se mueve con los ingresos."),
    ]
    for etiqueta, vals, fmt, clave, fuente in anuales:
        S.texto(f"A{fila}", etiqueta)
        for col, v in zip(COLP, vals):
            if v is not None:
                S.c(f"{col}{fila}", v, fmt, tipo="i")
        R[clave] = fila  # fila en Supuestos
        S.texto(f"I{fila}", fuente)
        fila += 1
    fila += 1
    S.texto(f"A{fila}", "Razones de capital de trabajo (promedio 2005-2007, año de 360 días)", negrita=True)
    fila += 1
    ct = [
        ("Ingresos por equipos / ingresos por servicios", "eq_ratio", F_PCT2, "Hist_eq_ratio",
         "Promedio histórico; coincide con el 7,5% del Anexo 1."),
        ("Días de cartera (sobre ingresos totales)", "dso", F_DIAS, "Hist_dso", "Promedio histórico; coincide con 41,67 del Anexo 1."),
        ("Días de inventario (sobre ingresos por equipos)", "dio", F_DIAS, "Hist_dio",
         "Promedio histórico; coincide con 154,36 del Anexo 1. Se calcula sobre ingresos de equipos, no sobre su costo."),
        ("Gastos prepagados / costos operativos en efectivo", "prep", F_PCT2, "Hist_prep",
         "Promedio histórico = 1,38% del Anexo 1. El anexo no dice la base; solo cuadra con costos operativos en efectivo."),
        ("Días de proveedores (sobre costos operativos en efectivo)", "dpo", F_DIAS, "Hist_dpo", "Promedio histórico; coincide con 35,54."),
        ("Días de ingresos diferidos (sobre ingresos por servicios)", "ddif", F_DIAS, "Hist_ddif",
         "Promedio histórico = 14,01. La nota del Anexo 1 dice 'costos operativos en efectivo', pero solo cuadra con ingresos por servicios."),
        ("Días de pasivos acumulados (sobre costos operativos en efectivo)", "dacum", F_DIAS, "Hist_dacum", "Promedio histórico; coincide con 6,85."),
    ]
    for etiqueta, clave, fmt, ref_hist, fuente in ct:
        S.texto(f"A{fila}", etiqueta)
        S.c(f"C{fila}", None, fmt, tipo="v", clave=clave)  # fórmula se completa al crear Historicos
        R["pend_" + clave] = (f"C{fila}", ref_hist)
        S.texto(f"I{fila}", fuente)
        fila += 1

    fila += 1
    S.seccion(fila, "4. Sinergias (Tablas 4 y 5)")
    fila += 1
    S.texto(f"A{fila}", "Año", negrita=True)
    for col, a in zip(COLP, ANIOS_P):
        S.c(f"{col}{fila}", str(a), tipo="t", negrita=True)
    fila += 1
    sin = [
        ("Costo de backhaul / costos de operación del sistema", [0.20] * 5, F_PCT, "bh_pct", "Caso, p. 6 y Tabla 4."),
        ("Reducción del costo de backhaul", [0.0, 0.07, 0.12, 0.222, 0.30], F_PCT, "bh_red", "Caso, Tabla 4."),
        ("Suscriptores empresariales promedio (millones)", [0.30, 0.50, 0.70, 1.00, 1.20], "0.00", "sus_emp", "Caso, Tabla 5."),
        ("Minutos promedio mensuales por suscriptor", [859, 885, 911, 939, 967], "#,##0", "min_emp", "Caso, Tabla 5."),
        ("Ingreso por minuto, clientes empresariales (US$)", [0.0506] * 5, "0.0000", "ipm_emp", "Caso, Tabla 5."),
    ]
    for etiqueta, vals, fmt, clave, fuente in sin:
        S.texto(f"A{fila}", etiqueta)
        for col, v in zip(COLP, vals):
            S.c(f"{col}{fila}", v, fmt, tipo="i")
        R[clave] = fila
        S.texto(f"I{fila}", fuente)
        fila += 1

    fila += 1
    S.seccion(fila, "5. Escenarios de riesgo (supuestos propios, motivados por el caso)")
    fila += 1
    for col, t in zip(["C", "D", "E", "F"], ["Base", "Precios", "Adquisición", "Combinado"]):
        S.c(f"{col}{fila}", t, tipo="t", negrita=True)
    R["esc_hdr"] = fila
    fila += 1
    esc = [
        ("Ajuste al crecimiento de ingresos por servicios (pp por año)", [0, -0.03, 0, -0.03], F_PCT, "esc_dg",
         "El ingreso por minuto cayó 11% en 2007 (Anexo 3). Supuesto propio: −3 pp de crecimiento cada año."),
        ("Ajuste a SG&A / ingresos totales (pp)", [0, 0, 0.02, 0.02], F_PCT, "esc_dsga",
         "El costo por cliente nuevo subió de US$385 a US$487 en 2007 (Anexo 3). Supuesto propio: +2 pp de SG&A."),
        ("Ajuste a capex / ingresos (pp)", [0, 0, 0, 0], F_PCT, "esc_dcapex", "Disponible para otros escenarios; en cero."),
    ]
    for etiqueta, vals, fmt, clave, fuente in esc:
        S.texto(f"A{fila}", etiqueta)
        for col, v in zip(["C", "D", "E", "F"], vals):
            S.c(f"{col}{fila}", v, fmt, tipo="i")
        R[clave] = fila
        S.texto(f"I{fila}", fuente)
        fila += 1

    def sup(clave, col):
        return f"Supuestos!${col}${R[clave]}"

    # ------------------------------------------------------------------ Historicos
    H = Hoja(wb, "Historicos", "Datos históricos 2005-2007, controles de consistencia e indicadores (millones de US$)",
             [50, 4, 12, 12, 12, 4, 14, 60])
    ws = H.ws
    H.seccion(3, "A. Estado de resultados (Anexo 4)")
    for col, a in COLH.items():
        H.c(f"{a}4", str(col), tipo="t", negrita=True)
    is_rows = [
        ("Ingresos por servicios", "S", [2827.022, 3214.41, 3679.237]),
        ("Ingresos por equipos", "EQ", [203.743, 258.745, 267.027]),
        ("Ingresos totales", "REV", "=C5+C6"),
        ("Costos de operación del sistema", "SYS", [604.093, 639.683, 717.075]),
        ("Costo de equipos vendidos", "COGS", [511.939, 568.903, 640.225]),
        ("Gastos de ventas, generales y de administración (SG&A)", "SGA", [1217.709, 1399.561, 1555.639]),
        ("EBITDA", "EBITDA", "=C7-C8-C9-C10"),
        ("Depreciación y amortización", "DA", [490.093, 555.525, 582.269]),
        ("EBIT (utilidad operativa)", "EBIT", "=C11-C12"),
        ("Gasto de intereses", "INT", [84.867, 93.674, 84.679]),
        ("Participación en utilidades de afiliadas", "AFIL", [66.719, 93.119, 90.033]),
        ("Ganancias (pérdidas) en inversiones", "GAN", [18.063, 50.84, 83.13]),
        ("Otros ingresos", "OTR", [54.501, -46.63, 6.961]),
        ("Utilidad antes de impuestos", "EBT", "=C13-C14+C15+C16+C17"),
        ("Impuestos", "TAX", [95.856, 120.604, 216.711]),
        ("Utilidad antes de interés minoritario", "IBMI", "=C18-C19"),
        ("Interés minoritario", "MI", [10.54, 13.044, 15.056]),
        ("Utilidad neta", "NI", "=C20-C21"),
    ]
    HR = {}
    fila = 5
    for etiqueta, k, vals in is_rows:
        H.texto(f"A{fila}", etiqueta, negrita=k in ("REV", "EBITDA", "EBIT", "NI"))
        for i, a in enumerate(["C", "D", "E"]):
            if isinstance(vals, list):
                H.c(f"{a}{fila}", vals[i], F_MM, tipo="i")
            else:
                H.c(f"{a}{fila}", vals.replace("C", a), F_MM, negrita=k in ("REV", "EBITDA", "EBIT", "NI"))
        HR[k] = fila
        fila += 1
    # control contra cifras publicadas (redondeadas) del PDF
    fila += 1
    H.texto(f"A{fila}", "Control: subtotales publicados en el Anexo 4 (PDF)", negrita=True)
    fila += 1
    pub_is = [("REV", [3030.8, 3473.2, 3946.3]), ("EBITDA", [697.0, 865.0, 1033.3]), ("EBIT", [206.9, 309.5, 451.1]),
              ("EBT", [261.3, 313.1, 546.5]), ("IBMI", [165.5, 192.5, 329.8]), ("NI", [155.0, 179.5, 314.7])]
    H.c(f"G{fila-1}", "Diferencia máx.", tipo="t", negrita=True)
    for k, vals in pub_is:
        H.texto(f"A{fila}", f"{ws['A'+str(HR[k])].value} publicado")
        for i, a in enumerate(["C", "D", "E"]):
            H.c(f"{a}{fila}", vals[i], F_MM, tipo="i")
        H.c(f"G{fila}", f"=MAX(ABS(C{fila}-C{HR[k]}),ABS(D{fila}-D{HR[k]}),ABS(E{fila}-E{HR[k]}))", F_MM)
        H.c(f"H{fila}", f'=IF(G{fila}<0.06,"Cuadra (redondeo)","NO CUADRA")')
        fila += 1
    R["chk_is"] = absref(ws, f"G{fila-1}")
    H.texto(f"A{fila}", "Diferencia máxima de todos los subtotales del estado de resultados")
    H.c(f"G{fila}", f"=MAX(G{fila-len(pub_is)}:G{fila-1})", F_MM, clave="chk_is_max")
    fila += 2

    H.seccion(fila, "B. Balance general (Anexo 5)")
    fila += 1
    for col, a in COLH.items():
        H.c(f"{a}{fila}", str(col), tipo="t", negrita=True)
    fila += 1
    bs = [
        ("Caja y equivalentes", "CASH", [29.003, 32.912, 204.533]),
        ("Valores negociables", "MKS", [0, 249.039, 16.352]),
        ("Cuentas por cobrar", "AR", [362.359, 407.438, 435.497]),
        ("Inventario", "INV", [92.748, 117.189, 100.99]),
        ("Gastos prepagados", "PREP", [32.068, 34.955, 41.588]),
        ("Impuestos diferidos (activo corriente)", "DTA", [8.218, 0, 18.566]),
        ("Otros activos corrientes", "OCA", [15.489, 13.385, 16.227]),
        ("Total activos corrientes", "TCA", None),
        ("Propiedad, planta y equipo", "PPE", [2553.029, 2628.848, 2595.096]),
        ("Licencias", "LIC", [1362.263, 1494.327, 1482.446]),
        ("Listas de clientes", "LIST", [47.649, 26.196, 15.375]),
        ("Títulos de renta variable negociables", "MES", [225.387, 4.873, 0]),
        ("Inversiones en entidades afiliadas", "INVAF", [172.093, 150.325, 157.693]),
        ("Nota por cobrar de largo plazo", "NOTE", [4.707, 4.541, 4.422]),
        ("Plusvalía (goodwill)", "GW", [481.235, 485.452, 491.316]),
        ("Otros activos de largo plazo", "OLTA", [29.985, 31.136, 31.773]),
        ("Total activos", "TA", None),
        ("Cuentas por pagar", "AP", [254.116, 254.871, 260.791]),
        ("Ingresos diferidos y depósitos", "DEF", [111.407, 123.344, 143.445]),
        ("Pasivos acumulados", "ACC", [42.865, 47.842, 59.224]),
        ("Impuestos por pagar", "TXP", [36.748, 26.913, 43.105]),
        ("Impuestos diferidos (pasivo corriente)", "DTLC", [0, 26.326, 0]),
        ("Pagaré (deuda de corto plazo)", "NP", [135, 35, 0]),
        ("Contrato a plazo (corriente)", "FWDC", [0, 159.856, 0]),
        ("Pasivo por derivados (corriente)", "DERC", [0, 88.84, 0]),
        ("Otros pasivos corrientes", "OCL", [82.585, 93.718, 97.678]),
        ("Total pasivos corrientes", "TCL", None),
        ("Deuda de largo plazo", "LTD", [1001.385, 1001.839, 1002.293]),
        ("Contratos a plazo (largo plazo)", "FWDL", [159.856, 0, 0]),
        ("Pasivo por derivados (largo plazo)", "DERL", [25.818, 0, 0]),
        ("Pasivo por impuestos diferidos", "DTL", [647.086, 601.535, 554.412]),
        ("Obligación por retiro de activos", "ARO", [90.224, 127.639, 126.844]),
        ("Otros pasivos diferidos", "ODL", [46.234, 62.914, 84.53]),
        ("Interés minoritario", "MINT", [41.871, 36.7, 43.396]),
        ("Capital: componente 1 (oculto en la fórmula del Excel del caso)", "CS1", [55.046, 55.046, 55.046]),
        ("Capital: componente 2 (oculto en la fórmula)", "CS2", [33.006, 33.006, 33.006]),
        ("Capital: componente 3, capital pagado adicional (oculto en la fórmula)", "CS3", [1286.964, 1290.829, 1316.042]),
        ("Acciones comunes y capital pagado (total publicado)", "CS", None),
        ("Utilidades retenidas: componente 1 (oculto en la fórmula)", "RE1", [1368.988, 1548.478, 1823.022]),
        ("Utilidades retenidas: componente 2 (oculto)", "RE2", [44.122, 80.382, 10.134]),
        ("Utilidades retenidas: componente 3 (oculto, se resta)", "RE3", [47.088, 14.462, 41.094]),
        ("Utilidades retenidas (total publicado)", "RE", None),
        ("Total pasivos y patrimonio", "TLE", None),
    ]
    for etiqueta, k, vals in bs:
        HR[k] = fila
        fila += 1
    for etiqueta, k, vals in bs:
        f_ = HR[k]
        tot = k in ("TCA", "TA", "TCL", "TLE", "CS", "RE")
        H.texto(f"A{f_}", etiqueta, negrita=tot)
        for i, a in enumerate(["C", "D", "E"]):
            if vals is not None:
                H.c(f"{a}{f_}", vals[i], F_MM, tipo="i")
            else:
                if k == "TCA":
                    fm = f"=SUM({a}{HR['CASH']}:{a}{HR['OCA']})"
                elif k == "TA":
                    fm = f"={a}{HR['TCA']}+SUM({a}{HR['PPE']}:{a}{HR['OLTA']})"
                elif k == "TCL":
                    fm = f"=SUM({a}{HR['AP']}:{a}{HR['OCL']})"
                elif k == "CS":
                    fm = f"={a}{HR['CS1']}+{a}{HR['CS2']}+{a}{HR['CS3']}"
                elif k == "RE":
                    fm = f"={a}{HR['RE1']}+{a}{HR['RE2']}-{a}{HR['RE3']}"
                else:  # TLE
                    fm = (f"={a}{HR['TCL']}+SUM({a}{HR['LTD']}:{a}{HR['MINT']})+{a}{HR['CS']}+{a}{HR['RE']}")
                H.c(f"{a}{f_}", fm, F_MM, negrita=True)
    fila += 1
    H.texto(f"A{fila}", "Control: activos − pasivos y patrimonio")
    for a in ["C", "D", "E"]:
        H.c(f"{a}{fila}", f"={a}{HR['TA']}-{a}{HR['TLE']}", F_MM)
    H.c(f"G{fila}", f"=MAX(ABS(C{fila}),ABS(D{fila}),ABS(E{fila}))", F_MM, clave="chk_bs")
    H.c(f"H{fila}", f'=IF(G{fila}<0.001,"El balance cuadra","NO CUADRA")')
    fila += 1
    pub_bs = [("TCA", [539.9, 854.9, 833.8]), ("TA", [5416.2, 5680.6, 5611.9]), ("TCL", [662.7, 856.7, 604.2]),
              ("CS", [1375.0, 1378.9, 1404.1]), ("RE", [1366.0, 1614.4, 1792.1]), ("TLE", [5416.2, 5680.6, 5611.9])]
    H.texto(f"A{fila}", "Control: totales publicados en el Anexo 5 (PDF)", negrita=True)
    H.c(f"G{fila}", "Diferencia máx.", tipo="t", negrita=True)
    fila += 1
    ini = fila
    for k, vals in pub_bs:
        H.texto(f"A{fila}", f"{ws['A'+str(HR[k])].value} – publicado")
        for i, a in enumerate(["C", "D", "E"]):
            H.c(f"{a}{fila}", vals[i], F_MM, tipo="i")
        H.c(f"G{fila}", f"=MAX(ABS(C{fila}-C{HR[k]}),ABS(D{fila}-D{HR[k]}),ABS(E{fila}-E{HR[k]}))", F_MM)
        H.c(f"H{fila}", f'=IF(G{fila}<0.06,"Cuadra (redondeo)","NO CUADRA")')
        fila += 1
    H.c(f"G{fila}", f"=MAX(G{ini}:G{fila-1})", F_MM, clave="chk_bs_pub")
    H.texto(f"A{fila}", "Diferencia máxima de los totales del balance")
    fila += 2

    H.seccion(fila, "C. Conciliaciones")
    fila += 1
    H.texto(f"A{fila}", "Patrimonio contable (capital + utilidades retenidas)")
    for a in ["C", "D", "E"]:
        H.c(f"{a}{fila}", f"={a}{HR['CS']}+{a}{HR['RE']}", F_MM)
    HR["PAT"] = fila
    fila += 1
    H.texto(f"A{fila}", "Variación de utilidades retenidas (componente 1)")
    for a, p in [("D", "C"), ("E", "D")]:
        H.c(f"{a}{fila}", f"={a}{HR['RE1']}-{p}{HR['RE1']}", F_MM)
    HR["dRE1"] = fila
    fila += 1
    H.texto(f"A{fila}", "Menos: utilidad neta del año")
    for a in ["D", "E"]:
        H.c(f"{a}{fila}", f"=-{a}{HR['NI']}", F_MM)
    fila += 1
    H.texto(f"A{fila}", "Diferencia = dividendos u otros cargos implícitos", negrita=True)
    for a in ["D", "E"]:
        H.c(f"{a}{fila}", f"={a}{fila-2}+{a}{fila-1}", F_MM, negrita=True)
    R["re_dif07"] = absref(ws, f"E{fila}")
    R["re_dif06"] = absref(ws, f"D{fila}")
    H.texto(f"H{fila}", "2006 cuadra; en 2007 faltan ~40 (dividendos u otro ajuste no informado).")
    fila += 1
    H.texto(f"A{fila}", "Variación de utilidades retenidas publicadas − utilidad neta")
    for a, p in [("D", "C"), ("E", "D")]:
        H.c(f"{a}{fila}", f"={a}{HR['RE']}-{p}{HR['RE']}-{a}{HR['NI']}", F_MM)
    R["re_pub_dif07"] = absref(ws, f"E{fila}")
    R["re_pub_dif06"] = absref(ws, f"D{fila}")
    H.texto(f"H{fila}", "La línea publicada mezcla utilidades retenidas con otros componentes (probablemente ORI).")
    fila += 2
    H.texto(f"A{fila}", "Activo fijo neto (PP&E + licencias + listas de clientes)")
    for a in ["C", "D", "E"]:
        H.c(f"{a}{fila}", f"={a}{HR['PPE']}+{a}{HR['LIC']}+{a}{HR['LIST']}", F_MM)
    HR["FA"] = fila
    fila += 1
    H.texto(f"A{fila}", "Capex implícito = Δ activo fijo neto + D&A")
    for a, p in [("D", "C"), ("E", "D")]:
        H.c(f"{a}{fila}", f"={a}{HR['FA']}-{p}{HR['FA']}+{a}{HR['DA']}", F_MM)
    HR["CAPEX"] = fila
    H.texto(f"H{fila}", "El caso no publica el capex histórico; se deduce del balance (supone sin ventas ni deterioros).")
    fila += 1
    H.texto(f"A{fila}", "Capex implícito / ingresos totales")
    for a in ["D", "E"]:
        H.c(f"{a}{fila}", f"={a}{HR['CAPEX']}/{a}{HR['REV']}", F_PCT)
    HR["CAPEXP"] = fila
    fila += 2

    # Indicadores
    H.seccion(fila, "D. Indicadores (módulo 2)")
    fila += 1
    for col, a in COLH.items():
        H.c(f"{a}{fila}", str(col), tipo="t", negrita=True)
    H.c(f"G{fila}", "Comparables", tipo="t", negrita=True)
    fila += 1
    IND = {}

    def ind(nombre, k, fm, fmt, desde="C", bench=None):
        nonlocal fila
        H.texto(f"A{fila}", nombre)
        for a in ["C", "D", "E"]:
            if a < desde:
                continue
            p = chr(ord(a) - 1)
            H.c(f"{a}{fila}", fm.replace("{a}", a).replace("{p}", p), fmt)
        if bench is not None:
            H.c(f"G{fila}", bench, fmt, tipo="v")
        IND[k] = fila
        R["ind_" + k] = fila
        fila += 1

    ind("Crecimiento de ingresos totales", "g_rev", f"={{a}}{HR['REV']}/{{p}}{HR['REV']}-1", F_PCT, "D")
    ind("Crecimiento de ingresos por servicios", "g_s", f"={{a}}{HR['S']}/{{p}}{HR['S']}-1", F_PCT, "D")
    ind("Margen EBITDA", "m_ebitda", f"={{a}}{HR['EBITDA']}/{{a}}{HR['REV']}", F_PCT, bench="=Comparables!M14")
    ind("Margen EBIT", "m_ebit", f"={{a}}{HR['EBIT']}/{{a}}{HR['REV']}", F_PCT, bench="=Comparables!L14")
    ind("Margen neto", "m_neto", f"={{a}}{HR['NI']}/{{a}}{HR['REV']}", F_PCT, bench="=Comparables!N14")
    ind("Margen de equipos (ingresos − costo) / ingresos por equipos", "m_eq", f"=({{a}}{HR['EQ']}-{{a}}{HR['COGS']})/{{a}}{HR['EQ']}", F_PCT)
    ind("SG&A / ingresos totales", "sga", f"={{a}}{HR['SGA']}/{{a}}{HR['REV']}", F_PCT)
    ind("ROA (utilidad neta / activos totales)", "roa", f"={{a}}{HR['NI']}/{{a}}{HR['TA']}", F_PCT)
    ind("ROE (utilidad neta / patrimonio al cierre)", "roe", f"={{a}}{HR['NI']}/{{a}}{HR['PAT']}", F_PCT)
    ind("ROIC (EBIT × (1 − t) / (patrimonio + deuda), definición de la nota 7)", "roic",
        f"={{a}}{HR['EBIT']}*(1-Supuestos!$C$6)/({{a}}{HR['PAT']}+{{a}}{HR['LTD']}+{{a}}{HR['NP']})", F_PCT)
    ind("Días de cartera (sobre ingresos totales, 360 días)", "dso", f"={{a}}{HR['AR']}/{{a}}{HR['REV']}*360", F_DIAS)
    ind("Días de inventario (sobre ingresos por equipos, convención del Anexo 1)", "dio", f"={{a}}{HR['INV']}/{{a}}{HR['EQ']}*360", F_DIAS)
    ind("Días de inventario (sobre costo de equipos)", "dio_c", f"={{a}}{HR['INV']}/{{a}}{HR['COGS']}*360", F_DIAS)
    ind("Costos operativos en efectivo (sistema + equipos + SG&A)", "co", f"={{a}}{HR['SYS']}+{{a}}{HR['COGS']}+{{a}}{HR['SGA']}", F_MM)
    ind("Gastos prepagados / costos operativos en efectivo", "prep", f"={{a}}{HR['PREP']}/{{a}}{fila-1}", F_PCT2)
    ind("Días de proveedores (sobre costos operativos en efectivo)", "dpo", f"={{a}}{HR['AP']}/{{a}}{IND['co']}*360", F_DIAS)
    ind("Días de ingresos diferidos (sobre ingresos por servicios)", "ddif", f"={{a}}{HR['DEF']}/{{a}}{HR['S']}*360", F_DIAS)
    ind("Días de pasivos acumulados (sobre costos operativos en efectivo)", "dacum", f"={{a}}{HR['ACC']}/{{a}}{IND['co']}*360", F_DIAS)
    ind("Ciclo de caja = cartera + inventario (costo) − proveedores", "ccc",
        f"={{a}}{IND['dso']}+{{a}}{IND['dio_c']}-{{a}}{IND['dpo']}", F_DIAS)
    ind("Ingresos por equipos / ingresos por servicios", "eq_ratio", f"={{a}}{HR['EQ']}/{{a}}{HR['S']}", F_PCT2)
    ind("Rotación de activos (ingresos / activos totales)", "rot", f"={{a}}{HR['REV']}/{{a}}{HR['TA']}", F_X)
    ind("Razón corriente", "rc", f"={{a}}{HR['TCA']}/{{a}}{HR['TCL']}", F_X)
    ind("Prueba ácida ((activo corriente − inventario) / pasivo corriente)", "pa",
        f"=({{a}}{HR['TCA']}-{{a}}{HR['INV']})/{{a}}{HR['TCL']}", F_X)
    ind("Deuda financiera / patrimonio contable", "de", f"=({{a}}{HR['LTD']}+{{a}}{HR['NP']})/{{a}}{HR['PAT']}", F_PCT,
        bench="=Comparables!I12")
    ind("Deuda financiera / EBITDA", "debitda", f"=({{a}}{HR['LTD']}+{{a}}{HR['NP']})/{{a}}{HR['EBITDA']}", F_X)
    ind("EBITDA / intereses", "cov", f"={{a}}{HR['EBITDA']}/{{a}}{HR['INT']}", F_X)
    ind("(Caja + valores negociables) / ingresos totales", "caja", f"=({{a}}{HR['CASH']}+{{a}}{HR['MKS']})/{{a}}{HR['REV']}", F_PCT)
    ind("Capex implícito / ingresos totales", "capex", f"={{a}}{HR['CAPEX']}/{{a}}{HR['REV']}", F_PCT, "D")
    ind("D&A / ingresos totales", "da", f"={{a}}{HR['DA']}/{{a}}{HR['REV']}", F_PCT)
    ind("Tasa efectiva de impuestos", "tax", f"={{a}}{HR['TAX']}/{{a}}{HR['EBT']}", F_PCT)
    # Datos de clientes (Anexo 3)
    fila += 1
    H.texto(f"A{fila}", "Datos de clientes (Anexo 3)", negrita=True)
    fila += 1
    a3 = [("Adiciones netas de clientes (miles)", "adds", [301, 310, 477], "#,##0"),
          ("Costo por cliente nuevo (US$)", "cpa", [372, 385, 487], "#,##0"),
          ("ARPU mensual (US$)", "arpu", [45.24, 47.23, 51.13], "0.00"),
          ("Minutos por cliente al mes", "mins", [625, 704, 859], "#,##0"),
          ("Tasa mensual de abandono (churn)", "churn", [0.021, 0.021, 0.017], F_PCT)]
    for etiqueta, k, vals, fmt in a3:
        H.texto(f"A{fila}", etiqueta)
        for i, a in enumerate(["C", "D", "E"]):
            H.c(f"{a}{fila}", vals[i], fmt, tipo="i")
        IND[k] = fila
        R["ind_" + k] = fila
        fila += 1
    ind("Costo total de clientes nuevos (miles de US$)", "cpa_tot", f"={{a}}{IND['adds']}*{{a}}{IND['cpa']}", "#,##0")
    ind("Ingreso por minuto (US$)", "rpm", f"={{a}}{IND['arpu']}/{{a}}{IND['mins']}", "0.0000")
    ind("Variación del ingreso por minuto", "g_rpm", f"={{a}}{IND['rpm']}/{{p}}{IND['rpm']}-1", F_PCT, "D")
    ind("Variación del costo por cliente nuevo", "g_cpa", f"={{a}}{IND['cpa']}/{{p}}{IND['cpa']}-1", F_PCT, "D")
    ind("Variación de los minutos por cliente", "g_mins", f"={{a}}{IND['mins']}/{{p}}{IND['mins']}-1", F_PCT, "D")
    ind("Variación del ARPU", "g_arpu", f"={{a}}{IND['arpu']}/{{p}}{IND['arpu']}-1", F_PCT, "D")
    H.texto(f"A{fila}", "Control: costo de clientes nuevos publicado (miles de US$)")
    for i, (a, v) in enumerate(zip(["C", "D", "E"], [111972, 119350, 232299])):
        H.c(f"{a}{fila}", v, "#,##0", tipo="i")
    H.c(f"G{fila}", f"=MAX(ABS(C{fila}-C{IND['cpa_tot']}),ABS(D{fila}-D{IND['cpa_tot']}),ABS(E{fila}-E{IND['cpa_tot']}))", "#,##0",
        clave="chk_a3")
    fila += 1
    H.texto(f"A{fila}", "Control: costo de equipos / ingresos por equipos publicado")
    for a, v in zip(["C", "D", "E"], [2.5126703739515, 2.19870142418211, 2.39760398761174]):
        H.c(f"{a}{fila}", v, F_PCT, tipo="i")
    H.c(f"G{fila}", f"=MAX(ABS(C{fila}-C{HR['COGS']}/C{HR['EQ']}),ABS(D{fila}-D{HR['COGS']}/D{HR['EQ']}),ABS(E{fila}-E{HR['COGS']}/E{HR['EQ']}))",
        "0.0000", clave="chk_a3b")
    fila += 1
    H.texto(f"A{fila}", "Tabla 2 del caso (publicada): rendimiento sobre activos operativos netos")
    for a, v in zip(["C", "D", "E"], [0.036, 0.05, 0.071]):
        H.c(f"{a}{fila}", v, F_PCT, tipo="i")
    IND["t2_rnoa"] = fila
    fila += 1
    H.texto(f"A{fila}", "Tabla 2: ROE publicado")
    for a, v in zip(["C", "D", "E"], [0.057, 0.06, 0.098]):
        H.c(f"{a}{fila}", v, F_PCT, tipo="i")
    H.c(f"G{fila}", f"=MAX(ABS(C{fila}-C{IND['roe']}),ABS(D{fila}-D{IND['roe']}),ABS(E{fila}-E{IND['roe']}))", F_PCT2, clave="chk_t2roe")
    fila += 1
    H.texto(f"A{fila}", "Tabla 2: rotación de activos publicada")
    for a, v in zip(["C", "D", "E"], [0.873, 0.944, 1.034]):
        H.c(f"{a}{fila}", v, F_PCT, tipo="i")
    IND["t2_rot"] = fila
    fila += 1
    H.texto(f"A{fila}", "Activos operativos netos implícitos en la Tabla 2 (ingresos / rotación)")
    for a in ["C", "D", "E"]:
        H.c(f"{a}{fila}", f"={a}{HR['REV']}/{a}{IND['t2_rot']}", F_MM)
    IND["t2_noa"] = fila
    R["ind_t2_noa"] = fila
    H.texto(f"H{fila}", "El caso no define 'activos operativos netos'; no se logró replicar la base exacta.")
    fila += 1
    H.texto(f"A{fila}", "EBIT × (1 − t) / activos operativos implícitos (debe igualar la Tabla 2)")
    for a in ["C", "D", "E"]:
        H.c(f"{a}{fila}", f"={a}{HR['EBIT']}*(1-Supuestos!$C$6)/{a}{IND['t2_noa']}", F_PCT)
    IND["t2_rnoa_calc"] = fila
    R["ind_t2_rnoa_calc"] = fila
    fila += 2

    # Anexo 2
    H.seccion(fila, "E. Mercados (Anexo 2, miles de personas)")
    fila += 1
    for col, t in zip(["C", "D", "E"], ["Población", "Clientes", "Penetración"]):
        H.c(f"{col}{fila}", t, tipo="t", negrita=True)
    fila += 1
    for nombre, datos in [("Mercados consolidados", [(65096, 3846), (11677, 1180), (2830, 518), (2287, 431), (481, 147)]),
                          ("Mercados en operación", [(32497, 3846), (7346, 1180), (2344, 518), (2287, 431), (481, 147)])]:
        ini = fila
        for reg, (p, c) in zip(["Central US", "Mid-Atlantic", "New England", "Northwest", "New York"], datos):
            H.texto(f"A{fila}", f"{nombre}: {reg}")
            H.c(f"C{fila}", p, "#,##0", tipo="i")
            H.c(f"D{fila}", c, "#,##0", tipo="i")
            H.c(f"E{fila}", f"=D{fila}/C{fila}", F_PCT)
            fila += 1
        H.texto(f"A{fila}", f"{nombre}: total", negrita=True)
        H.c(f"C{fila}", f"=SUM(C{ini}:C{fila-1})", "#,##0", negrita=True)
        H.c(f"D{fila}", f"=SUM(D{ini}:D{fila-1})", "#,##0", negrita=True)
        H.c(f"E{fila}", f"=D{fila}/C{fila}", F_PCT, negrita=True)
        R["pen_" + ("cons" if nombre.startswith("Mercados c") else "oper")] = absref(ws, f"E{fila}")
        fila += 1
    fila += 1

    # Flujo de caja histórico
    H.seccion(fila, "F. Flujo de caja histórico, método indirecto (módulo 3)")
    fila += 1
    for a, t in [("D", "2006"), ("E", "2007")]:
        H.c(f"{a}{fila}", t, tipo="t", negrita=True)
    fila += 1

    def dlt(k, signo=1):
        return {a: f"{'' if signo == 1 else '-'}({a}{HR[k]}-{p}{HR[k]})" for a, p in [("D", "C"), ("E", "D")]}

    cf_lineas = [
        ("Utilidad neta", "cf_ni", {a: f"{a}{HR['NI']}" for a in "DE"}),
        ("+ Depreciación y amortización", "cf_da", {a: f"{a}{HR['DA']}" for a in "DE"}),
        ("− Aumento de cartera, inventario, prepagados y otros activos corrientes", "cf_nwca",
         {a: f"-(({a}{HR['AR']}+{a}{HR['INV']}+{a}{HR['PREP']}+{a}{HR['OCA']})-({p}{HR['AR']}+{p}{HR['INV']}+{p}{HR['PREP']}+{p}{HR['OCA']}))"
          for a, p in [("D", "C"), ("E", "D")]}),
        ("+ Aumento de proveedores, ingresos diferidos, acumulados, impuestos por pagar y otros", "cf_nwcl",
         {a: f"(({a}{HR['AP']}+{a}{HR['DEF']}+{a}{HR['ACC']}+{a}{HR['TXP']}+{a}{HR['OCL']})-({p}{HR['AP']}+{p}{HR['DEF']}+{p}{HR['ACC']}+{p}{HR['TXP']}+{p}{HR['OCL']}))"
          for a, p in [("D", "C"), ("E", "D")]}),
        ("+ Variación neta de impuestos diferidos", "cf_dt",
         {a: f"(({a}{HR['DTLC']}+{a}{HR['DTL']}-{a}{HR['DTA']})-({p}{HR['DTLC']}+{p}{HR['DTL']}-{p}{HR['DTA']}))"
          for a, p in [("D", "C"), ("E", "D")]}),
        ("+ Aumento de retiro de activos y otros pasivos diferidos", "cf_aro",
         {a: f"(({a}{HR['ARO']}+{a}{HR['ODL']})-({p}{HR['ARO']}+{p}{HR['ODL']}))" for a, p in [("D", "C"), ("E", "D")]}),
    ]
    ini_cfo = fila
    for etiqueta, k, fms in cf_lineas:
        H.texto(f"A{fila}", etiqueta)
        for a in "DE":
            H.c(f"{a}{fila}", "=" + fms[a], F_MM)
        R[k] = fila
        fila += 1
    H.texto(f"A{fila}", "Flujo de caja operativo", negrita=True)
    for a in "DE":
        H.c(f"{a}{fila}", f"=SUM({a}{ini_cfo}:{a}{fila-1})", F_MM, negrita=True)
    R["cf_cfo"] = fila
    fila += 1
    ini_cfi = fila
    cfi = [
        ("− Capex implícito (activo fijo, licencias y listas de clientes)", {a: f"-{a}{HR['CAPEX']}" for a in "DE"}),
        ("− Aumento de plusvalía (adquisiciones)", dlt("GW", -1)),
        ("− Compra neta de valores negociables y títulos de renta variable",
         {a: f"-(({a}{HR['MKS']}+{a}{HR['MES']})-({p}{HR['MKS']}+{p}{HR['MES']}))" for a, p in [("D", "C"), ("E", "D")]}),
        ("− Aumento de inversiones en afiliadas, nota por cobrar y otros activos LP",
         {a: f"-(({a}{HR['INVAF']}+{a}{HR['NOTE']}+{a}{HR['OLTA']})-({p}{HR['INVAF']}+{p}{HR['NOTE']}+{p}{HR['OLTA']}))"
          for a, p in [("D", "C"), ("E", "D")]}),
    ]
    for etiqueta, fms in cfi:
        H.texto(f"A{fila}", etiqueta)
        for a in "DE":
            H.c(f"{a}{fila}", "=" + fms[a], F_MM)
        fila += 1
    H.texto(f"A{fila}", "Flujo de caja de inversión", negrita=True)
    for a in "DE":
        H.c(f"{a}{fila}", f"=SUM({a}{ini_cfi}:{a}{fila-1})", F_MM, negrita=True)
    R["cf_cfi"] = fila
    fila += 1
    ini_cff = fila
    cff = [
        ("+ Variación de deuda (pagaré + largo plazo)",
         {a: f"(({a}{HR['NP']}+{a}{HR['LTD']})-({p}{HR['NP']}+{p}{HR['LTD']}))" for a, p in [("D", "C"), ("E", "D")]}),
        ("+ Variación de contratos a plazo y derivados",
         {a: f"(({a}{HR['FWDC']}+{a}{HR['DERC']}+{a}{HR['FWDL']}+{a}{HR['DERL']})-({p}{HR['FWDC']}+{p}{HR['DERC']}+{p}{HR['FWDL']}+{p}{HR['DERL']}))"
          for a, p in [("D", "C"), ("E", "D")]}),
        ("+ Variación de interés minoritario", dlt("MINT")),
        ("+ Aportes de capital", dlt("CS")),
        ("+ Otros movimientos patrimoniales (Δ utilidades retenidas − utilidad neta)",
         {a: f"(({a}{HR['RE']}-{p}{HR['RE']})-{a}{HR['NI']})" for a, p in [("D", "C"), ("E", "D")]}),
    ]
    for etiqueta, fms in cff:
        H.texto(f"A{fila}", etiqueta)
        for a in "DE":
            H.c(f"{a}{fila}", "=" + fms[a], F_MM)
        fila += 1
    H.texto(f"A{fila}", "Flujo de caja de financiación", negrita=True)
    for a in "DE":
        H.c(f"{a}{fila}", f"=SUM({a}{ini_cff}:{a}{fila-1})", F_MM, negrita=True)
    R["cf_cff"] = fila
    fila += 1
    H.texto(f"A{fila}", "Variación de caja según flujo", negrita=True)
    for a in "DE":
        H.c(f"{a}{fila}", f"={a}{R['cf_cfo']}+{a}{R['cf_cfi']}+{a}{R['cf_cff']}", F_MM, negrita=True)
    R["cf_dcash"] = fila
    fila += 1
    H.texto(f"A{fila}", "Variación de caja según balance")
    for a, p in [("D", "C"), ("E", "D")]:
        H.c(f"{a}{fila}", f"={a}{HR['CASH']}-{p}{HR['CASH']}", F_MM)
    fila += 1
    H.texto(f"A{fila}", "Control: diferencia")
    for a in "DE":
        H.c(f"{a}{fila}", f"={a}{fila-2}-{a}{fila-1}", F_MM)
    H.c(f"G{fila}", f"=MAX(ABS(D{fila}),ABS(E{fila}))", F_MM, clave="chk_cf_hist")
    fila += 1
    H.texto(f"A{fila}", "Flujo de caja libre operativo = EBIT×(1−t) + D&A − capex − Δ capital de trabajo operativo", negrita=True)
    for a, p in [("D", "C"), ("E", "D")]:
        nwc_a = f"({a}{HR['AR']}+{a}{HR['INV']}+{a}{HR['PREP']}-{a}{HR['AP']}-{a}{HR['DEF']}-{a}{HR['ACC']})"
        nwc_p = f"({p}{HR['AR']}+{p}{HR['INV']}+{p}{HR['PREP']}-{p}{HR['AP']}-{p}{HR['DEF']}-{p}{HR['ACC']})"
        H.c(f"{a}{fila}", f"={a}{HR['EBIT']}*(1-Supuestos!$C$6)+{a}{HR['DA']}-{a}{HR['CAPEX']}-({nwc_a}-{nwc_p})", F_MM, negrita=True)
    R["cf_fcf"] = fila
    fila += 1
    H.texto(f"A{fila}", "Capital de trabajo operativo neto (CxC + inventario + prepagados − CxP − diferidos − acumulados)")
    for a in "CDE":
        H.c(f"{a}{fila}", f"={a}{HR['AR']}+{a}{HR['INV']}+{a}{HR['PREP']}-{a}{HR['AP']}-{a}{HR['DEF']}-{a}{HR['ACC']}", F_MM)
    HR["NWC"] = fila
    fila += 1

    # Promedios para Supuestos
    H.texto(f"A{fila}", "Promedios 2005-2007 usados en la proyección", negrita=True)
    fila += 1
    prom = {"Hist_eq_ratio": IND["eq_ratio"], "Hist_dso": IND["dso"], "Hist_dio": IND["dio"], "Hist_prep": IND["prep"],
            "Hist_dpo": IND["dpo"], "Hist_ddif": IND["ddif"], "Hist_dacum": IND["dacum"]}
    for k, r in prom.items():
        H.texto(f"A{fila}", ws[f"A{r}"].value)
        H.c(f"E{fila}", f"=AVERAGE(C{r}:E{r})", F_NUM)
        R[k] = absref(ws, f"E{fila}")
        fila += 1
    for k in ["eq_ratio", "dso", "dio", "prep", "dpo", "ddif", "dacum"]:
        celda, ref_hist = R["pend_" + k]
        S.ws[celda].value = f"={R[ref_hist]}"
    for k, v in HR.items():
        R["H_" + k] = v

    # ------------------------------------------------------------------ Comparables
    C = Hoja(wb, "Comparables", "Comparables inalámbricos (Anexo 7) y Tabla 3", [30, 12, 11, 9, 9, 8, 10, 9, 9, 9, 10, 9, 9, 9, 10, 10, 10, 9, 9, 9, 9])
    ws = C.ws
    ws["A2"] = "Unidades del Anexo 7 rotuladas como miles de US$; por magnitud parecen millones. Las razones no dependen de la unidad."
    ws["A2"].font = Font(name=FONT, size=9, italic=True)
    hdr = ["Empresa", "Patrimonio de mercado", "Deuda neta", "D/V publicado", "D/P publicado", "Beta patrimonio", "Ingresos", "EBIT",
           "EBITDA", "Utilidad neta", "D/V calculado", "Margen EBIT", "Margen EBITDA", "Margen neto", "Valor empresa (VE)", "VE/EBITDA",
           "VE/EBIT", "VE/Ingresos", "P/U", "D/P calculado"]
    for i, t in enumerate(hdr):
        C.c(f"{get_column_letter(i+1)}4", t, tipo="t", negrita=True)
        ws[f"{get_column_letter(i+1)}4"].alignment = Alignment(wrap_text=True)
    comps = [
        ("Universal Mobile", 118496.867, 69130, 0.368444035256422, 0.583390951593682, 0.86, 43882, 11795, 16949, 3794),
        ("Neuberger Wireless", 189469.94865, 79351.251, 0.295182266515118, 0.418806526129283, 0.89, 42684, 7020, 14099, 4102.8),
        ("Agile Connections", 21078.873, 5080.008393, 0.1941982272361, 0.241, 1.17, 34698, 1631.22018518518, 9914, -29.7420000000006),
        ("Big Country Communications", 26285.354631, 8335.0859534901, 0.240756206818009, 0.3171, 0.97, 38896.458, 6701.8597134, 12614.1213294, 3383.991846),
        ("Rocky Mountain Wireless", 7359.89929668, 3267.79528772592, 0.307479224376731, 0.443999999999999, 1.13, 4063.519, 510.05290488, 1028.4766589, 239.747621),
    ]
    for i, comp in enumerate(comps):
        r = 5 + i
        C.texto(f"A{r}", comp[0])
        fmts = [F_MM0, F_MM0, F_PCT, F_PCT, "0.00", F_MM0, F_MM0, F_MM0, F_MM0]
        for j, v in enumerate(comp[1:]):
            C.c(f"{get_column_letter(j+2)}{r}", v, fmts[j], tipo="i")
        C.c(f"K{r}", f"=C{r}/(B{r}+C{r})", F_PCT)
        C.c(f"L{r}", f"=H{r}/G{r}", F_PCT)
        C.c(f"M{r}", f"=I{r}/G{r}", F_PCT)
        C.c(f"N{r}", f"=J{r}/G{r}", F_PCT)
        C.c(f"O{r}", f"=B{r}+C{r}", F_MM0)
        C.c(f"P{r}", f"=O{r}/I{r}", F_X1)
        C.c(f"Q{r}", f"=O{r}/H{r}", F_X1)
        C.c(f"R{r}", f"=O{r}/G{r}", F_X)
        C.c(f"S{r}", f'=IF(J{r}>0,B{r}/J{r},"n.s.")', F_X1)
        C.c(f"T{r}", f"=C{r}/B{r}", F_PCT)
    C.texto("A11", "Promedio", negrita=True)
    C.texto("A12", "Promedio (D/P) y mediana", negrita=True)
    for col in "DEFKLMNPQRST":
        C.c(f"{col}11", f"=AVERAGE({col}5:{col}9)", ws[f"{col}5"].number_format, negrita=True)
        C.c(f"{col}13", f"=MEDIAN({col}5:{col}9)", ws[f"{col}5"].number_format)
    C.texto("A13", "Mediana")
    C.c("I12", "=E11", F_PCT)
    C.texto("A12", "D/P promedio publicado (referencia para indicadores)")
    C.texto("A14", "Tabla 3 publicada (promedio de márgenes)")
    C.c("L14", 0.156, F_PCT, tipo="i")
    C.c("M14", 0.316, F_PCT, tipo="i")
    C.c("N14", 0.066, F_PCT, tipo="i")
    C.texto("A15", "Control: diferencia Tabla 3 vs. recalculado")
    C.c("L15", "=L14-L11", F_PCT2)
    C.c("M15", "=M14-M11", F_PCT2)
    C.c("N15", "=N14-N11", F_PCT2)
    C.c("O15", "=MAX(ABS(L15),ABS(M15),ABS(N15))", F_PCT2, clave="chk_t3")
    C.texto("A16", "Control: D/V publicado − calculado (máx.)")
    C.c("O16", "=MAX(ABS(D5-K5),ABS(D6-K6),ABS(D7-K7),ABS(D8-K8),ABS(D9-K9))", F_PCT2, clave="chk_dv")
    C.texto("A17", "Inconsistencia: D/P promedio publicado vs. D/P implícito en el D/V promedio")
    C.c("D17", "=D11/(1-D11)", F_PCT, clave="de_implicito")
    C.c("E17", "=E11", F_PCT)
    R["dv_prom"] = "Comparables!$D$11"
    R["dv_med"] = "Comparables!$D$13"
    R["de_prom"] = "Comparables!$E$11"
    R["beta_prom"] = "Comparables!$F$11"
    R["evebitda_prom"] = "Comparables!$P$11"
    R["evebitda_med"] = "Comparables!$P$13"
    R["pe_med"] = "Comparables!$S$13"
    # Tabla 3 de AirThread
    C.texto("A18", "AirThread 2007 (Tabla 3: EBIT 11,4%, EBITDA 26,2%, neto 8,0%)")
    C.c("L18", f"=Historicos!E{IND['m_ebit']}", F_PCT, tipo="v")
    C.c("M18", f"=Historicos!E{IND['m_ebitda']}", F_PCT, tipo="v")
    C.c("N18", f"=Historicos!E{IND['m_neto']}", F_PCT, tipo="v")

    # Betas desapalancadas
    C.texto("A21", "Betas desapalancadas por empresa", negrita=True)
    for col, t in zip("BCDE", ["E/V", "1 Hamada (βd=0)", "2 Harris-Pringle (βd=0)", "3 Harris-Pringle (βd>0)"]):
        C.c(f"{col}22", t, tipo="t", negrita=True)
        ws[f"{col}22"].alignment = Alignment(wrap_text=True)
    for i in range(5):
        r = 23 + i
        C.c(f"A{r}", f"=A{5+i}")
        C.c(f"B{r}", f"=1-K{5+i}", F_PCT)
        C.c(f"C{r}", f"=F{5+i}/(1+(1-Supuestos!$C$6)*T{5+i})", F_NUM)
        C.c(f"D{r}", f"=F{5+i}*B{r}", F_NUM)
        C.c(f"E{r}", f"=F{5+i}*B{r}+Costo_Capital!$C$7*K{5+i}", F_NUM)
    C.texto("A28", "Promedio", negrita=True)
    C.texto("A29", "Mediana")
    for col in "CDE":
        C.c(f"{col}28", f"=AVERAGE({col}23:{col}27)", F_NUM, negrita=True)
        C.c(f"{col}29", f"=MEDIAN({col}23:{col}27)", F_NUM)
    C.c("B28", "=AVERAGE(B23:B27)", F_PCT)
    C.texto("A30", "Nota: se usan D/V y D/P calculados con deuda neta, como en el anexo; el anexo aclara que reflejan promedios de 3 años.")
    ws["A30"].font = Font(name=FONT, size=9, italic=True)

    # ------------------------------------------------------------------ Costo de capital
    K = Hoja(wb, "Costo_Capital", "Costo de capital: insumos, opciones y elección", [58, 4, 12, 12, 12, 12, 12, 70])
    ws = K.ws
    K.seccion(3, "A. Insumos elegidos")
    k_rows = [
        ("Tasa libre de riesgo (Tesoro 10 años)", f"={R['rf']}", F_PCT2, "k_rf", "v"),
        ("Prima de riesgo de mercado", f"={R['erp']}", F_PCT2, "k_erp", "v"),
        ("Costo de la deuda antes de impuestos (Kd)", f"={R['kd']}", F_PCT2, "k_kd", "v"),
        ("Beta de la deuda implícita en el CAPM = (Kd − rf) / PRM", "=(C6-C4)/C5", F_NUM, "k_bd", "f"),
        ("Tasa de impuestos", f"={R['t']}", F_PCT, "k_t", "v"),
        ("Beta desapalancada (activos) según selector", None, F_NUM, "k_bu", "f"),
        ("Costo de capital no apalancado Ku = rf + βu × PRM", "=C4+C9*C5", F_PCT2, "ku", "f"),
        ("Estructura objetivo D/V (post 2012, valor terminal)", None, F_PCT, "L", "f"),
        ("D/P objetivo = L / (1 − L)", "=C11/(1-C11)", F_PCT, "k_de", "f"),
        ("Beta de la deuda usada al reapalancar", f"=IF({R['sel_beta']}=3,C7,0)", F_NUM, "k_bd_used", "f"),
        ("Beta del patrimonio reapalancada", None, F_NUM, "k_be", "f"),
        ("Costo del patrimonio Ke = rf + βe × PRM", "=C4+C14*C5", F_PCT2, "ke", "f"),
        ("CPPC (WACC) objetivo = E/V × Ke + D/V × Kd × (1 − t)", "=(1-C11)*C15+C11*C6*(1-C8)", F_PCT2, "wacc", "f"),
        ("Control (solo Harris-Pringle): Ku − t × Kd × D/V", "=C10-C8*C6*C11", F_PCT2, "wacc_hp", "f"),
        ("Prima por tamaño", 0, F_PCT2, "k_size", "i"),
        ("Prima por riesgo regulatorio", 0, F_PCT2, "k_reg", "i"),
        ("Prima por riesgo país", 0, F_PCT2, "k_pais", "i"),
    ]
    for i, (etq, fm, fmt, clave, tipo) in enumerate(k_rows):
        r = 4 + i
        K.texto(f"A{r}", etq)
        K.c(f"C{r}", fm, fmt, tipo=tipo, clave=clave)
    # selector-dependent formulas (selector rows in Supuestos: beta=19, dv=20)
    ws["C9"] = f"=CHOOSE({R['sel_beta']},Comparables!C28,Comparables!D28,Comparables!E28)"
    ws["C11"] = f"=CHOOSE({R['sel_dv']},{R['dv_prom']},{R['dv_med']},{R['dv_max']})"
    ws["C14"] = f"=IF({R['sel_beta']}=1,C9*(1+(1-C8)*C12),C9+(C9-C13)*C12)"
    K.texto("H18", "El caso no da primas por tamaño, regulación ni país: se dejan en 0 (ver justificación en la sección B).")
    K.texto("H7", "Kd = rf + βd × PRM ⇒ βd = 0,25. La deuda BBB+ tiene riesgo sistemático; ignorarlo subestima Ku.")

    # B. Tabla de opciones con efecto en valor (fórmula cerrada del valor del patrimonio)
    K.seccion(22, "B. Opciones defendibles por insumo y su efecto en el valor del patrimonio (millones de US$)")
    for col, t in zip("ABCDEFG", ["Insumo / opción", "", "Valor", "Ku", "CPPC objetivo", "Patrimonio", "Δ vs. base"]):
        K.c(f"{col}23", t, tipo="t", negrita=True)
    K.texto("H23", "Fuente / comentario", negrita=True)
    R["k_opt_start"] = 24

    # ------------------------------------------------------------------ Deuda
    D = Hoja(wb, "Deuda", "Deuda de adquisición: préstamo amortizable a 10 años con pago bullet al mes 60 (Anexo 6)",
             [8, 12, 13, 11, 11, 11, 12, 13, 8, 4, 30, 12, 12, 12, 12, 12])
    ws = D.ws
    D.texto("A3", "EBITDA 2007")
    D.c("C3", f"=Historicos!E{HR['EBITDA']}", F_MM, tipo="v", clave="d_ebitda07")
    D.texto("A4", "Intereses máximos = EBITDA / cobertura")
    D.c("C4", f"=C3/{R['cob']}", F_MM, clave="d_intmax")
    D.texto("A5", "Deuda inicial = intereses máximos / Kd")
    D.c("C5", f"=C4/{R['kd']}", F_MM, clave="d0", negrita=True)
    D.texto("A6", "Cuota mensual (amortización a 120 meses)")
    D.c("C6", f"=PMT({R['kd']}/12,{R['plazo']},-C5)", F_MM, clave="cuota")
    D.texto("E3", "Anexo 6 publicado: deuda inicial 3.758, cuota 41, saldo antes del bullet 2.166.")
    for col, t in zip("ABCDEFGHI", ["Mes", "Fecha", "Saldo inicial", "Cuota", "Interés", "Capital", "Pago bullet", "Saldo final", "Año"]):
        D.c(f"{col}8", t, tipo="t", negrita=True)
    for m in range(1, 61):
        r = 8 + m
        D.c(f"A{r}", m, "0")
        D.c(f"B{r}", f"=DATE(2008,A{r}+1,0)", "dd/mm/yyyy")
        D.c(f"C{r}", "=C5" if m == 1 else f"=H{r-1}", F_MM)
        D.c(f"D{r}", "=$C$6", F_MM)
        D.c(f"E{r}", f"=C{r}*{R['kd']}/12", F_MM)
        D.c(f"F{r}", f"=D{r}-E{r}", F_MM)
        D.c(f"G{r}", f"=IF(A{r}={R['mes_bullet']},C{r}-F{r},0)", F_MM)
        D.c(f"H{r}", f"=C{r}-F{r}-G{r}", F_MM)
        D.c(f"I{r}", f"=YEAR(B{r})", "0")
    # Resumen anual
    D.texto("K8", "Resumen anual", negrita=True)
    for col, a in zip("LMNOP", ANIOS_P):
        D.c(f"{col}8", str(a), tipo="t", negrita=True)
    D.texto("K9", "Intereses")
    D.texto("K10", "Amortización ordinaria de capital")
    D.texto("K11", "Pago bullet (refinanciado)")
    D.texto("K12", "Saldo al cierre antes del bullet")
    D.texto("K13", "Escudo fiscal = intereses × t")
    D.texto("K14", "Saldo publicado en el Anexo 6")
    D.texto("K15", "Control: diferencia (fila 16 − fila 14)")
    D.texto("K16", "Saldo comparable con el Anexo 6 (2012: inicio del mes 60)")
    for col, a, pub in zip("LMNOP", ANIOS_P, [3468, 3161, 2838, 2496, 2166]):
        D.c(f"{col}9", f"=SUMIF($I$9:$I$68,{a},$E$9:$E$68)", F_MM)
        D.c(f"{col}10", f"=SUMIF($I$9:$I$68,{a},$F$9:$F$68)", F_MM)
        D.c(f"{col}11", f"=SUMIF($I$9:$I$68,{a},$G$9:$G$68)", F_MM)
        D.c(f"{col}12", f"=INDEX($C$9:$C$68,{(a-2007)*12})-INDEX($F$9:$F$68,{(a-2007)*12})", F_MM)
        # Anexo 6 muestra el saldo de fin de año; en 2012 muestra el saldo al inicio del mes 60 (antes de la última cuota)
        D.c(f"{col}16", f"=INDEX($H$9:$H$68,{(a-2007)*12})" if a < 2012 else "=INDEX($C$9:$C$68,60)", F_MM)
        D.c(f"{col}13", f"={col}9*{R['t']}", F_MM)
        D.c(f"{col}14", pub, F_MM0, tipo="i")
        D.c(f"{col}15", f"={col}16-{col}14", F_MM)
    D.c("Q15", "=MAX(ABS(L15),ABS(M15),ABS(N15),ABS(O15),ABS(P15))", F_MM, clave="chk_deuda")
    D.texto("K17", "Cobertura EBITDA 2007 / intereses del año 1")
    D.c("L17", "=C3/L9", F_X)
    D.texto("K18", "Nota: el Anexo 6 rotula 84 períodos restantes en el mes 41 (deberían ser 79); no afecta la cuota.")
    ws["K18"].font = Font(name=FONT, size=9, italic=True)
    R["d_int_row"] = 9
    R["d_amort_row"] = 10
    R["d_bal_row"] = 12
    R["d_ts_row"] = 13

    # ------------------------------------------------------------------ Proyeccion (bloque FCF) + estados proforma
    P = Hoja(wb, "Proyeccion", "Proyección 2008-2012 sin sinergias y estados financieros proforma (millones de US$)",
             [58, 4, 12, 12, 12, 12, 12, 12, 12, 50])

    def bloque_fcf(hoja, fila0, col_esc="C", titulo="Flujo de caja libre (escenario base)"):
        """Escribe el bloque de proyección operativa y valoración APV. Devuelve dict de filas."""
        ws_ = hoja.ws
        hoja.seccion(fila0, titulo)
        f = fila0 + 1
        hoja.texto(f"A{f}", "Año", negrita=True)
        hoja.c(f"C{f}", "2007 real", tipo="t", negrita=True)
        for col, a in zip(COLP, ANIOS_P):
            hoja.c(f"{col}{f}", str(a), tipo="t", negrita=True)
        hoja.c(f"I{f}", "2013 normal.", tipo="t", negrita=True)
        rows = {}
        nombres = ["S", "gS", "EQ", "REV", "SYS", "COGS", "SGA", "EBITDA", "mEBITDA", "DA", "EBIT", "TAXOP", "NOPAT", "CO",
                   "AR", "INV", "PREP", "AP", "DEF", "ACC", "NWC", "DNWC", "CAPEX", "FCF", "IC", "ROIC", "RR"]
        for i, n in enumerate(nombres):
            rows[n] = f + 1 + i
        r = rows
        esc = lambda k: f"Supuestos!${col_esc}${R[k]}"  # noqa: E731
        etiquetas = {
            "S": "Ingresos por servicios", "gS": "Crecimiento de ingresos por servicios", "EQ": "Ingresos por equipos",
            "REV": "Ingresos totales", "SYS": "Costos de operación del sistema", "COGS": "Costo de equipos vendidos",
            "SGA": "SG&A", "EBITDA": "EBITDA", "mEBITDA": "Margen EBITDA", "DA": "Depreciación y amortización",
            "EBIT": "EBIT", "TAXOP": "Impuestos operativos (EBIT × t)", "NOPAT": "Utilidad operativa después de impuestos (NOPAT)",
            "CO": "Costos operativos en efectivo", "AR": "Cuentas por cobrar", "INV": "Inventario", "PREP": "Gastos prepagados",
            "AP": "Cuentas por pagar", "DEF": "Ingresos diferidos", "ACC": "Pasivos acumulados",
            "NWC": "Capital de trabajo operativo neto", "DNWC": "Aumento del capital de trabajo operativo neto",
            "CAPEX": "Capex", "FCF": "Flujo de caja libre no apalancado", "IC": "Capital invertido contable (patrimonio + deuda)",
            "ROIC": "ROIC (NOPAT / capital invertido al inicio del año)",
            "RR": "Tasa de reinversión = (capex − D&A + Δ capital de trabajo) / NOPAT",
        }
        for n in nombres:
            hoja.texto(f"A{r[n]}", etiquetas[n], negrita=n in ("REV", "EBITDA", "NOPAT", "FCF"))
        # 2007 real
        h = lambda k: f"Historicos!$E${HR[k]}"  # noqa: E731
        for n, fm in [("S", h("S")), ("EQ", h("EQ")), ("REV", h("REV")), ("SYS", h("SYS")), ("COGS", h("COGS")),
                      ("SGA", h("SGA")), ("EBITDA", h("EBITDA")), ("DA", h("DA")), ("EBIT", h("EBIT")), ("AR", h("AR")),
                      ("INV", h("INV")), ("PREP", h("PREP")), ("AP", h("AP")), ("DEF", h("DEF")), ("ACC", h("ACC"))]:
            hoja.c(f"C{r[n]}", "=" + fm, F_MM, tipo="v")
        hoja.c(f"C{r['gS']}", f"=Historicos!$E${IND['g_s']}", F_PCT, tipo="v")
        hoja.c(f"C{r['mEBITDA']}", f"=C{r['EBITDA']}/C{r['REV']}", F_PCT)
        hoja.c(f"C{r['CO']}", f"=C{r['SYS']}+C{r['COGS']}+C{r['SGA']}", F_MM)
        hoja.c(f"C{r['NWC']}", f"=C{r['AR']}+C{r['INV']}+C{r['PREP']}-C{r['AP']}-C{r['DEF']}-C{r['ACC']}", F_MM)
        hoja.c(f"C{r['IC']}", f"=Historicos!$E${HR['PAT']}+Historicos!$E${HR['LTD']}+Historicos!$E${HR['NP']}", F_MM, tipo="v")
        for col, p in zip(COLP, ["C"] + COLP[:-1]):
            hoja.c(f"{col}{r['gS']}", f"=Supuestos!{col}${R['g_serv']}+{esc('esc_dg')}", F_PCT)
            hoja.c(f"{col}{r['S']}", f"={p}{r['S']}*(1+{col}{r['gS']})", F_MM)
            hoja.c(f"{col}{r['EQ']}", f"={col}{r['S']}*{R['eq_ratio']}", F_MM)
            hoja.c(f"{col}{r['REV']}", f"={col}{r['S']}+{col}{r['EQ']}", F_MM, negrita=True)
            hoja.c(f"{col}{r['SYS']}", f"={col}{r['S']}*Supuestos!{col}${R['sysop_pct']}", F_MM)
            hoja.c(f"{col}{r['COGS']}", f"={col}{r['EQ']}*Supuestos!{col}${R['cogs_pct']}", F_MM)
            hoja.c(f"{col}{r['SGA']}", f"={col}{r['REV']}*(Supuestos!{col}${R['sga_pct']}+{esc('esc_dsga')})", F_MM)
            hoja.c(f"{col}{r['EBITDA']}", f"={col}{r['REV']}-{col}{r['SYS']}-{col}{r['COGS']}-{col}{r['SGA']}", F_MM, negrita=True)
            hoja.c(f"{col}{r['mEBITDA']}", f"={col}{r['EBITDA']}/{col}{r['REV']}", F_PCT)
            hoja.c(f"{col}{r['DA']}", f"={col}{r['REV']}*Supuestos!{col}${R['dya_pct']}", F_MM)
            hoja.c(f"{col}{r['EBIT']}", f"={col}{r['EBITDA']}-{col}{r['DA']}", F_MM)
            hoja.c(f"{col}{r['TAXOP']}", f"={col}{r['EBIT']}*{R['t']}", F_MM)
            hoja.c(f"{col}{r['NOPAT']}", f"={col}{r['EBIT']}-{col}{r['TAXOP']}", F_MM, negrita=True)
            hoja.c(f"{col}{r['CO']}", f"={col}{r['SYS']}+{col}{r['COGS']}+{col}{r['SGA']}", F_MM)
            hoja.c(f"{col}{r['AR']}", f"={col}{r['REV']}*{R['dso']}/360", F_MM)
            hoja.c(f"{col}{r['INV']}", f"={col}{r['EQ']}*{R['dio']}/360", F_MM)
            hoja.c(f"{col}{r['PREP']}", f"={col}{r['CO']}*{R['prep']}", F_MM)
            hoja.c(f"{col}{r['AP']}", f"={col}{r['CO']}*{R['dpo']}/360", F_MM)
            hoja.c(f"{col}{r['DEF']}", f"={col}{r['S']}*{R['ddif']}/360", F_MM)
            hoja.c(f"{col}{r['ACC']}", f"={col}{r['CO']}*{R['dacum']}/360", F_MM)
            hoja.c(f"{col}{r['NWC']}", f"={col}{r['AR']}+{col}{r['INV']}+{col}{r['PREP']}-{col}{r['AP']}-{col}{r['DEF']}-{col}{r['ACC']}", F_MM)
            hoja.c(f"{col}{r['DNWC']}", f"={col}{r['NWC']}-{p}{r['NWC']}", F_MM)
            hoja.c(f"{col}{r['CAPEX']}", f"={col}{r['REV']}*(Supuestos!{col}${R['capex_pct']}+{esc('esc_dcapex')})", F_MM)
            hoja.c(f"{col}{r['FCF']}", f"={col}{r['NOPAT']}+{col}{r['DA']}-{col}{r['CAPEX']}-{col}{r['DNWC']}", F_MM, negrita=True)
            hoja.c(f"{col}{r['IC']}", f"={p}{r['IC']}+{col}{r['CAPEX']}-{col}{r['DA']}+{col}{r['DNWC']}", F_MM)
            hoja.c(f"{col}{r['ROIC']}", f"={col}{r['NOPAT']}/{p}{r['IC']}", F_PCT)
            hoja.c(f"{col}{r['RR']}", f"=({col}{r['CAPEX']}-{col}{r['DA']}+{col}{r['DNWC']})/{col}{r['NOPAT']}", F_PCT)
        # Valor terminal y APV
        f = r["RR"] + 2
        lin = {}

        def l(n, etq, fm, fmt=F_MM, neg=False, col="H"):
            nonlocal f
            hoja.texto(f"A{f}", etq, negrita=neg)
            hoja.c(f"{col}{f}", fm, fmt, negrita=neg)
            lin[n] = f"{col}{f}"
            f += 1

        l("g_fund", "Crecimiento fundamental = ROIC 2012 × tasa de reinversión 2012", f"=H{r['ROIC']}*H{r['RR']}", F_PCT2)
        l("g", "Crecimiento a perpetuidad usado (g)", f"=IF({R['sel_g']}=1,H{f-1},{R['sel_g_manual']})", F_PCT2, True)
        l("rr_t", "Tasa de reinversión de largo plazo = g / ROIC 2012", f"={lin['g']}/H{r['ROIC']}", F_PCT)
        hoja.c(f"I{r['NOPAT']}", f"=H{r['NOPAT']}*(1+{lin['g']})", F_MM)
        hoja.c(f"I{r['FCF']}", f"=I{r['NOPAT']}*(1-{lin['rr_t']})", F_MM, negrita=True)
        l("tv", "Valor terminal 2012 = FCF 2013 / (CPPC − g)", f"=I{r['FCF']}/({R['wacc']}-{lin['g']})", F_MM)
        l("tv_mult", "Valor terminal 2012 por múltiplo VE/EBITDA (contraste)",
          f"=H{r['EBITDA']}*IF({R['sel_mult']}=1,{R['evebitda_prom']},{R['evebitda_med']})", F_MM)
        l("pv_fcf", "VP de los FCF 2008-2012 a Ku", f"=NPV({R['ku']},D{r['FCF']}:H{r['FCF']})", F_MM)
        l("pv_tv", "VP del valor terminal a Ku", f"={lin['tv']}/(1+{R['ku']})^5", F_MM)
        l("pv_ts", "VP de los escudos fiscales 2008-2012 a Kd", f"=NPV({R['kd']},Deuda!$L$13:$P$13)", F_MM)
        l("v_ops", "Valor de las operaciones (APV)", f"={lin['pv_fcf']}+{lin['pv_tv']}+{lin['pv_ts']}", F_MM, True)
        l("nooper", "+ Activos no operativos", "=Valoracion!$C$20", F_MM)
        l("claims", "− Deuda existente e interés minoritario", "=Valoracion!$C$26", F_MM)
        l("equity", "Valor del patrimonio (100%, antes de descuento por iliquidez)",
          f"={lin['v_ops']}+{lin['nooper']}-{lin['claims']}", F_MM, True)
        return r, lin

    BR, BL = bloque_fcf(P, 3)
    for col in COLP:
        S.c(f"{col}{R['dya_pct']}", f"={col}{R['dya']}/Proyeccion!{col}{BR['REV']}", F_PCT2)
    R["P_rows"] = BR
    for k, v in BL.items():
        R["P_" + k] = f"Proyeccion!${v[0]}${v[1:]}"

    # Estados proforma con estructura de adquisición
    ws = P.ws
    f = int(BL["equity"][1:]) + 3
    P.seccion(f, "Estados financieros proforma con la deuda de adquisición (variable de cierre: caja / línea de crédito)")
    f += 1
    P.c(f"C{f}", "2007 proforma", tipo="t", negrita=True)
    for col, a in zip(COLP, ANIOS_P):
        P.c(f"{col}{f}", str(a), tipo="t", negrita=True)
    P.texto(f"J{f}", "Supuestos del proforma", negrita=True)
    f += 1
    PF = {}
    orden = ["EBIT", "INTD", "INTR", "EBT", "TAX", "NI", "_", "CASH", "AR", "INV", "PREP", "OCA", "FA", "GW", "ONOA", "TA", "_",
             "AP", "DEF", "ACC", "OCL", "REVOLV", "DEBT", "OLTL", "MINT", "EQ", "TLE", "CHK", "_",
             "CFO", "CFI", "PRINC", "DREV", "DCASH", "CHK2", "DISP", "_",
             "i_mebit", "i_mneto", "i_roe", "i_debitda", "i_cov", "i_dv"]
    for n in orden:
        if n != "_":
            PF[n] = f
        f += 1
    et = {
        "EBIT": "EBIT", "INTD": "Intereses de la deuda de adquisición", "INTR": "Intereses de la línea de crédito (Kd sobre saldo inicial)",
        "EBT": "Utilidad antes de impuestos", "TAX": "Impuestos", "NI": "Utilidad neta (solo operaciones)",
        "CASH": "Caja", "AR": "Cuentas por cobrar", "INV": "Inventario", "PREP": "Gastos prepagados",
        "OCA": "Otros activos corrientes e impuestos diferidos (constantes)", "FA": "Activo fijo neto, licencias y listas de clientes",
        "GW": "Plusvalía (constante)", "ONOA": "Afiliadas, nota por cobrar y otros activos LP (constantes)", "TA": "Total activos",
        "AP": "Cuentas por pagar", "DEF": "Ingresos diferidos", "ACC": "Pasivos acumulados",
        "OCL": "Impuestos por pagar y otros pasivos corrientes (constantes)", "REVOLV": "Línea de crédito (variable de cierre)",
        "DEBT": "Deuda de adquisición (2012: saldo a refinanciar)", "OLTL": "Impuestos diferidos, retiro de activos y otros pasivos LP (constantes)",
        "MINT": "Interés minoritario (constante)", "EQ": "Patrimonio", "TLE": "Total pasivos y patrimonio", "CHK": "Control: activos − pasivos y patrimonio",
        "CFO": "Flujo operativo = utilidad neta + D&A − Δ capital de trabajo", "CFI": "Flujo de inversión = − capex",
        "PRINC": "Amortización ordinaria de la deuda", "DREV": "Variación de la línea de crédito", "DCASH": "Variación de caja",
        "CHK2": "Control: variación de caja del flujo − del balance",
        "DISP": "Caja disponible antes de línea de crédito (caja inicial − línea inicial + flujos − amortización)",
        "i_mebit": "Indicador: margen EBIT", "i_mneto": "Indicador: margen neto", "i_roe": "Indicador: ROE (utilidad neta / patrimonio inicial)",
        "i_debitda": "Indicador: deuda total / EBITDA", "i_cov": "Indicador: EBITDA / intereses totales",
        "i_dv": "Indicador: deuda total / valor de las operaciones (solo 2007)",
    }
    for n, row in PF.items():
        P.texto(f"A{row}", et[n], negrita=n in ("NI", "TA", "TLE", "EQ", "DCASH"))
    h = lambda k: f"Historicos!$E${HR[k]}"  # noqa: E731
    # 2007 proforma
    P.c(f"C{PF['EBIT']}", f"=C{BR['EBIT']}", F_MM)
    P.c(f"C{PF['CASH']}", 0, F_MM, tipo="i", nota="Supuesto: la caja y los valores negociables de 2007 se usan en la transacción (se valoran como activo no operativo).")
    for n, k in [("AR", "AR"), ("INV", "INV"), ("PREP", "PREP"), ("AP", "AP"), ("DEF", "DEF"), ("ACC", "ACC")]:
        P.c(f"C{PF[n]}", f"=C{BR[k]}", F_MM)
    P.c(f"C{PF['OCA']}", f"={h('OCA')}+{h('DTA')}", F_MM, tipo="v")
    P.c(f"C{PF['FA']}", f"=Historicos!$E${HR['FA']}", F_MM, tipo="v")
    P.c(f"C{PF['GW']}", f"={h('GW')}", F_MM, tipo="v")
    P.c(f"C{PF['ONOA']}", f"={h('INVAF')}+{h('NOTE')}+{h('OLTA')}+{h('MES')}", F_MM, tipo="v")
    P.c(f"C{PF['OCL']}", f"={h('TXP')}+{h('OCL')}+{h('DTLC')}", F_MM, tipo="v")
    P.c(f"C{PF['REVOLV']}", 0, F_MM, tipo="i")
    P.c(f"C{PF['DEBT']}", f"={R['d0']}", F_MM, tipo="v")
    P.c(f"C{PF['OLTL']}", f"={h('DTL')}+{h('ARO')}+{h('ODL')}", F_MM, tipo="v")
    P.c(f"C{PF['MINT']}", f"={h('MINT')}", F_MM, tipo="v")
    P.c(f"C{PF['EQ']}", f"=Historicos!$E${HR['PAT']}-({R['d0']}-{h('LTD')}-{h('NP')})-({h('CASH')}+{h('MKS')})", F_MM,
        nota="Patrimonio 2007 menos la deuda nueva neta de la refinanciada, menos la caja usada en la compra. Sin contabilidad de compra (no se registra la nueva plusvalía).")
    for col, p, a in zip(COLP, ["C"] + COLP[:-1], ANIOS_P):
        dcol = "LMNOP"[a - 2008]
        P.c(f"{col}{PF['EBIT']}", f"={col}{BR['EBIT']}", F_MM)
        P.c(f"{col}{PF['INTD']}", f"=Deuda!{dcol}$9", F_MM, tipo="v")
        P.c(f"{col}{PF['INTR']}", f"={p}{PF['REVOLV']}*{R['kd']}", F_MM)
        P.c(f"{col}{PF['EBT']}", f"={col}{PF['EBIT']}-{col}{PF['INTD']}-{col}{PF['INTR']}", F_MM)
        P.c(f"{col}{PF['TAX']}", f"={col}{PF['EBT']}*{R['t']}", F_MM)
        P.c(f"{col}{PF['NI']}", f"={col}{PF['EBT']}-{col}{PF['TAX']}", F_MM, negrita=True)
        P.c(f"{col}{PF['DISP']}", f"={p}{PF['CASH']}-{p}{PF['REVOLV']}+{col}{PF['CFO']}+{col}{PF['CFI']}-{col}{PF['PRINC']}", F_MM)
        P.c(f"{col}{PF['CASH']}", f"=MAX(0,{col}{PF['DISP']})", F_MM)
        P.c(f"{col}{PF['REVOLV']}", f"=MAX(0,-{col}{PF['DISP']})", F_MM)
        for n, k in [("AR", "AR"), ("INV", "INV"), ("PREP", "PREP"), ("AP", "AP"), ("DEF", "DEF"), ("ACC", "ACC")]:
            P.c(f"{col}{PF[n]}", f"={col}{BR[k]}", F_MM)
        for n in ["OCA", "GW", "ONOA", "OCL", "OLTL", "MINT"]:
            P.c(f"{col}{PF[n]}", f"={p}{PF[n]}", F_MM)
        P.c(f"{col}{PF['FA']}", f"={p}{PF['FA']}+{col}{BR['CAPEX']}-{col}{BR['DA']}", F_MM)
        P.c(f"{col}{PF['DEBT']}", f"=Deuda!{dcol}$12", F_MM, tipo="v")
        P.c(f"{col}{PF['EQ']}", f"={p}{PF['EQ']}+{col}{PF['NI']}", F_MM)
        P.c(f"{col}{PF['CFO']}", f"={col}{PF['NI']}+{col}{BR['DA']}-{col}{BR['DNWC']}", F_MM)
        P.c(f"{col}{PF['CFI']}", f"=-{col}{BR['CAPEX']}", F_MM)
        P.c(f"{col}{PF['PRINC']}", f"=Deuda!{dcol}$10", F_MM, tipo="v")
        P.c(f"{col}{PF['DREV']}", f"={col}{PF['REVOLV']}-{p}{PF['REVOLV']}", F_MM)
        P.c(f"{col}{PF['DCASH']}", f"={col}{PF['CFO']}+{col}{PF['CFI']}-{col}{PF['PRINC']}+{col}{PF['DREV']}", F_MM, negrita=True)
        P.c(f"{col}{PF['CHK2']}", f"={col}{PF['DCASH']}-({col}{PF['CASH']}-{p}{PF['CASH']})", F_MM)
        P.c(f"{col}{PF['i_roe']}", f"={col}{PF['NI']}/{p}{PF['EQ']}", F_PCT)
        P.c(f"{col}{PF['i_cov']}", f"={col}{BR['EBITDA']}/({col}{PF['INTD']}+{col}{PF['INTR']})", F_X)
        P.c(f"{col}{PF['i_mneto']}", f"={col}{PF['NI']}/{col}{BR['REV']}", F_PCT)
    for col in ["C"] + COLP:
        P.c(f"{col}{PF['TA']}", f"={col}{PF['CASH']}+{col}{PF['AR']}+{col}{PF['INV']}+{col}{PF['PREP']}+{col}{PF['OCA']}+{col}{PF['FA']}+{col}{PF['GW']}+{col}{PF['ONOA']}", F_MM, negrita=True)
        P.c(f"{col}{PF['TLE']}", f"={col}{PF['AP']}+{col}{PF['DEF']}+{col}{PF['ACC']}+{col}{PF['OCL']}+{col}{PF['REVOLV']}+{col}{PF['DEBT']}+{col}{PF['OLTL']}+{col}{PF['MINT']}+{col}{PF['EQ']}", F_MM, negrita=True)
        P.c(f"{col}{PF['CHK']}", f"={col}{PF['TA']}-{col}{PF['TLE']}", F_MM)
        P.c(f"{col}{PF['i_mebit']}", f"={col}{BR['EBIT']}/{col}{BR['REV']}", F_PCT)
        P.c(f"{col}{PF['i_debitda']}", f"=({col}{PF['DEBT']}+{col}{PF['REVOLV']})/{col}{BR['EBITDA']}", F_X)
    P.c(f"C{PF['i_dv']}", f"=C{PF['DEBT']}/{R['P_v_ops']}", F_PCT)
    P.c(f"I{PF['CHK']}", f"=MAX(ABS(C{PF['CHK']}),ABS(D{PF['CHK']}),ABS(E{PF['CHK']}),ABS(F{PF['CHK']}),ABS(G{PF['CHK']}),ABS(H{PF['CHK']}))", F_MM, clave="chk_pf_bs")
    P.c(f"I{PF['CHK2']}", f"=MAX(ABS(D{PF['CHK2']}),ABS(E{PF['CHK2']}),ABS(F{PF['CHK2']}),ABS(G{PF['CHK2']}),ABS(H{PF['CHK2']}))", F_MM, clave="chk_pf_cf")
    notas_pf = [
        "Solo operaciones: no se proyectan afiliadas, ganancias en inversiones, otros ingresos ni interés minoritario (se valoran aparte).",
        "La deuda existente (1.002) se refinancia con la deuda de adquisición (3.758).",
        "Caja inicial 0: la caja de 2007 se usa en la compra. La caja no genera intereses (conservador).",
        "Sin dividendos: todo el excedente queda como caja; los faltantes los cubre la línea de crédito a Kd.",
        "En 2012 el saldo bullet se refinancia (no sale caja); a partir de ahí la deuda sigue la estructura objetivo.",
    ]
    for i, t in enumerate(notas_pf):
        P.texto(f"J{PF['EBIT'] + i}", t)
    R["PF"] = PF

    # ------------------------------------------------------------------ Sinergias
    SY = Hoja(wb, "Sinergias", "Sinergias con American Cable (Tablas 4 y 5), millones de US$", [58, 4, 12, 12, 12, 12, 12, 12, 12, 50])
    ws = SY.ws
    SY.c("C3", "2007", tipo="t", negrita=True)
    for col, a in zip(COLP, ANIOS_P):
        SY.c(f"{col}3", str(a), tipo="t", negrita=True)
    SY.c("I3", "2013", tipo="t", negrita=True)
    sy = ["SYS", "BHP", "BHC", "RED", "SAV", "_", "SUB", "MIN", "TMIN", "RPM", "BREV", "MRG", "BEBITDA", "_", "TOT", "TOTAT", "TOTR"]
    SR = {}
    rr = 4
    for n in sy:
        if n != "_":
            SR[n] = rr
        rr += 1
    ets = {"SYS": "Costos de operación del sistema (sin sinergias)", "BHP": "Backhaul / costos del sistema", "BHC": "Costo estimado de backhaul",
           "RED": "Reducción del costo de backhaul", "SAV": "Ahorro de backhaul (antes de impuestos)",
           "SUB": "Suscriptores empresariales promedio (millones)", "MIN": "Minutos mensuales por suscriptor",
           "TMIN": "Minutos mensuales totales (millones)", "RPM": "Ingreso por minuto (US$)", "BREV": "Ingreso anual adicional por clientes empresariales",
           "MRG": "Margen EBITDA aplicado a los ingresos adicionales (selector)", "BEBITDA": "EBITDA adicional por clientes empresariales",
           "TOT": "Sinergias totales antes de impuestos", "TOTAT": "Sinergias después de impuestos", "TOTR": "Sinergias después de impuestos × % realizado"}
    for n, row in SR.items():
        SY.texto(f"A{row}", ets[n], negrita=n in ("SAV", "BREV", "TOTR"))
    for col in COLP:
        SY.c(f"{col}{SR['SYS']}", f"=Proyeccion!{col}{BR['SYS']}", F_MM, tipo="v")
        SY.c(f"{col}{SR['BHP']}", f"=Supuestos!{col}${R['bh_pct']}", F_PCT, tipo="v")
        SY.c(f"{col}{SR['BHC']}", f"={col}{SR['SYS']}*{col}{SR['BHP']}", F_MM)
        SY.c(f"{col}{SR['RED']}", f"=Supuestos!{col}${R['bh_red']}", F_PCT, tipo="v")
        SY.c(f"{col}{SR['SAV']}", f"={col}{SR['BHC']}*{col}{SR['RED']}", F_MM, negrita=True)
        SY.c(f"{col}{SR['SUB']}", f"=Supuestos!{col}${R['sus_emp']}", "0.00", tipo="v")
        SY.c(f"{col}{SR['MIN']}", f"=Supuestos!{col}${R['min_emp']}", "#,##0", tipo="v")
        SY.c(f"{col}{SR['TMIN']}", f"={col}{SR['SUB']}*{col}{SR['MIN']}", "#,##0.0")
        SY.c(f"{col}{SR['RPM']}", f"=Supuestos!{col}${R['ipm_emp']}", "0.0000", tipo="v")
        SY.c(f"{col}{SR['BREV']}", f"={col}{SR['TMIN']}*{col}{SR['RPM']}*12", F_MM, negrita=True)
        SY.c(f"{col}{SR['MRG']}", f"=IF({R['sel_margen_sin']}=1,Proyeccion!{col}{BR['mEBITDA']},1-Supuestos!{col}${R['sysop_pct']}-Supuestos!{col}${R['sga_pct']})", F_PCT)
        SY.c(f"{col}{SR['BEBITDA']}", f"={col}{SR['BREV']}*{col}{SR['MRG']}", F_MM)
        SY.c(f"{col}{SR['TOT']}", f"={col}{SR['SAV']}+{col}{SR['BEBITDA']}", F_MM)
        SY.c(f"{col}{SR['TOTAT']}", f"={col}{SR['TOT']}*(1-{R['t']})", F_MM)
        SY.c(f"{col}{SR['TOTR']}", f"={col}{SR['TOTAT']}*{R['realiz_sin']}", F_MM, negrita=True)
    # Separación en dos componentes
    rr = SR["TOTR"] + 2
    SY.texto(f"A{rr}", "Valoración (descontadas a Ku; perpetuidad desde 2013 creciendo a g, sin inversión adicional)", negrita=True)
    rr += 1
    for col, t in zip("CDE", ["Backhaul", "Ingresos", "Total"]):
        SY.c(f"{col}{rr}", t, tipo="t", negrita=True)
    rr += 1
    SY.texto(f"A{rr}", "VP de los flujos 2008-2012")
    SY.c(f"C{rr}", f"=NPV({R['ku']},D{SR['SAV']}:H{SR['SAV']})*(1-{R['t']})*{R['realiz_sin']}", F_MM)
    SY.c(f"D{rr}", f"=NPV({R['ku']},D{SR['BEBITDA']}:H{SR['BEBITDA']})*(1-{R['t']})*{R['realiz_sin']}", F_MM)
    SY.c(f"E{rr}", f"=C{rr}+D{rr}", F_MM)
    rr += 1
    SY.texto(f"A{rr}", "VP del valor terminal")
    SY.c(f"C{rr}", f"=H{SR['SAV']}*(1-{R['t']})*{R['realiz_sin']}*(1+{R['P_g']})/({R['ku']}-{R['P_g']})/(1+{R['ku']})^5", F_MM)
    SY.c(f"D{rr}", f"=H{SR['BEBITDA']}*(1-{R['t']})*{R['realiz_sin']}*(1+{R['P_g']})/({R['ku']}-{R['P_g']})/(1+{R['ku']})^5", F_MM)
    SY.c(f"E{rr}", f"=C{rr}+D{rr}", F_MM)
    rr += 1
    SY.texto(f"A{rr}", "Valor presente de las sinergias", negrita=True)
    for col, k in zip("CDE", ["sin_bh", "sin_rev", "sin_tot"]):
        SY.c(f"{col}{rr}", f"={col}{rr-2}+{col}{rr-1}", F_MM, negrita=True, clave=k)
    rr += 1
    SY.texto(f"A{rr}", "Control: total = VP de la fila 'sinergias × % realizado' + valor terminal")
    SY.c(f"E{rr}", f"=E{rr-1}-(NPV({R['ku']},D{SR['TOTR']}:H{SR['TOTR']})+H{SR['TOTR']}*(1+{R['P_g']})/({R['ku']}-{R['P_g']})/(1+{R['ku']})^5)", F_MM, clave="chk_sin")
    rr += 1
    SY.texto(f"A{rr}", "VP de las sinergias con 100% de realización (para sensibilidad)")
    SY.c(f"E{rr}", f"=NPV({R['ku']},D{SR['TOTAT']}:H{SR['TOTAT']})+H{SR['TOTAT']}*(1+{R['P_g']})/({R['ku']}-{R['P_g']})/(1+{R['ku']})^5", F_MM, clave="sin_100")
    rr += 2
    for t in [
        "Supuestos propios: (1) el caso solo da ingresos empresariales, no su margen; base = margen EBITDA proyectado de ATC;",
        "(2) no se asume capex ni capital de trabajo adicional (el caso dice que el negocio empresarial usa capacidad ociosa de la red);",
        "(3) ambos tipos de sinergia se descuentan a Ku; las de ingresos son más inciertas, por eso se muestran por separado;",
        "(4) los costos de integración no aparecen en el caso y no se incluyen (sesgo al alza del valor de las sinergias).",
    ]:
        SY.texto(f"A{rr}", t, italica=True)
        rr += 1
    R["SR"] = SR

    # ------------------------------------------------------------------ Valoracion
    V = Hoja(wb, "Valoracion", "Valoración de AirThread: APV (método de American Cable) y contrastes (millones de US$)",
             [62, 4, 13, 13, 13, 13, 4, 70])
    ws = V.ws
    V.seccion(3, "A. Valor de las operaciones (APV, sin sinergias)")
    vrows = [
        (5, "Costo de capital no apalancado (Ku)", f"={R['ku']}", F_PCT2, "v", "v_ku"),
        (6, "CPPC objetivo para el valor terminal", f"={R['wacc']}", F_PCT2, "v", "v_wacc"),
        (7, "Crecimiento a perpetuidad (g)", f"={R['P_g']}", F_PCT2, "v", "v_g"),
        (8, "VP de los flujos libres 2008-2012", f"={R['P_pv_fcf']}", F_MM, "v", "v_pvfcf"),
        (9, "VP del valor terminal (perpetuidad)", f"={R['P_pv_tv']}", F_MM, "v", "v_pvtv"),
        (10, "VP de los escudos fiscales de la deuda de adquisición 2008-2012", f"={R['P_pv_ts']}", F_MM, "v", "v_pvts"),
        (11, "Valor de las operaciones (APV)", f"={R['P_v_ops']}", F_MM, "v", "v_ops"),
        (12, "Valor terminal / valor de las operaciones", "=C9/C11", F_PCT, "f", "v_tvshare"),
    ]
    for r_, etq, fm, fmt, tipo, clave in vrows:
        V.texto(f"A{r_}", etq, negrita=r_ == 11)
        V.c(f"C{r_}", fm, fmt, tipo=tipo, clave=clave, negrita=r_ == 11)
    V.seccion(14, "B. Activos no operativos y otros derechos")
    nooper = [
        (15, "Caja y valores negociables 2007", f"=Historicos!$E${HR['CASH']}+Historicos!$E${HR['MKS']}", "no_caja",
         "Caso, p. 6: probablemente se usarán para financiar la compra. Se suma completa."),
        (16, "Participación en utilidades de afiliadas 2007", f"=Historicos!$E${HR['AFIL']}", "no_afil_ni", "Anexo 4."),
        (17, "Valor de las afiliadas", f"=IF({R['sel_nooper']}=1,C16*{R['pe']},Historicos!$E${HR['INVAF']})", "no_afil",
         "Base: P/U de la industria × utilidad (caso, nota 3). Alternativa: valor en libros (157,7)."),
        (18, "Nota por cobrar de largo plazo", f"=Historicos!$E${HR['NOTE']}", "no_nota", "Anexo 5, a valor en libros."),
        (19, "Títulos de renta variable negociables", f"=Historicos!$E${HR['MES']}", "no_mes", "Anexo 5 (cero en 2007)."),
        (20, "Total activos no operativos", "=C15+C17+C18+C19", "no_total", ""),
        (22, "Deuda existente (largo plazo + pagaré), se refinancia", f"=Historicos!$E${HR['LTD']}+Historicos!$E${HR['NP']}", "cl_deuda", "Anexo 5."),
        (23, "Interés minoritario en resultados 2007", f"=Historicos!$E${HR['MI']}", "cl_mi_ni", "Anexo 4."),
        (24, "Valor del interés minoritario", f"=IF({R['sel_nooper']}=1,C23*{R['pe']},Historicos!$E${HR['MINT']})", "cl_mi",
         "Coherente con afiliadas: P/U de la industria. Alternativa: libros (43,4)."),
        (26, "Total deuda e interés minoritario", "=C22+C24", "cl_total", ""),
    ]
    for r_, etq, fm, clave, nota in nooper:
        V.texto(f"A{r_}", etq, negrita=r_ in (20, 26))
        V.c(f"C{r_}", fm, F_MM, clave=clave, negrita=r_ in (20, 26))
        V.texto(f"H{r_}", nota)
    V.seccion(28, "C. Valor del patrimonio")
    vals = [
        (29, "Valor de las operaciones (APV)", "=C11", "e_ops"),
        (30, "+ Activos no operativos", "=C20", "e_noop"),
        (31, "Valor total de la empresa", "=C29+C30", "e_ve"),
        (32, "− Deuda existente e interés minoritario", "=C26", "e_claims"),
        (33, "Valor del patrimonio, 100%, sin sinergias (valor intrínseco independiente)", "=C31-C32", "e_eq"),
        (34, "Descuento por iliquidez aplicado", f"={R['iliq']}", "e_iliq"),
        (35, "Valor del patrimonio después del descuento", "=C33*(1-C34)", "e_eq_iliq"),
        (36, "+ VP sinergias de backhaul", f"={R['sin_bh']}", "e_sin_bh"),
        (37, "+ VP sinergias de ingresos empresariales", f"={R['sin_rev']}", "e_sin_rev"),
        (38, "Valor del patrimonio con todas las sinergias", "=C35+C36+C37", "e_eq_sin"),
        (39, "Control: bloque de la hoja Proyeccion − fila 33", f"={R['P_equity']}-C33", "chk_val"),
    ]
    for r_, etq, fm, clave in vals:
        V.texto(f"A{r_}", etq, negrita=r_ in (33, 38))
        V.c(f"C{r_}", fm, F_PCT if r_ == 34 else F_MM, clave=clave, negrita=r_ in (33, 38))
    V.seccion(41, "D. Contrastes")
    ctr = [
        (42, "Valor terminal por múltiplo VE/EBITDA de comparables (2012)", f"={R['P_tv_mult']}", F_MM, "x_tvm"),
        (43, "Múltiplo VE/EBITDA usado", f"=IF({R['sel_mult']}=1,{R['evebitda_prom']},{R['evebitda_med']})", F_X1, "x_mult"),
        (44, "Múltiplo VE/EBITDA 2012 implícito en el valor terminal por perpetuidad", f"={R['P_tv']}/Proyeccion!H{BR['EBITDA']}", F_X1, "x_mult_impl"),
        (45, "Valor de las operaciones con valor terminal por múltiplo", f"=C8+C10+C42/(1+C5)^5", F_MM, "x_vops_mult"),
        (46, "Patrimonio con valor terminal por múltiplo (sin sinergias)", "=C45+C30-C32", F_MM, "x_eq_mult"),
        (47, "Crecimiento implícito en el múltiplo: g = (VT×CPPC − FCF₂₀₁₃ con reinversión) …", None, F_PCT2, "x_g_impl"),
        (48, "CPPC-DCF a estructura objetivo constante: valor de las operaciones",
         f"=NPV(C6,Proyeccion!D{BR['FCF']}:H{BR['FCF']})+Proyeccion!{R['P_tv'].split('!')[1].replace('$','')}/(1+C6)^5", F_MM, "x_vops_wacc"),
        (49, "Patrimonio por CPPC-DCF (sin sinergias)", "=C48+C30-C32", F_MM, "x_eq_wacc"),
        (50, "Diferencia APV − CPPC-DCF (valor del endeudamiento inicial por encima del objetivo, neto)", "=C33-C49", F_MM, "x_dif_apv"),
        (51, "VE/EBITDA 2007 implícito en el valor de las operaciones", f"=C11/Historicos!$E${HR['EBITDA']}", F_X1, "x_ev_ebitda07"),
        (52, "VE/EBITDA 2008 implícito en el valor de las operaciones", f"=C11/Proyeccion!D{BR['EBITDA']}", F_X1, "x_ev_ebitda08"),
        (53, "Valor del patrimonio por múltiplo VE/EBITDA 2007 actual (mediana)", f"=Historicos!$E${HR['EBITDA']}*{R['evebitda_med']}+C30-C32", F_MM, "x_eq_mult07"),
        (54, "Valor del patrimonio por P/U de la industria × utilidad neta 2007", f"=Historicos!$E${HR['NI']}*{R['pe']}", F_MM, "x_eq_pe"),
        (55, "Deuda de adquisición / valor de las operaciones (límite 50%)", f"={R['d0']}/C11", F_PCT, "x_dv0"),
        (56, "Deuda de adquisición / valor total de la empresa con activos no operativos", f"={R['d0']}/C31", F_PCT, "x_dv0_tot"),
        (63, "Rendimiento anual exigido sobre el valor de las operaciones = Ku × valor (APV)", "=C5*C11", F_MM, "x_rend"),
        (64, "Afiliadas / valor del patrimonio sin sinergias", "=C17/C33", F_PCT, "x_afil_sh"),
        (65, "Sinergias de ingresos / total de sinergias", "=C37/(C36+C37)", F_PCT, "x_sinrev_sh"),
        (66, "Prima del precio máximo sobre el valor independiente", "=C60/C33-1", F_PCT, "x_prima_max"),
        (67, "VE/EBITDA 2007 implícito en el precio máximo (con deuda existente y minoritario)", f"=(C60+C32-C30)/Historicos!$E${HR['EBITDA']}", F_X1, "x_mult_max"),
    ]
    for r_, etq, fm, fmt, clave in ctr:
        V.texto(f"A{r_}", etq)
        if fm is not None:
            V.c(f"C{r_}", fm, fmt, clave=clave)
    # g implícito en el múltiplo: VT = NOPAT12(1+g)(1-g/ROIC)/(CPPC-g) -> resolver cuadrática
    # VT(CPPC-g) = N(1+g)(1-g/Ro) -> N/Ro g^2 + g(VT + N/Ro - N) ... se resuelve abajo
    V.ws["A47"] = "Crecimiento implícito en el múltiplo (resuelve VT = NOPAT₂₀₁₂(1+g)(1−g/ROIC)/(CPPC−g))"
    nn = f"Proyeccion!$H${BR['NOPAT']}"
    ro = f"Proyeccion!$H${BR['ROIC']}"
    # a g^2 + b g + c = 0 con a = N/Ro, b = -(VT + N - N/Ro), c = VT*CPPC - N
    V.c("E47", f"={nn}/{ro}", F_NUM)
    V.c("F47", f"=-(C42+{nn}-{nn}/{ro})", F_NUM)
    V.c("D47", f"=C42*C6-{nn}", F_NUM)
    V.c("C47", "=(-F47-SQRT(F47^2-4*E47*D47))/(2*E47)", F_PCT2, clave="x_g_impl")
    V.texto("H47", "Celdas D-F: coeficientes c, a, b de la cuadrática.")

    V.seccion(57, "E. Rango de precio para la negociación (patrimonio 100%, millones de US$)")
    neg = [
        (58, "Apertura: valor independiente con el descuento por iliquidez de una salida a bolsa (10%, nota 6)",
         f"=C33*(1-{R['ipo']})", "n_apertura"),
        (59, "Objetivo: valor independiente + 50% de las sinergias de backhaul", f"=C33+{R['cesion_bh']}*C36", "n_objetivo"),
        (60, "Precio máximo (retiro): valor independiente + 100% backhaul (sin pagar sinergias de ingresos)", "=C33+C36", "n_maximo"),
        (61, "Referencia: valor con todas las sinergias (no pagar)", "=C33+C36+C37", "n_todas"),
    ]
    for r_, etq, fm, clave in neg:
        V.texto(f"A{r_}", etq, negrita=r_ == 60)
        V.c(f"C{r_}", fm, F_MM, clave=clave, negrita=r_ == 60)
    V.texto("H58", "Costo de salir a bolsa: 10% (nota 6 del caso).")
    V.texto("H59", "Comparte con el vendedor la mitad de la sinergia más cierta.")
    V.texto("H60", "Por encima, ACC transfiere al vendedor sinergias que solo ACC crea y que son inciertas.")

    # ------------------------------------------------------------------ Costo_Capital B: opciones (necesita filas de Proyeccion)
    ws = K.ws
    fcf_rng = f"Proyeccion!$D${BR['FCF']}:$H${BR['FCF']}"
    n12 = f"Proyeccion!$H${BR['NOPAT']}"
    g = R["P_g"]
    pvts = R["P_pv_ts"]
    noop = "Valoracion!$C$20"
    cl = "Valoracion!$C$26"
    fcf13 = f"Proyeccion!$I${BR['FCF']}"

    def eq_formula(ku, wacc, gg=g, f13=fcf13):
        return f"=NPV({ku},{fcf_rng})+{pvts}+{f13}/(({wacc})-{gg})/(1+{ku})^5+{noop}-{cl}"

    opciones = []
    # (etiqueta, valor: número = insumo o texto = vínculo, Ku, CPPC, fuente); {v} = celda de la columna C de la fila
    bu1, bu2, bu3 = "Comparables!$C$28", "Comparables!$D$28", "Comparables!$E$28"
    rf, erp, kd, t, L, de = "$C$4", "$C$5", "$C$6", "$C$8", "$C$11", "$C$12"
    opciones.append(("Beta: 1 Hamada, βd = 0", bu1, "={rf}+{v}*{erp}",
                     "=(1-{L})*({rf}+{v}*(1+(1-{t})*{de})*{erp})+{L}*{kd}*(1-{t})", "Supone deuda fija y sin riesgo; incoherente con Kd > rf."))
    opciones.append(("Beta: 2 Harris-Pringle, βd = 0", bu2, "={rf}+{v}*{erp}", "={rf}+{v}*{erp}-{t}*{kd}*{L}",
                     "Deuda reequilibrada; ignora el riesgo de la deuda."))
    opciones.append(("Beta: 3 Harris-Pringle, βd = 0,25 (elegida)", bu3, "={rf}+{v}*{erp}", "={rf}+{v}*{erp}-{t}*{kd}*{L}",
                     "Coherente con Kd = 5,5% y con un valor terminal a CPPC con deuda reequilibrada."))
    prm_ku = "={rf}+{v}*Comparables!$D$28+({kd}-{rf})*(1-Comparables!$B$28)"
    opciones.append(("PRM 4% (supuesto propio, sensibilidad)", 0.04, prm_ku, "=D{r}-{t}*{kd}*{L}",
                     "El caso solo da 5%. Ku = rf + PRM × promedio(E/V × βe) + (Kd − rf) × D/V promedio (método 3)."))
    opciones.append(("PRM 6% (supuesto propio, sensibilidad)", 0.06, prm_ku, "=D{r}-{t}*{kd}*{L}", "Ídem."))
    opciones.append(("D/V objetivo: promedio comparables (elegida)", R["dv_prom"], "=$C$10", "=$C$10-{t}*{kd}*{v}",
                     "Anexo 7; el caso pide converger a la norma de la industria."))
    opciones.append(("D/V objetivo: mediana comparables", R["dv_med"], "=$C$10", "=$C$10-{t}*{kd}*{v}", "Anexo 7."))
    opciones.append(("D/V objetivo: 50% (máximo de R&R) a perpetuidad", R["dv_max"], "=$C$10", "=$C$10-{t}*{kd}*{v}",
                     "Agresivo: R&R lo plantea para la compra, no como estructura permanente."))
    opciones.append(("Primas por tamaño, regulación o país", 0.0, "=$C$10+{v}", "=$C$16+{v}",
                     "No se suman: los comparables son operadores de EE. UU. con la misma regulación (FCC); el caso no da prima por tamaño."))
    r_ = 24
    for etq, val, ku_f, wacc_f, fuente in opciones:
        K.texto(f"A{r_}", etq)
        fmt = F_NUM if etq.startswith("Beta") else F_PCT2
        if isinstance(val, float):
            K.c(f"C{r_}", val, fmt, tipo="i")
        else:
            K.c(f"C{r_}", "=" + val, fmt, tipo="v")
        sub = dict(rf=rf, erp=erp, kd=kd, t=t, L=L, de=de, v=f"C{r_}", r=r_)
        K.c(f"D{r_}", ku_f.format(**sub), F_PCT2)
        K.c(f"E{r_}", wacc_f.format(**sub), F_PCT2)
        K.c(f"F{r_}", eq_formula(f"D{r_}", f"E{r_}"), F_MM)
        K.c(f"G{r_}", f"=F{r_}-Valoracion!$C$33", F_MM)
        K.texto(f"H{r_}", fuente)
        r_ += 1
    R["k_opt_end"] = r_ - 1
    r_ += 1
    K.texto(f"A{r_}", "Variable decisiva: crecimiento a perpetuidad y método del valor terminal (ver hojas Sensibilidad y Valoracion).", italica=True)
    r_ += 1
    K.texto(f"A{r_}", "Moneda e inflación: flujos en US$ nominales; rf nominal en US$; el caso no da inflación. g nominal.", italica=True)
    r_ += 1
    K.texto(f"A{r_}", "Plazo de rf: 10 años, el único que da el caso; los flujos son perpetuos (un bono a 20-30 años sería más coherente).", italica=True)

    # ------------------------------------------------------------------ Escenarios
    E = Hoja(wb, "Escenarios", "Escenarios de riesgo: el mismo bloque de cálculo con ajustes (millones de US$)",
             [58, 4, 12, 12, 12, 12, 12, 12, 12, 30])
    esc_cols = ["C", "D", "E", "F"]
    esc_names = ["Base", "Precios", "Adquisición", "Combinado"]
    fila0 = 3
    res = {}
    for col_esc, nombre in zip(esc_cols, esc_names):
        rws, ln = bloque_fcf(E, fila0, col_esc=col_esc, titulo=f"Escenario: {nombre}")
        res[nombre] = (rws, ln)
        fila0 = int(ln["equity"][1:]) + 3
    E.seccion(fila0, "Resumen de escenarios")
    fila0 += 1
    for i, n in enumerate(esc_names):
        E.c(f"{'CDEF'[i]}{fila0}", n, tipo="t", negrita=True)
    resumen = [("EBITDA 2012", "EBITDA", None), ("Margen EBITDA 2012", "mEBITDA", None), ("FCF 2012", "FCF", None),
               ("Crecimiento a perpetuidad (g)", None, "g"), ("Valor de las operaciones (APV)", None, "v_ops"),
               ("Valor del patrimonio sin sinergias", None, "equity")]
    for etq, rowk, lnk in resumen:
        fila0 += 1
        E.texto(f"A{fila0}", etq)
        for i, n in enumerate(esc_names):
            rws, ln = res[n]
            ref = f"H{rws[rowk]}" if rowk else ln[lnk]
            fmt = F_PCT if rowk == "mEBITDA" else (F_PCT2 if lnk == "g" else F_MM)
            E.c(f"{'CDEF'[i]}{fila0}", f"={ref}", fmt, clave=f"esc_{lnk or rowk}_{i}")
    fila0 += 1
    E.texto(f"A{fila0}", "Δ patrimonio vs. base")
    for i in range(4):
        E.c(f"{'CDEF'[i]}{fila0}", f"={'CDEF'[i]}{fila0-1}-$C${fila0-1}", F_MM, clave=f"esc_delta_{i}")
    fila0 += 1
    E.texto(f"A{fila0}", "Control: escenario base de esta hoja − valor de la hoja Valoracion")
    E.c(f"C{fila0}", f"=C{fila0-2}-Valoracion!$C$33", F_MM, clave="chk_esc")

    # ------------------------------------------------------------------ Sensibilidad
    SE = Hoja(wb, "Sensibilidad", "Sensibilidad del valor del patrimonio sin sinergias (millones de US$)",
              [34, 12, 12, 12, 12, 12, 12, 12, 40])
    ws = SE.ws
    SE.texto("A3", "1. Ku (filas) × crecimiento a perpetuidad g (columnas). CPPC = Ku − t × Kd × D/V; FCF₂₀₁₃ = NOPAT₂₀₁₂(1+g)(1 − g/ROIC).", negrita=True)
    gs = [0.015, 0.02, 0.025, 0.03, 0.035, 0.04]
    kus = [0.0750, 0.0775, 0.0800, 0.0825, 0.0850, 0.0875, 0.0900]
    for j, gv in enumerate(gs):
        SE.c(f"{get_column_letter(2+j)}4", gv, F_PCT, tipo="i")
    for i, kv in enumerate(kus):
        r_ = 5 + i
        SE.c(f"A{r_}", kv, F_PCT2, tipo="i")
        for j in range(len(gs)):
            col = get_column_letter(2 + j)
            f13 = f"{n12}*(1+{col}$4)*(1-{col}$4/Proyeccion!$H${BR['ROIC']})"
            SE.c(f"{col}{r_}", f"=NPV($A{r_},{fcf_rng})+{pvts}+{f13}/(($A{r_}-{R['t']}*{R['kd']}*{R['L']})-{col}$4)/(1+$A{r_})^5+{noop}-{cl}", F_MM0)
    SE.texto("A12", "Base: Ku y g del modelo →")
    SE.c("B12", f"={R['ku']}", F_PCT2, tipo="v")
    SE.c("C12", f"={R['P_g']}", F_PCT2, tipo="v")
    SE.c("D12", "=Valoracion!$C$33", F_MM0, tipo="v")

    SE.texto("A14", "2. Múltiplo VE/EBITDA terminal (valor terminal por múltiplo, descontado a Ku)", negrita=True)
    mults = [4, 5, 6, 7, 8, 9, 10, 11]
    for j, m in enumerate(mults):
        col = get_column_letter(2 + j)
        SE.c(f"{col}15", m, F_X1, tipo="i")
        SE.c(f"{col}16", f"=Valoracion!$C$8+Valoracion!$C$10+Proyeccion!$H${BR['EBITDA']}*{col}15/(1+{R['ku']})^5+{noop}-{cl}", F_MM0)
    SE.texto("A15", "Múltiplo")
    SE.texto("A16", "Patrimonio")

    SE.texto("A18", "3. Descuento por iliquidez", negrita=True)
    descs = [("Ninguno (base)", "=0"), ("Costo de salir a bolsa", f"={R['ipo']}"), ("Anexo 8 (lectura)", f"={R['iliq_ex8']}"),
             ("Regla práctica", f"={R['iliq_rule']}"), ("Académico", f"={R['iliq_acad']}")]
    for j, (n, fm) in enumerate(descs):
        col = get_column_letter(2 + j)
        SE.c(f"{col}19", n, tipo="t", negrita=True)
        ws[f"{col}19"].alignment = Alignment(wrap_text=True)
        SE.c(f"{col}20", fm, F_PCT, tipo="v")
        SE.c(f"{col}21", f"=Valoracion!$C$33*(1-{col}20)", F_MM0, clave=f"iliq_val_{j}")
    SE.texto("A20", "Descuento")
    SE.texto("A21", "Patrimonio")

    SE.texto("G18", "5. Realización de las sinergias", negrita=True)
    for j, x in enumerate([0.0, 0.5, 1.0]):
        col = get_column_letter(8 + j)
        SE.c(f"{col}20", x, F_PCT, tipo="i")
        SE.c(f"{col}21", f"=Valoracion!$C$35+{col}20*{R['sin_100']}", F_MM0, clave=f"sin_real_{j}")
    SE.texto("G20", "Realizado")
    SE.texto("G21", "Patrimonio")
    SE.texto("A23", "4. Tornado: efecto de cada decisión sobre el patrimonio sin sinergias", negrita=True)
    for col, t in zip("BCDE", ["Bajo", "Alto", "Rango", "Supuesto"]):
        SE.c(f"{col}24", t, tipo="t", negrita=True)
    base = "Valoracion!$C$33"
    ku_ = R["ku"]
    wacc_ = R["wacc"]
    dg = R["torn_dg"]
    torn = [
        ("g ± 0,5 pp", eq_formula(ku_, wacc_, f"({g}-{dg})", f"{n12}*(1+{g}-{dg})*(1-({g}-{dg})/Proyeccion!$H${BR['ROIC']})"),
         eq_formula(ku_, wacc_, f"({g}+{dg})", f"{n12}*(1+{g}+{dg})*(1-({g}+{dg})/Proyeccion!$H${BR['ROIC']})"), "g del modelo ± 0,5 pp"),
        ("Método de beta (opciones 1 a 3)", f"=MIN(Costo_Capital!$F$24:$F$26)", f"=MAX(Costo_Capital!$F$24:$F$26)", "Hamada / Harris-Pringle"),
        ("PRM 4% - 6%", "=MIN(Costo_Capital!$F$27:$F$28)", "=MAX(Costo_Capital!$F$27:$F$28)", "Supuesto propio"),
        ("D/V objetivo (mediana a 50%)", "=MIN(Costo_Capital!$F$29:$F$31)", "=MAX(Costo_Capital!$F$29:$F$31)", "Anexo 7 y R&R"),
        ("Valor terminal: perpetuidad vs. múltiplo", f"=MIN({base},Valoracion!$C$46)", f"=MAX({base},Valoracion!$C$46)", "Anexo 7"),
        ("Afiliadas y minoritario: libros vs. P/U",
         f"={base}-Valoracion!$C$17+Historicos!$E${HR['INVAF']}+Valoracion!$C$24-Historicos!$E${HR['MINT']}", f"={base}", "Nota 3"),
        ("Descuento por iliquidez 0% - 25%", f"={base}*(1-{R['iliq_rule']})", f"={base}", "Caso, p. 8"),
        ("Escenario combinado de riesgo", f"=Escenarios!$F${int(R['esc_equity_3'].split('$')[-1])}", f"={base}", "Hoja Escenarios"),
    ]
    for i, (n, lo, hi, nota) in enumerate(torn):
        r_ = 25 + i
        SE.texto(f"A{r_}", n)
        SE.c(f"B{r_}", lo, F_MM0)
        SE.c(f"C{r_}", hi, F_MM0)
        SE.c(f"D{r_}", f"=ABS(C{r_}-B{r_})", F_MM0, clave=f"torn_{i}")
        SE.texto(f"E{r_}", nota)
        R[f"torn_lo_{i}"] = absref(ws, f"B{r_}")
        R[f"torn_hi_{i}"] = absref(ws, f"C{r_}")
    SE.texto("A34", "Variable con mayor rango")
    SE.c("B34", "=INDEX(A25:A32,MATCH(MAX(D25:D32),D25:D32,0))", None, clave="torn_max")

    # ------------------------------------------------------------------ Verificacion (se llena con Python) y Control
    VE = Hoja(wb, "Verificacion", "Verificación cruzada: Python independiente vs. Excel", [56, 16, 16, 14, 12])
    for col, t in zip("ABCDE", ["Cifra", "Excel", "Python", "Diferencia", "¿Coincide?"]):
        VE.c(f"{col}3", t, tipo="t", negrita=True)
    CT = Hoja(wb, "Control", "Tabla de control: cifra del informe → celda del Excel", [16, 70, 36, 16])
    for col, t in zip("ABCD", ["Clave", "Descripción", "Celda", "Valor"]):
        CT.c(f"{col}3", t, tipo="t", negrita=True)

    # Controles globales
    CT.texto("F3", "Controles de consistencia", negrita=True)
    checks = [("Subtotales del estado de resultados vs. publicados", "chk_is_max", 0.06),
              ("Balance histórico cuadra", "chk_bs", 0.001), ("Totales del balance vs. publicados", "chk_bs_pub", 0.06),
              ("Flujo histórico concilia con Δ caja", "chk_cf_hist", 0.001), ("Tabla 3 vs. recalculado", "chk_t3", 0.0006),
              ("D/V publicado vs. calculado", "chk_dv", 0.0001), ("Costo de clientes nuevos (Anexo 3)", "chk_a3", 0.5),
              ("ROE de la Tabla 2", "chk_t2roe", 0.0006), ("Saldos de deuda vs. Anexo 6", "chk_deuda", 0.6),
              ("Balance proforma cuadra", "chk_pf_bs", 0.001), ("Flujo proforma concilia", "chk_pf_cf", 0.001),
              ("Bloque de valoración = Valoracion", "chk_val", 0.001), ("Escenario base = Valoracion", "chk_esc", 0.001),
              ("Sinergias: total consistente", "chk_sin", 0.001)]
    for i, (n, k, tol) in enumerate(checks):
        r_ = 4 + i
        CT.texto(f"F{r_}", n)
        CT.c(f"G{r_}", f"={R[k]}", "0.0000", tipo="v")
        CT.c(f"H{r_}", f'=IF(ABS(G{r_})<={tol},"OK","REVISAR")')
    CT.texto(f"F{4+len(checks)}", "Todos los controles", negrita=True)
    CT.c(f"H{4+len(checks)}", f'=IF(COUNTIF(H4:H{3+len(checks)},"OK")={len(checks)},"OK","REVISAR")', negrita=True, clave="chk_all")
    CT.ws.column_dimensions["F"].width = 48
    CT.ws.column_dimensions["G"].width = 12

    wb.save(ruta)
    return R


if __name__ == "__main__":
    import json
    refs = construir(sys.argv[1])
    out = {k: v for k, v in refs.items() if isinstance(v, (str, int))}
    out["P_rows"] = refs["P_rows"]
    out["PF"] = refs["PF"]
    out["SR"] = refs["SR"]
    with open(sys.argv[1].replace(".xlsx", "_refs.json"), "w") as fh:
        json.dump(out, fh, indent=1, ensure_ascii=False)
    print("ok")
