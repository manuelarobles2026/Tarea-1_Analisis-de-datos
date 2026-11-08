"""Aplica las correcciones a la solución de Radio One y marca cada celda tocada.

Naranja = celda existente modificada. Verde = celda nueva.
Uso: python corregir.py <entrada.xlsx> <salida.xlsx>
"""
import sys
from copy import copy

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

SRC, DST = sys.argv[1], sys.argv[2]
MOD = PatternFill("solid", fgColor="FFFFC000")   # naranja
NEW = PatternFill("solid", fgColor="FF92D050")   # verde

wb = openpyxl.load_workbook(SRC)
log = []  # (hoja, celdas, tipo, antes, después, motivo)


def put(ws, ref, value, motivo, fmt_from=None, fmt=None, bold=None, log_it=True):
    c = ws[ref]
    before = c.value
    kind = "Modificada" if before not in (None, "") else "Nueva"
    if fmt_from:
        src = ws[fmt_from]
        c.number_format = src.number_format
        c.font = copy(src.font)
        c.alignment = copy(src.alignment)
        c.border = copy(src.border)
    if fmt:
        c.number_format = fmt
    if bold is not None:
        f = copy(c.font)
        f = Font(name=f.name, size=f.size, bold=bold, italic=f.italic, color=f.color)
        c.font = f
    c.value = value
    if c.font.color is not None and c.font.color.rgb == "FFFFFFFF":
        c.font = Font(name=c.font.name, size=c.font.size, bold=c.font.b, color="FF000000")
    c.fill = MOD if kind == "Modificada" else NEW
    if log_it:
        log.append((ws.title, ref, kind, before, value, motivo))


def put_row(ws, cols, row, values, motivo, fmt_from_row=None, label=None):
    """Escribe varias celdas de una fila y registra una sola entrada en el log."""
    befores = []
    kinds = set()
    for col, v in zip(cols, values):
        ref = f"{col}{row}"
        befores.append(ws[ref].value)
        kinds.add("Modificada" if ws[ref].value not in (None, "") else "Nueva")
        put(ws, ref, v, motivo, fmt_from=f"{col}{fmt_from_row}" if fmt_from_row else None, log_it=False)
    rng = f"{cols[0]}{row}:{cols[-1]}{row}"
    log.append((ws.title, rng, "/".join(sorted(kinds)), befores[0], values[0], motivo))


SUP = "'Supuestos y alcance'"
T = f"{SUP}!$B$6"

# ---------------------------------------------------------------- Supuestos
ws = wb["Supuestos y alcance"]
put(ws, "A10", "Beta del activo del sector (Anexo 8, promedio)",
    "Aclarar que 0,75 es beta del ACTIVO (ya desapalancada con beta de deuda 0,25), no una beta de Hamada.")
put(ws, "A20", "Beta de la deuda (Anexo 8, nota c)", "Supuesto del propio Anexo 8, necesario para reapalancar de forma consistente.", fmt_from="A19")
put(ws, "B20", 0.25, "Idem.", fmt_from="B10")
put(ws, "A21", "Año inicial del beneficio fiscal", "Nota b del Anexo 9: amortización deducible de 2001 a 2015.", fmt_from="A19")
put(ws, "B21", 2001, "Idem.", fmt_from="B19")

# ---------------------------------------------------------------- WACC
ws = wb["WACC"]
put(ws, "A15", "Beta del equity (fórmula Anexo 8)",
    "Se reemplaza Hamada por la fórmula con la que el caso construye las betas del activo.")
put(ws, "B15", f"=(B14-{SUP}!B20*B9)/B10",
    "βA = βE·E/V + βD·D/V  ⇒  βE = (βA − βD·D/V)/(E/V). Consistente con la definición del Anexo 8.")

# ---------------------------------------------------------------- FCL incremental
ws = wb["FCL incremental"]
put(ws, "A13", "Amortización fiscal (informativa)",
    "La amortización ya no entra al FCL; su ahorro fiscal dura sólo 2001-2015 y se valora por separado.")
put(ws, "A15", "Impuestos sin escudo (EBITDA × t)",
    "Flujo operativo con impuestos completos; así el valor terminal no perpetúa un escudo fiscal que vence en 2015.")
put_row(ws, "BCDEF", 15, [f"={c}11*{T}" for c in "BCDEF"],
        "Antes: MAX(0, EBIT·t), que (a) metía el escudo de amortización en el FCL y (b) perdía las pérdidas fiscales.")
put(ws, "A16", "NOPAT sin escudo", "Coherente con la nueva fila 15.")
put_row(ws, "BCDEF", 16, [f"={c}11-{c}15" for c in "BCDEF"], "NOPAT = EBITDA − impuestos operativos.")
put(ws, "A17", "(+) D&A (no aplica: parte del EBITDA)", "Evita doble conteo con la nueva definición de NOPAT.")
put_row(ws, "BCDEF", 17, [0, 0, 0, 0, 0], "Ídem.")
put(ws, "A25", "FCL incremental (2000 = sólo capital de trabajo)",
    "Aclarar el tratamiento del año 0.")
put(ws, "B25", "=-B23",
    "Antes: B16+B17-B18-B23 incluía el NOPAT de 2000 (~US$41 M), que genera Clear Channel antes del cierre. "
    "En t0 Radio One sólo aporta el capital de trabajo inicial.")
put(ws, "B31",
    "En t0 (2000) sólo se reconoce el desembolso de capital de trabajo inicial (lo exige el caso). "
    "El flujo operativo de 2000 queda para el vendedor. Las filas 15-17 calculan el flujo con impuestos completos "
    "sobre el EBITDA; el escudo fiscal de la amortización (2001-2015) se suma aparte en 'Valoracion DCF'.",
    "Nota actualizada con la nueva lógica.")

# ---------------------------------------------------------------- Valoración DCF
ws = wb["Valoracion DCF"]
W = "WACC!$B$21"
put(ws, "A10", "Escudo fiscal anual (90.000 × t)",
    "Nota b del Anexo 9: amortización deducible de ~US$90 M al año entre 2001 y 2015.", fmt_from="A9")
put(ws, "B10", f"={SUP}!B16*{SUP}!B6", "Ídem.", fmt_from="B9")
put(ws, "A11", "VP escudo fiscal 2001-2015", "Se valora el escudo sólo por los 15 años en que existe.",
    fmt_from="A9")
put(ws, "B11", f"=B10*(1-(1+{W})^-({SUP}!B19-{SUP}!B21+1))/{W}", "Anualidad de 15 pagos descontada al WACC.",
    fmt_from="B9")
put(ws, "A18", "Valor empresa perpetuidad + escudo", "Se suma el VP del escudo fiscal.")
put(ws, "B18", "=SUM(C9:F9)+B17+B9+B11", "Antes no sumaba B11 (el escudo estaba, mal, dentro del FCL perpetuo).")
put(ws, "C18", "=B18/'FCL incremental'!$B$8", "Múltiplo implícito sobre FCE 2000, comparable con precedentes.",
    fmt="0.0\"x FCE 2000\"")
put(ws, "A24", "Valor empresa múltiplo + escudo", "Se suma el VP del escudo fiscal.")
put(ws, "B24", "=SUM(C9:F9)+B23+B9+B11",
    "Antes no sumaba B11. Ojo: los múltiplos de mercado pueden incorporar parte de escudos fiscales; ver Registro de cambios.")
put(ws, "C24", "=B24/'FCL incremental'!$B$8", "Múltiplo implícito sobre FCE 2000.", fmt="0.0\"x FCE 2000\"")
put(ws, "A27", "Sensibilidad DCF (perpetuidad + escudo)", "Título actualizado.")
for r in range(29, 34):
    vals = []
    for col in "BCDEF":
        w = f"{col}$28"
        vals.append(
            f"=$B$9+SUMPRODUCT($C$6:$F$6,1/(1+{w})^$C$7:$F$7)+($F$6*(1+$A{r})/({w}-$A{r}))/(1+{w})^4"
            f"+$B$10*(1-(1+{w})^-({SUP}!$B$19-{SUP}!$B$21+1))/{w}")
    put_row(ws, "BCDEF", r, vals, "Se agrega el VP del escudo fiscal a cada celda de la sensibilidad.")

# ---------------------------------------------------------------- Múltiplos
ws = wb["Multiplos y transacciones"]
E8 = "'Exhibit 8'"
put(ws, "C7", f"=MEDIAN({E8}!C3:C16)", "Se calcula con fórmula (antes 16,9 digitado; mismo valor).")
put(ws, "F7", f"=MEDIAN({E8}!D3:D16)", "Antes 18,2 digitado; la mediana real del Anexo 8 es 17,8.")
put(ws, "A7", "Mediana sector", "Ya no es aproximada.")
put(ws, "A9", "Promedio sin ROIA ni HBCCA", "Excluye al comprador y al valor atípico (42x).", fmt_from="A7")
put(ws, "B9", "='FCL incremental'!C8", "Ídem.", fmt_from="B7")
put(ws, "C9", f"=(SUM({E8}!C3:C16)-{E8}!C12-{E8}!C14)/12", "Ídem.", fmt_from="C7")
put(ws, "D9", "=B9*C9", "Ídem.", fmt_from="D7")
put(ws, "E9", "='FCL incremental'!C11", "Ídem.", fmt_from="E7")
put(ws, "F9", f"=(SUM({E8}!D3:D16)-{E8}!D12-{E8}!D14)/12", "Ídem.", fmt_from="F7")
put(ws, "G9", "=E9*F9", "Ídem.", fmt_from="G7")
put(ws, "A24", "Oferta - piso (DCF perpetuidad)",
    "El caso define el piso como el valor del flujo de caja de las emisoras por separado.")
put(ws, "B24", "=B19", "Antes: MIN de los cinco métodos.")
put(ws, "A25", "Oferta - techo / precio de retiro",
    "Techo anclado al valor para Radio One, no a una mediana arbitraria.")
put(ws, "B25", "=MIN(B20,B23)",
    "Antes: mediana de los 5 métodos. Ahora: el menor entre el DCF con múltiplo terminal y el precedente más alto (Infinity 21,5x).")
put(ws, "C24", "=B24/'FCL incremental'!$B$8", "Múltiplo implícito.", fmt="0.0\"x FCE 2000\"")
put(ws, "C25", "=B25/'FCL incremental'!$B$8", "Múltiplo implícito.", fmt="0.0\"x FCE 2000\"")
put(ws, "A26", "Oferta de apertura (20x FCE)",
    "Punto de partida competitivo dentro del rango.", fmt_from="A24")
put(ws, "B26", "=D15", "Ídem.", fmt_from="B24")
put(ws, "C26", "=B26/'FCL incremental'!$B$8", "Ídem.", fmt="0.0\"x FCE 2000\"")
put(ws, "A27", "Ref. no recomendada: 30x FCE",
    "Máximo teórico del caso si se paga con acciones; incorpora expectativas y genera circularidad.", fmt_from="A24")
put(ws, "B27", "=30*B13", "Ídem.", fmt_from="B24")
put(ws, "C27", "=B27/'FCL incremental'!$B$8", "Ídem.", fmt="0.0\"x FCE 2000\"")

# ---------------------------------------------------------------- Resumen ejecutivo
ws = wb["Resumen ejecutivo"]
put(ws, "A13", "Rango de oferta - piso (DCF perpetuidad)", "Etiqueta alineada con la nueva definición.")
put(ws, "A14", "Rango de oferta - techo (precio de retiro)", "Ídem.")
put(ws, "A15", "Oferta de apertura sugerida (20x FCE)", "Nuevo indicador.", fmt_from="A14")
put(ws, "B15", "='Multiplos y transacciones'!B26", "Ídem.", fmt_from="B14")
put(ws, "A16", "VP escudo fiscal 2001-2015 (incluido arriba)", "Nuevo indicador.", fmt_from="A14")
put(ws, "B16", "='Valoracion DCF'!B11", "Ídem.", fmt_from="B14")
put(ws, "A18",
    "La venta forzada de activos urbanos de Clear Channel es una oportunidad poco frecuente y encaja con la estrategia de "
    "Radio One. Corregido el modelo (sin el flujo de 2000 y con el escudo fiscal valorado sólo para 2001-2015), las "
    "emisoras valen unos US$875 M tal como operan (≈14x FCE 2000) y hasta ~US$1.337 M para Radio One con un múltiplo "
    "terminal de salida (≈21,5x, igual a Infinity). Se recomienda abrir en 20x FCE 2000 (~US$1.244 M) y no superar "
    "~21,5x. Pagar más sólo se justifica con mejoras operativas verificables; el 30x de la propia acción no es un "
    "ancla válida porque refleja la expectativa de esta compra.",
    "Conclusión actualizada con los valores corregidos.")

# ---------------------------------------------------------------- Análisis estratégico
ws = wb["Analisis estrategico"]
put(ws, "A22",
    "Ofertar entre el piso (DCF perpetuidad, ~14x FCE 2000) y el precio de retiro (~21,5x, DCF con múltiplo terminal y "
    "precedente Infinity), abriendo en 20x FCE 2000, condicionado a due diligence por mercado, confirmación del "
    "beneficio fiscal de US$90 M anuales (2001-2015), retención del equipo comercial y un límite explícito de "
    "apalancamiento. Evitar justificar el precio con 30x, porque ese múltiplo incorpora expectativas del propio "
    "negocio y genera circularidad.",
    "Recomendación alineada con el rango corregido.")

# ---------------------------------------------------------------- Registro de cambios
rc = wb.create_sheet("Registro de cambios", 1)
rc["A1"] = "Registro de cambios sobre la solución original"
rc["A1"].font = Font(bold=True, size=14)
rc["A3"] = "Leyenda:"
rc["A3"].font = Font(bold=True)
rc["B3"] = "Celda modificada"
rc["B3"].fill = MOD
rc["C3"] = "Celda nueva"
rc["C3"].fill = NEW
rc["A5"] = ("Correcciones de fondo: (1) el FCL de 2000 sólo incluye la inversión en capital de trabajo; "
            "(2) el flujo operativo paga impuestos completos y el escudo fiscal de la amortización se valora aparte "
            "sólo para 2001-2015; (3) la beta se reapalanca con la fórmula del Anexo 8 (beta de deuda 0,25). "
            "Además: mediana EBITDA corregida, promedio del sector sin atípicos, rango de oferta redefinido. "
            "Advertencia: el DCF con múltiplo terminal suma el escudo fiscal completo; si se considera que el "
            "múltiplo de 15x EBITDA ya incorpora escudos fiscales, ese valor estaría algo sobrestimado.")
rc["A5"].alignment = Alignment(wrap_text=True, vertical="top")
rc.merge_cells("A5:F5")
rc.row_dimensions[5].height = 95
hdr = ["Hoja", "Celda(s)", "Tipo", "Antes", "Después", "Motivo"]
for i, h in enumerate(hdr, 1):
    c = rc.cell(row=7, column=i, value=h)
    c.font = Font(bold=True, color="FFFFFFFF")
    c.fill = PatternFill("solid", fgColor="FF17365D")
for r, (sh, ref, kind, before, after, why) in enumerate(log, 8):
    vals = [sh, ref, kind, "" if before is None else str(before), str(after), why]
    for i, v in enumerate(vals, 1):
        c = rc.cell(row=r, column=i, value=v)
        c.data_type = "s"  # mostrar fórmulas como texto
        c.alignment = Alignment(wrap_text=True, vertical="top")
    rc.cell(row=r, column=2).fill = NEW if kind == "Nueva" else MOD
for col, w in zip("ABCDEF", [24, 12, 12, 40, 50, 60]):
    rc.column_dimensions[col].width = w

wb.calculation.fullCalcOnLoad = True
wb.save(DST)
print(f"{len(log)} entradas en el registro")
