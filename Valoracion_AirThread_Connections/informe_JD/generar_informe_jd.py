"""Informe de 2 páginas para la Junta Directiva, con las cifras del Excel del Grupo 3.

Cada cifra se lee de una celda del Excel del grupo (o se deriva de celdas, indicado en la
tabla de control). Uso: python generar_informe_jd.py ruta/Caso_2_Grupo_3_1.xlsx carpeta_salida
"""
import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from openpyxl import load_workbook

AZUL = "1F3864"
CELDAS = {
    "wacc": ("WACC_COMPARABLES", "K37"), "g": ("VALORACIÓN", "B5"), "ops": ("VALORACIÓN", "B8"),
    "tv_share": ("VALORACIÓN", "B9"), "noop": ("VALORACIÓN", "B16"), "claims": ("VALORACIÓN", "B25"),
    "eq": ("VALORACIÓN", "B26"), "sin_bh": ("VALORACIÓN", "B28"), "sin_rev": ("VALORACIÓN", "B29"),
    "eq_sin": ("VALORACIÓN", "B30"), "mult": ("VALORACIÓN", "B33"), "eq_mult": ("VALORACIÓN", "B36"),
    "g_impl": ("VALORACIÓN", "B37"), "ev_ebitda07": ("VALORACIÓN", "B38"), "eq_pe": ("VALORACIÓN", "B41"),
    "dv_ops": ("VALORACIÓN", "B42"), "apertura": ("VALORACIÓN", "B45"), "objetivo": ("VALORACIÓN", "B46"),
    "maximo": ("VALORACIÓN", "B47"), "afil": ("VALORACIÓN", "B13"), "afil_sh": ("VALORACIÓN", "B49"),
    "sinrev_sh": ("VALORACIÓN", "B50"), "prima": ("VALORACIÓN", "B51"), "mult_max": ("VALORACIÓN", "B52"),
    "m_ebitda": ("Exhibit 4", "J65"), "m_comp": ("WACC_COMPARABLES", "R13"), "rpm06": ("Exhibit 3", "F73"),
    "rpm07": ("Exhibit 3", "G73"), "cpa06": ("Exhibit 3", "F60"), "cpa07": ("Exhibit 3", "G60"),
    "deuda": ("Exhibit 6", "H8"), "linea11": ("PROYECCIONES", "F85"), "ahorro_bh12": ("SINERGIAS", "H8"),
    "ipo": ("PROYECCIONES", "K7"), "ebitda07": ("PROYECCIONES", "B26"),
}
DERIVADAS = {  # clave: (fórmula en celdas, función)
    "rpm_var": ("Exhibit 3!G73 / Exhibit 3!F73 − 1", lambda v: v["rpm07"] / v["rpm06"] - 1),
    "cpa_var": ("Exhibit 3!G60 / Exhibit 3!F60 − 1", lambda v: v["cpa07"] / v["cpa06"] - 1),
    "sin_tot": ("VALORACIÓN!B28 + VALORACIÓN!B29", lambda v: v["sin_bh"] + v["sin_rev"]),
    "mitad_rev": ("VALORACIÓN!B29 / 2", lambda v: v["sin_rev"] / 2),
}


def n(x, d=0):
    s = f"{abs(x):,.{d}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("−" if x < 0 else "") + s


def p(x, d=1):
    return n(x * 100, d) + "%"


def sombrear(celda, color):
    tcPr = celda._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color)
    tcPr.append(shd)


def bordes(celda, color="BFC9DC"):
    tcPr = celda._tc.get_or_add_tcPr()
    b = OxmlElement("w:tcBorders")
    for lado in ("top", "left", "bottom", "right"):
        e = OxmlElement(f"w:{lado}")
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), "4")
        e.set(qn("w:color"), color)
        b.append(e)
    tcPr.append(b)


def margen_celda(celda, cm=0.12):
    tcPr = celda._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for lado in ("left", "right"):
        e = OxmlElement(f"w:{lado}")
        e.set(qn("w:w"), str(int(cm * 567)))
        e.set(qn("w:type"), "dxa")
        mar.append(e)
    tcPr.append(mar)


def texto(par, partes, tam=10.5):
    """partes: lista de (texto, negrita)."""
    for t, b in partes:
        r = par.add_run(t)
        r.bold = b
        r.font.size = Pt(tam)
    return par


def main(ruta_xlsx, salida):
    wb = load_workbook(ruta_xlsx, data_only=True)
    v = {k: wb[h][c].value for k, (h, c) in CELDAS.items()}
    for k, val in v.items():
        assert isinstance(val, (int, float)), (k, val)
    for k, (_, f) in DERIVADAS.items():
        v[k] = f(v)

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.59), Cm(27.94)
    sec.left_margin = sec.right_margin = Cm(1.8)
    sec.top_margin, sec.bottom_margin = Cm(1.6), Cm(1.5)
    st = doc.styles["Normal"]
    st.font.name = "Arial"
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    st.font.size = Pt(9.5)
    st.paragraph_format.space_after = Pt(2)
    st.paragraph_format.space_before = Pt(0)
    st.paragraph_format.line_spacing = 1.05

    def parrafo(partes=None, tam=10.5, despues=4, alin=None):
        par = doc.add_paragraph()
        par.paragraph_format.space_after = Pt(despues)
        if alin:
            par.alignment = alin
        if partes:
            texto(par, partes, tam)
        return par

    def titulo_seccion(t):
        par = parrafo(despues=1.5)
        par.paragraph_format.space_before = Pt(10)
        par.paragraph_format.keep_with_next = True
        r = par.add_run(t)
        r.bold = True
        r.font.size = Pt(12.5)
        r.font.color.rgb = RGBColor.from_string(AZUL)
        pPr = par._p.get_or_add_pPr()
        bdr = OxmlElement("w:pBdr")
        bot = OxmlElement("w:bottom")
        bot.set(qn("w:val"), "single")
        bot.set(qn("w:sz"), "6")
        bot.set(qn("w:color"), AZUL)
        bdr.append(bot)
        pPr.append(bdr)

    def viñeta(partes, tam=10.5):
        par = doc.add_paragraph(style="List Bullet")
        par.paragraph_format.space_after = Pt(3)
        par.paragraph_format.left_indent = Cm(0.6)
        texto(par, partes, tam)
        return par

    def numerada(partes, tam=10.5):
        par = doc.add_paragraph(style="List Number")
        par.paragraph_format.space_after = Pt(4)
        par.paragraph_format.left_indent = Cm(0.6)
        texto(par, partes, tam)
        return par

    # Encabezado
    par = parrafo([("MEMORANDO A LA JUNTA DIRECTIVA · AMERICAN CABLE COMMUNICATIONS", True)], tam=8, despues=0)
    par.runs[0].font.color.rgb = RGBColor.from_string("666666")
    par = parrafo([("De: Desarrollo de Negocios (Grupo 3) · Diciembre de 2007 · Cifras en millones de US$", False)], tam=8, despues=4)
    par.runs[0].font.color.rgb = RGBColor.from_string("666666")
    par = parrafo(despues=4)
    r = par.add_run(f"Compra de AirThread: ofertar, pero no pagar más de US${n(v['maximo'])} millones")
    r.bold = True
    r.font.size = Pt(17)
    r.font.color.rgb = RGBColor.from_string(AZUL)

    # Recuadro de recomendación
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = t.cell(0, 0)
    sombrear(c, "E8EEF7")
    bordes(c, AZUL)
    pr = c.paragraphs[0]
    pr.paragraph_format.space_after = Pt(3)
    pr.paragraph_format.space_before = Pt(3)
    texto(pr, [("Recomendación: ", True),
               (f"presentar una oferta por el 100% de AirThread Connections (ATC). Abrir en US${n(v['apertura'])} millones, "
                f"buscar cerrar en US${n(v['objetivo'])} millones y retirarse si el precio supera US${n(v['maximo'])} millones. "
                "Ese tope paga lo que ATC vale por sí sola más el ahorro de backhaul que solo ACC puede generar con su fibra; "
                "no paga por adelantado las sinergias de ingresos, que son inciertas.", False)])
    parrafo(despues=0).paragraph_format.space_after = Pt(0)

    # 1. Situación
    titulo_seccion("1. El problema: cuánto pagar sin destruir valor")
    parrafo([("ACC necesita una oferta inalámbrica para competir con los operadores que venden paquetes. ATC le da esa oferta, "
              "pero el mayor riesgo de la operación es pagar de más. Los hechos que importan:", False)])
    viñeta([("ATC es rentable, pero está en desventaja. ", True),
            (f"Su margen EBITDA es {p(v['m_ebitda'])}, frente a {p(v['m_comp'])} de sus rivales. En 2007 su ingreso por minuto "
             f"cayó {p(-v['rpm_var'])} y el costo de captar un cliente subió {p(v['cpa_var'])}.", False)])
    viñeta([(f"Por sí sola vale US${n(v['eq'])} millones ", True),
            (f"(patrimonio, {n(v['ev_ebitda07'], 1)} veces su EBITDA de 2007). Es un flujo de caja descontado con WACC de "
             f"{p(v['wacc'], 2)} y crecimiento perpetuo de {p(v['g'], 2)}; el {p(v['tv_share'], 0)} del valor está después de 2012.", False)])
    viñeta([(f"Con ACC vale US${n(v['sin_tot'])} millones más. ", True),
            (f"US${n(v['sin_bh'])} millones salen del backhaul (usar la fibra de ACC ahorra US${n(v['ahorro_bh12'])} millones al año "
             f"en 2012) y US${n(v['sin_rev'])} millones de clientes empresariales: el {p(v['sinrev_sh'], 0)} de las sinergias, "
             "sin un margen documentado.", False)])
    viñeta([("El mercado sugiere precios mucho mayores. ", True),
            (f"Con el múltiplo mediano de los comparables ({n(v['mult'], 1)} veces el EBITDA), ATC valdría US${n(v['eq_mult'])} "
             f"millones, pero eso exige que crezca {p(v['g_impl'])} al año para siempre.", False)])
    viñeta([("La deuda propuesta por el banco es alta. ", True),
            (f"US${n(v['deuda'])} millones equivalen al {p(v['dv_ops'])} del valor de las operaciones (límite: 50%), y el flujo de "
             f"ATC no alcanza para las cuotas: en 2011 ya necesita US${n(v['linea11'])} millones de línea de crédito.", False)])

    # Tabla de valor
    filas = [("Valor de las operaciones (flujo de caja descontado)", v["ops"], False),
             ("+ Caja, afiliadas y nota por cobrar", v["noop"], False),
             ("− Deuda existente e interés minoritario", -v["claims"], False),
             ("Valor del patrimonio de ATC por sí sola", v["eq"], True),
             ("+ Sinergias de backhaul", v["sin_bh"], False),
             ("+ Sinergias de clientes empresariales", v["sin_rev"], False),
             ("Valor del patrimonio con todas las sinergias", v["eq_sin"], True)]
    t = doc.add_table(rows=len(filas) + 1, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    enc = [("Cómo se llega al valor", False), ("US$ millones", True)]
    for j, (txt, der) in enumerate(enc):
        cel = t.cell(0, j)
        sombrear(cel, AZUL)
        pr = cel.paragraphs[0]
        r = pr.add_run(txt)
        r.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(255, 255, 255)
        if der:
            pr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for i, (etq, val, neg) in enumerate(filas, start=1):
        for j, txt in enumerate([etq, n(val)]):
            cel = t.cell(i, j)
            bordes(cel)
            margen_celda(cel)
            if neg:
                sombrear(cel, "E8EEF7")
            pr = cel.paragraphs[0]
            pr.paragraph_format.space_after = Pt(0)
            r = pr.add_run(txt)
            r.bold = neg
            r.font.size = Pt(9.5)
            if j == 1:
                pr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for fila in t.rows:
        fila.cells[0].width = Cm(9.5)
        fila.cells[1].width = Cm(2.8)

    # 2. Alternativas
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    doc.paragraphs[-1].paragraph_format.space_after = Pt(0)
    titulo_seccion("2. Alternativas")
    cols = ["", "A. Pagar solo el valor de ATC", "B. Compartir solo el backhaul", "C. Pagar todas las sinergias",
            "D. No comprar: acuerdo de backhaul"]
    filas = [
        ("Precio", f"Hasta {n(v['eq'])}", f"{n(v['apertura'])} a {n(v['maximo'])}", f"Hasta {n(v['eq_sin'])}", "Sin compra"),
        ("Valor que conserva ACC", f"Todas las sinergias ({n(v['sin_tot'])})",
         f"Al menos las de ingresos ({n(v['sin_rev'])})", "Ninguno",
         f"Parte del ahorro de backhaul ({n(v['ahorro_bh12'])} al año en 2012)"),
        ("Implicaciones", "El vendedor, anclado en múltiplos de mercado, probablemente no la acepte.",
         f"Prima de {p(v['prima'])} sobre el valor de ATC ({n(v['mult_max'], 1)} veces el EBITDA). El vendedor cobra algo que no "
         "podría lograr por su cuenta.",
         f"Gana cualquier subasta, pero si solo se logra la mitad de las sinergias de ingresos, ACC pierde {n(v['mitad_rev'])}.",
         "No requiere deuda, pero ACC sigue sin oferta inalámbrica."),
        ("Riesgo principal", "Perder ATC frente a otro comprador.", "Que el vendedor no baje de los precios del mercado.",
         "Pagar de más: transfiere al vendedor todo el valor que crea ACC.", "Que un competidor compre ATC."),
    ]
    t = doc.add_table(rows=len(filas) + 1, cols=5)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    anchos = [Cm(2.6), Cm(3.4), Cm(3.7), Cm(3.4), Cm(3.4)]
    for j, txt in enumerate(cols):
        cel = t.cell(0, j)
        sombrear(cel, "2E7D32" if j == 2 else AZUL)
        margen_celda(cel)
        pr = cel.paragraphs[0]
        pr.paragraph_format.space_after = Pt(0)
        r = pr.add_run(txt + (" (recomendada)" if j == 2 else ""))
        r.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(255, 255, 255)
    for i, fila in enumerate(filas, start=1):
        for j, txt in enumerate(fila):
            cel = t.cell(i, j)
            bordes(cel)
            margen_celda(cel)
            if j == 2:
                sombrear(cel, "EAF4EA")
            elif j == 0:
                sombrear(cel, "F2F2F2")
            pr = cel.paragraphs[0]
            pr.paragraph_format.space_after = Pt(0)
            pr.paragraph_format.line_spacing = 1.0
            r = pr.add_run(txt)
            r.font.size = Pt(9)
            r.bold = j == 0
    for fila in t.rows:
        for j, w in enumerate(anchos):
            fila.cells[j].width = w

    # 3. Recomendación
    titulo_seccion(f"3. Recomendación: alternativa B, con tope de US${n(v['maximo'])} millones")
    numerada([("El precio es disciplinado. ", True),
              (f"US${n(v['maximo'])} millones equivalen a {n(v['mult_max'], 1)} veces el EBITDA de 2007 y a una prima de "
               f"{p(v['prima'])} sobre el valor de ATC. El P/U histórico de la industria da un valor similar "
               f"(US${n(v['eq_pe'])} millones); los múltiplos actuales de los comparables suponen un crecimiento que ATC no "
               "puede sostener con sus márgenes.", False)])
    numerada([("Solo se paga lo que ACC controla. ", True),
              ("El ahorro de backhaul depende de la fibra de ACC y es la sinergia más cierta. Las sinergias de clientes "
               f"empresariales (US${n(v['sin_rev'])} millones) quedan como margen de seguridad para ACC.", False)])
    numerada([("Estrategia de negociación. ", True),
              (f"Abrir en US${n(v['apertura'])} millones: el valor de ATC menos el {p(v['ipo'], 0)} que le costaría salir a bolsa, "
               f"que es la alternativa real del vendedor. Buscar US${n(v['objetivo'])} millones, que comparte la mitad del ahorro "
               "de backhaul. Si el vendedor pide más, ofrecer un pago adicional condicionado a que se cumplan las metas de "
               "clientes empresariales, en vez de subir el precio fijo.", False)])
    numerada([("Condiciones para firmar. ", True),
              (f"(a) Revisar las afiliadas, que aportan US${n(v['afil'])} millones ({p(v['afil_sh'], 0)} del valor). "
               "(b) Que la deuda no supere el 50% del valor: reducir el préstamo propuesto o aportar más capital, y garantizar "
               "la línea de crédito que ATC necesitará. "
               f"(c) Si el precio supera US${n(v['maximo'])} millones, retirarse y negociar con ATC un acuerdo comercial de "
               "backhaul (alternativa D).", False)])

    titulo_seccion("Referencias")
    parrafo([("Stafford, E. y Heilprin, J. (2011, rev. 2012). ", False), ("Valuation of AirThread Connections", False),
             (". Harvard Business School, caso 4263. Cifras: modelo financiero del Grupo 3 (Caso_2_Grupo_3_1.xlsx).", False)],
            tam=9)
    doc.paragraphs[-1].runs[1].italic = True

    salida = Path(salida)
    salida.mkdir(parents=True, exist_ok=True)
    doc.save(salida / "Informe_JD_AirThread.docx")

    # Tabla de control
    lineas = ["# Tabla de control: cifra del informe → celda del Excel del Grupo 3", "",
              "Archivo: `Caso_2_Grupo_3_1.xlsx`. Las cifras derivadas se calculan solo con celdas del archivo.", "",
              "| Clave | Celda | Valor |", "|---|---|---:|"]
    for k, (h, c) in CELDAS.items():
        lineas.append(f"| {k} | `{h}!{c}` | {n(v[k], 4)} |")
    for k, (f, _) in DERIVADAS.items():
        lineas.append(f"| {k} | {f} | {n(v[k], 4)} |")
    (salida / "tabla_control_JD.md").write_text("\n".join(lineas) + "\n", encoding="utf-8")
    print("ok")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
