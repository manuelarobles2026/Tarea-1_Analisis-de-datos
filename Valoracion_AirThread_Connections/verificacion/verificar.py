"""Verificación independiente del modelo de AirThread.

Recalcula desde los datos del caso (sin leer fórmulas del Excel) las cifras clave,
las compara con los valores recalculados del Excel y escribe la hoja 'Verificacion'.

Uso: python verificar.py ../modelo/modelo_airthread.xlsx
"""
import sys

import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import Font

# ---------------------------------------------------------------- datos del caso
t, kd, spread, erp = 0.40, 0.055, 0.0125, 0.05
rf = kd - spread
S07 = 3679.237
hist = {  # 2005, 2006, 2007
    "S": [2827.022, 3214.41, 3679.237], "EQ": [203.743, 258.745, 267.027],
    "SYS": [604.093, 639.683, 717.075], "COGS": [511.939, 568.903, 640.225], "SGA": [1217.709, 1399.561, 1555.639],
    "DA": [490.093, 555.525, 582.269], "AR": [362.359, 407.438, 435.497], "INV": [92.748, 117.189, 100.99],
    "PREP": [32.068, 34.955, 41.588], "AP": [254.116, 254.871, 260.791], "DEF": [111.407, 123.344, 143.445],
    "ACC": [42.865, 47.842, 59.224],
}
h = {k: np.array(v) for k, v in hist.items()}
REVh = h["S"] + h["EQ"]
COh = h["SYS"] + h["COGS"] + h["SGA"]
EBITDA07 = REVh[2] - COh[2]
EBIT07 = EBITDA07 - h["DA"][2]
NI07 = 314.734
eq_ratio = np.mean(h["EQ"] / h["S"])
dso = np.mean(h["AR"] / REVh * 360)
dio = np.mean(h["INV"] / h["EQ"] * 360)
prep = np.mean(h["PREP"] / COh)
dpo = np.mean(h["AP"] / COh * 360)
ddif = np.mean(h["DEF"] / h["S"] * 360)
dacum = np.mean(h["ACC"] / COh * 360)

g_serv = np.array([0.14, 0.14, 0.125, 0.10, 0.07])
capex_pct = np.array([0.14, 0.14, 0.15, 0.1525, 0.155])
DA_ex1 = np.array([705.229319748866, 803.961424513707, 867.438753154676, 922.376540854473, 952.910384965518])

comps = np.array([  # E, D neta, beta, EBITDA
    [118496.867, 69130, 0.86, 16949], [189469.94865, 79351.251, 0.89, 14099], [21078.873, 5080.008393, 1.17, 9914],
    [26285.354631, 8335.0859534901, 0.97, 12614.1213294], [7359.89929668, 3267.79528772592, 1.13, 1028.4766589]])
E, Dn, be, ebitda_c = comps.T
DV = Dn / (E + Dn)
DE = Dn / E
bd = (kd - rf) / erp
bu_hamada = np.mean(be / (1 + (1 - t) * DE))
bu_hp0 = np.mean(be * (1 - DV))
bu_hp = np.mean(be * (1 - DV) + bd * DV)
ku = rf + bu_hp * erp
L = np.mean(DV)
wacc_formula = (1 - L) * (rf + (bu_hp + (bu_hp - bd) * L / (1 - L)) * erp) + L * kd * (1 - t)
wacc = ku - t * kd * L  # Harris-Pringle
mult_med = np.median((E + Dn) / ebitda_c)


# ---------------------------------------------------------------- deuda (bucle mensual propio)
D0 = EBITDA07 / 5.0 / kd
r_m = kd / 12
pago = D0 * r_m / (1 - (1 + r_m) ** -120)
bal = D0
int_anual = np.zeros(5)
princ_anual = np.zeros(5)
for m in range(1, 61):
    i = bal * r_m
    p = pago - i
    int_anual[(m - 1) // 12] += i
    princ_anual[(m - 1) // 12] += p
    bal -= p
bullet = bal
pv_ts = sum(int_anual[k] * t / (1 + kd) ** (k + 1) for k in range(5))


def proyectar(dg=0.0, dsga=0.0):
    S = S07 * np.cumprod(1 + g_serv + dg)
    EQr = S * eq_ratio
    REV = S + EQr
    SYS, COGS, SGA = 0.2 * S, 2.4 * EQr, (0.4 + dsga) * REV
    EBITDA = REV - SYS - COGS - SGA
    S_base = S07 * np.cumprod(1 + g_serv)
    DA = REV * DA_ex1 / (S_base * (1 + eq_ratio))
    EBIT = EBITDA - DA
    NOPAT = EBIT * (1 - t)
    CO = SYS + COGS + SGA
    NWC = REV * dso / 360 + EQr * dio / 360 + CO * prep - CO * dpo / 360 - S * ddif / 360 - CO * dacum / 360
    NWC07 = 435.497 + 100.99 + 41.588 - 260.791 - 143.445 - 59.224
    dNWC = np.diff(np.concatenate([[NWC07], NWC]))
    CAPEX = REV * capex_pct
    FCF = NOPAT + DA - CAPEX - dNWC
    IC = 1404.094 + 1792.062 + 1002.293 + np.cumsum(CAPEX - DA + dNWC)
    IC_ini = np.concatenate([[1404.094 + 1792.062 + 1002.293], IC[:-1]])
    roic = NOPAT[-1] / IC_ini[-1]
    rr = (CAPEX[-1] - DA[-1] + dNWC[-1]) / NOPAT[-1]
    g = roic * rr
    fcf13 = NOPAT[-1] * (1 + g) * (1 - g / roic)
    tv = fcf13 / (wacc - g)
    disc_ku = (1 + ku) ** -np.arange(1, 6)
    v_ops = FCF @ disc_ku + tv * disc_ku[-1] + pv_ts
    return dict(S=S, REV=REV, EBITDA=EBITDA, EBIT=EBIT, NOPAT=NOPAT, FCF=FCF, dNWC=dNWC, CAPEX=CAPEX, DA=DA,
                roic=roic, rr=rr, g=g, tv=tv, v_ops=v_ops, fcf13=fcf13, disc=disc_ku, SYS=SYS)


b = proyectar()
nooper = 204.533 + 16.352 + 90.033 * 19.095 + 4.422
claims = 1002.293 + 15.056 * 19.095
equity = b["v_ops"] + nooper - claims
tv_mult = b["EBITDA"][-1] * mult_med
eq_mult = (b["FCF"] @ b["disc"]) + pv_ts + tv_mult * b["disc"][-1] + nooper - claims

# g implícito por bisección (independiente de la fórmula cuadrática del Excel)
lo, hi = 0.0, wacc - 1e-6
f = lambda g: b["NOPAT"][-1] * (1 + g) * (1 - g / b["roic"]) / (wacc - g) - tv_mult  # noqa: E731
for _ in range(200):
    mid = (lo + hi) / 2
    lo, hi = (mid, hi) if f(mid) < 0 else (lo, mid)
g_impl = (lo + hi) / 2

# sinergias
bh = b["SYS"] * 0.2 * np.array([0, 0.07, 0.12, 0.222, 0.30])
brev = np.array([0.3, 0.5, 0.7, 1.0, 1.2]) * np.array([859, 885, 911, 939, 967]) * 0.0506 * 12
margen = b["EBITDA"] / b["REV"]


def vp_sin(x):
    x = x * (1 - t)
    return x @ b["disc"] + x[-1] * (1 + b["g"]) / (ku - b["g"]) * b["disc"][-1]


sin_bh, sin_rev = vp_sin(bh), vp_sin(brev * margen)

# escenario combinado
c = proyectar(dg=-0.03, dsga=0.02)
eq_comb = c["v_ops"] + nooper - claims

# proforma: línea de crédito al cierre de 2012
caja, rev_ = 0.0, 0.0
for k in range(5):
    ni = (b["EBIT"][k] - int_anual[k] - rev_ * kd) * (1 - t)
    disp = caja - rev_ + ni + b["DA"][k] - b["dNWC"][k] - b["CAPEX"][k] - princ_anual[k]
    caja, rev_ = max(0, disp), max(0, -disp)
revolver12 = rev_

# WACC-DCF a estructura constante
disc_w = (1 + wacc) ** -np.arange(1, 6)
eq_wacc = b["FCF"] @ disc_w + b["tv"] * disc_w[-1] + nooper - claims

py = {
    "rf": rf, "beta_u_hamada": bu_hamada, "beta_u_hp0": bu_hp0, "beta_u": bu_hp, "ku": ku, "wacc": wacc,
    "wacc_formula_completa": wacc_formula, "d0": D0, "cuota": pago, "int_2008": int_anual[0], "bullet": bullet,
    "pv_ts": pv_ts, "rev_2012": b["REV"][-1], "ebitda_2012": b["EBITDA"][-1], "fcf_2008": b["FCF"][0],
    "fcf_2012": b["FCF"][-1], "roic_2012": b["roic"], "g": b["g"], "tv": b["tv"], "pv_fcf": b["FCF"] @ b["disc"],
    "v_ops": b["v_ops"], "nooper": nooper, "claims": claims, "equity": equity, "sin_bh": sin_bh, "sin_rev": sin_rev,
    "eq_sin": equity + sin_bh + sin_rev, "mult_med": mult_med, "eq_mult": eq_mult, "g_impl": g_impl,
    "eq_comb": eq_comb, "revolver_2012": revolver12, "eq_wacc": eq_wacc, "dv0_ops": D0 / b["v_ops"],
}

# celdas del Excel con la misma cifra
excel_ref = {
    "rf": ("Supuestos", "C9"), "beta_u_hamada": ("Comparables", "C28"), "beta_u_hp0": ("Comparables", "D28"),
    "beta_u": ("Costo_Capital", "C9"), "ku": ("Costo_Capital", "C10"), "wacc": ("Costo_Capital", "C17"),
    "wacc_formula_completa": ("Costo_Capital", "C16"), "d0": ("Deuda", "C5"), "cuota": ("Deuda", "C6"),
    "int_2008": ("Deuda", "L9"), "bullet": ("Deuda", "P11"), "pv_ts": ("Valoracion", "C10"),
    "rev_2012": ("Proyeccion", "H8"), "ebitda_2012": ("Proyeccion", "H12"), "fcf_2008": ("Proyeccion", "D28"),
    "fcf_2012": ("Proyeccion", "H28"), "roic_2012": ("Proyeccion", "H30"), "g": ("Valoracion", "C7"),
    "tv": ("Proyeccion", "H36"), "pv_fcf": ("Valoracion", "C8"), "v_ops": ("Valoracion", "C11"),
    "nooper": ("Valoracion", "C20"), "claims": ("Valoracion", "C26"), "equity": ("Valoracion", "C33"),
    "sin_bh": ("Valoracion", "C36"), "sin_rev": ("Valoracion", "C37"), "eq_sin": ("Valoracion", "C38"),
    "mult_med": ("Comparables", "P13"), "eq_mult": ("Valoracion", "C46"), "g_impl": ("Valoracion", "C47"),
    "eq_comb": ("Escenarios", "F186"), "revolver_2012": ("Proyeccion", "H70"), "eq_wacc": ("Valoracion", "C49"),
    "dv0_ops": ("Valoracion", "C55"),
}

if __name__ == "__main__":
    ruta = sys.argv[1]
    vals = load_workbook(ruta, data_only=True)
    filas = []
    for k, v in py.items():
        hoja, celda = excel_ref[k]
        ex = vals[hoja][celda].value
        filas.append((k, hoja, celda, ex, v, abs(ex - v)))
        print(f"{k:24s} {hoja}!{celda:6s} excel={ex:14.6f} python={v:14.6f} dif={abs(ex - v):.2e}")
    peor = max(f[-1] for f in filas)
    print("diferencia máxima:", peor)
    wb = load_workbook(ruta)
    ws = wb["Verificacion"]
    for i, (k, hoja, celda, ex, v, d) in enumerate(filas):
        r = 4 + i
        ws[f"A{r}"] = f"{k} ({hoja}!{celda})"
        ws[f"B{r}"] = f"='{hoja}'!{celda}"
        ws[f"C{r}"] = float(v)
        ws[f"C{r}"].font = Font(name="Arial", size=10, color="0000FF")
        ws[f"D{r}"] = f"=ABS(B{r}-C{r})"
        ws[f"E{r}"] = f'=IF(D{r}<0.0001,"Sí","NO")'
        for col in "ABDE":
            ws[f"{col}{r}"].font = Font(name="Arial", size=10)
        for col in "BCD":
            ws[f"{col}{r}"].number_format = "#,##0.000000"
    r = 5 + len(filas)
    ws[f"A{r}"] = "Valores de la columna C: calculados con verificacion/verificar.py (implementación independiente en Python)."
    ws[f"A{r}"].font = Font(name="Arial", size=9, italic=True)
    ws[f"A{r+1}"] = "Todas coinciden (diferencia máxima en la columna D)"
    ws[f"D{r+1}"] = f"=MAX(D4:D{3+len(filas)})"
    ws[f"D{r+1}"].number_format = "0.0E+00"
    ws[f"E{r+1}"] = f'=IF(COUNTIF(E4:E{3+len(filas)},"Sí")={len(filas)},"Sí","NO")'
    wb.save(ruta)
