"""Cálculos independientes del caso Nelson Nurseries (miles de USD).

Sirve para verificar el Excel: los mismos números se calculan aquí en Python
y en el libro con fórmulas, y main compara ambos.
"""

ANIOS = [2016, 2017, 2018, 2019]

# Exhibit 1 del caso
H = {
    "rev":   [788.5, 807.6, 908.2, 1048.8],
    "cogs":  [402.9, 428.8, 437.7, 503.4],
    "sga":   [301.2, 302.0, 356.0, 404.5],
    "dep":   [34.2, 38.4, 36.3, 40.9],
    "tax":   [17.6, 13.1, 26.2, 39.2],
    "cash":  [120.1, 105.2, 66.8, 9.4],
    "ar":    [90.6, 99.5, 119.5, 146.4],
    "inv":   [468.3, 507.6, 523.4, 656.9],
    "oca":   [20.9, 19.3, 22.6, 20.9],
    "nfa":   [332.1, 332.5, 384.3, 347.9],
    "ap":    [6.0, 5.3, 4.5, 5.0],
    "wp":    [19.7, 22.0, 22.1, 24.4],
    "op":    [10.2, 15.4, 16.6, 17.9],
    "capex": [22.0, 38.8, 88.1, 4.5],
    "purch": [140.8, 145.2, 161.2, 185.1],
}
# Exhibit 2 (publicado), para comprobar
EXH2 = {
    "crec": [0.029, 0.024, 0.125, 0.155], "mb": [.489, .469, .518, .520],
    "mo": [.064, .048, .086, .095], "mn": [.041, .031, .057, .058],
    "roa": [.032, .024, .047, .051], "roe": [.033, .025, .048, .054],
    "dso": [41.9, 45.0, 48.0, 50.9], "dio": [424.2, 432.1, 436.5, 476.3],
    "dpo": [15.6, 13.3, 10.2, 9.9], "rotaf": [2.4, 2.4, 2.4, 3.0],
}
BENCH = {"crec": -.018, "mb": .489, "mo": .076, "mn": .028, "roa": .029,
         "roe": .040, "dso": 21.8, "dio": 386.3, "dpo": 26.9, "rotaf": 2.7}


def historico():
    r = {k: list(v) for k, v in H.items()}
    n = range(4)
    r["gp"] = [r["rev"][i] - r["cogs"][i] for i in n]
    r["ebit"] = [r["gp"][i] - r["sga"][i] - r["dep"][i] for i in n]
    r["ni"] = [r["ebit"][i] - r["tax"][i] for i in n]
    r["ca"] = [r["cash"][i] + r["ar"][i] + r["inv"][i] + r["oca"][i] for i in n]
    r["ta"] = [r["ca"][i] + r["nfa"][i] for i in n]
    r["cl"] = [r["ap"][i] + r["wp"][i] + r["op"][i] for i in n]
    r["eq"] = [r["ta"][i] - r["cl"][i] for i in n]
    return r


def indicadores(r):
    n = range(4)
    k = {}
    k["crec"] = [None] + [r["rev"][i] / r["rev"][i - 1] - 1 for i in range(1, 4)]
    k["mb"] = [r["gp"][i] / r["rev"][i] for i in n]
    k["mo"] = [r["ebit"][i] / r["rev"][i] for i in n]
    k["mn"] = [r["ni"][i] / r["rev"][i] for i in n]
    k["tasa"] = [r["tax"][i] / r["ebit"][i] for i in n]
    k["sga_pct"] = [r["sga"][i] / r["rev"][i] for i in n]
    k["roa"] = [r["ni"][i] / r["ta"][i] for i in n]
    k["roe"] = [r["ni"][i] / r["eq"][i] for i in n]
    k["dso"] = [r["ar"][i] / r["rev"][i] * 365 for i in n]
    k["dio"] = [r["inv"][i] / r["cogs"][i] * 365 for i in n]
    k["dpo"] = [r["ap"][i] / r["purch"][i] * 365 for i in n]
    k["ccc"] = [k["dso"][i] + k["dio"][i] - k["dpo"][i] for i in n]
    k["rotaf"] = [r["rev"][i] / r["nfa"][i] for i in n]
    k["rotat"] = [r["rev"][i] / r["ta"][i] for i in n]
    k["cash_pct"] = [r["cash"][i] / r["rev"][i] for i in n]
    k["liq"] = [r["ca"][i] / r["cl"][i] for i in n]
    k["acida"] = [(r["cash"][i] + r["ar"][i]) / r["cl"][i] for i in n]
    k["ktno"] = [r["ar"][i] + r["inv"][i] + r["oca"][i] - r["cl"][i] for i in n]
    k["ktno_pct"] = [k["ktno"][i] / r["rev"][i] for i in n]
    return k


def flujos(r):
    """Flujo de caja indirecto 2017-2019 y conciliación con la caja."""
    out = []
    for i in range(1, 4):
        d = lambda c: r[c][i] - r[c][i - 1]
        dktno = d("ar") + d("inv") + d("oca") - d("cl")
        cfo = r["ni"][i] + r["dep"][i] - dktno
        fcf = cfo - r["capex"][i]
        out.append(dict(anio=ANIOS[i], ni=r["ni"][i], dep=r["dep"][i], d_ar=d("ar"),
                        d_inv=d("inv"), d_oca=d("oca"), d_cl=d("cl"), dktno=dktno,
                        cfo=cfo, capex=r["capex"][i], fcf=fcf, d_cash=d("cash")))
    return out


def proyeccion(r, g=0.30, dso=None, dio=None, sga_fijo=0.0, caja_obj=0.08,
               capex=75.0, dep=46.0, tasa_int=0.065, d_cogs=0.0, dpo=None, g_inv=None):
    """Proyección 2020 con los ratios de 2019 salvo lo que se cambie.

    g_inv: crecimiento para el que se arma el inventario y las compras (por
    defecto = g). d_cogs: cambio en puntos del costo de ventas sobre ventas.
    La deuda paga `tasa_int` sobre el saldo de cierre; se resuelve en forma
    cerrada para evitar la circularidad: D = AFN_sin_interes / (1 - i*(1-t)).
    Si D < 0, es excedente de caja.
    """
    i9 = 3
    g_inv = g if g_inv is None else g_inv
    rev = r["rev"][i9] * (1 + g)
    cogs_pct = r["cogs"][i9] / r["rev"][i9] + d_cogs
    cogs = rev * cogs_pct
    sga = sga_fijo + rev * (r["sga"][i9] - sga_fijo) / r["rev"][i9]
    ebit = rev - cogs - sga - dep
    t = r["tax"][i9] / r["ebit"][i9]
    dso = r["ar"][i9] / r["rev"][i9] * 365 if dso is None else dso
    dio = r["inv"][i9] / r["cogs"][i9] * 365 if dio is None else dio
    dpo = r["ap"][i9] / r["purch"][i9] * 365 if dpo is None else dpo
    purch = r["purch"][i9] * (1 + g_inv)
    p = dict(rev=rev, cogs=cogs, gp=rev - cogs, sga=sga, dep=dep, ebit=ebit, t=t)
    p["cash_obj"] = caja_obj * rev
    p["ar"] = rev * dso / 365
    p["inv"] = r["rev"][i9] * (1 + g_inv) * cogs_pct * dio / 365
    p["oca"] = rev * r["oca"][i9] / r["rev"][i9]
    p["nfa"] = r["nfa"][i9] + capex - dep
    p["ap"] = purch * dpo / 365
    p["wp"] = rev * r["wp"][i9] / r["rev"][i9]
    p["op"] = rev * r["op"][i9] / r["rev"][i9]
    p["cl"] = p["ap"] + p["wp"] + p["op"]
    activos_op = p["cash_obj"] + p["ar"] + p["inv"] + p["oca"] + p["nfa"]
    p["afn0"] = activos_op - p["cl"] - (r["eq"][i9] + ebit * (1 - t))
    d = p["afn0"] / (1 - tasa_int * (1 - t)) if p["afn0"] > 0 else p["afn0"]
    p["debt"] = max(d, 0)
    p["excedente"] = max(-d, 0)
    p["int"] = p["debt"] * tasa_int
    p["tax"] = (ebit - p["int"]) * t
    p["ni"] = ebit - p["int"] - p["tax"]
    p["eq"] = r["eq"][i9] + p["ni"]
    p["cash"] = p["cash_obj"] + p["excedente"]
    p["ca"] = p["cash"] + p["ar"] + p["inv"] + p["oca"]
    p["ta"] = p["ca"] + p["nfa"]
    p["check"] = p["ta"] - (p["cl"] + p["debt"] + p["eq"])
    p["purch"] = purch
    # indicadores 2020
    p["crec"] = g
    p["mb"] = p["gp"] / rev
    p["mo"] = ebit / rev
    p["mn"] = p["ni"] / rev
    p["roa"] = p["ni"] / p["ta"]
    p["roe"] = p["ni"] / p["eq"]
    p["dso"] = p["ar"] / rev * 365
    p["dio"] = p["inv"] / cogs * 365
    p["dpo"] = p["ap"] / purch * 365
    p["ccc"] = p["dso"] + p["dio"] - p["dpo"]
    p["rotaf"] = rev / p["nfa"]
    p["cash_pct"] = p["cash"] / rev
    p["liq"] = p["ca"] / (p["cl"] + p["debt"])
    p["de"] = p["debt"] / p["eq"]
    p["cobertura"] = ebit / p["int"] if p["int"] else None
    p["deuda_ebitda"] = p["debt"] / (ebit + dep)
    dktno = (p["ar"] - r["ar"][i9]) + (p["inv"] - r["inv"][i9]) + (p["oca"] - r["oca"][i9]) - (p["cl"] - r["cl"][i9])
    p["dktno"] = dktno
    p["cfo"] = p["ni"] + dep - dktno
    p["fcf"] = p["cfo"] - capex
    return p


CAJA_2019 = H["cash"][3] / H["rev"][3]  # 0,9% de las ventas

ESCENARIOS = [
    ("Base: ratios 2019", {}),
    ("Base sin reponer caja (0,9%)", {"caja_obj": CAJA_2019}),
    ("Medida: cobranza DSO 45", {"dso": 45}),
    ("Medida: inventario DIO 450", {"dio": 450}),
    ("Base con salario dueños fijo", {"sga_fijo": 50}),
    ("Plan combinado (DSO 45 + DIO 450)", {"dso": 45, "dio": 450}),
    ("Plan recomendado: combinado + caja 3% con línea comprometida", {"dso": 45, "dio": 450, "caja_obj": 0.03}),
    ("Riesgo: alza salario H2A (+2 pp costo)", {"d_cogs": 0.02}),
    ("Riesgo: ventas +15%, inventario para +30%", {"g": 0.15, "g_inv": 0.30}),
]
CRECIMIENTOS = [0.0, 0.05, 0.10, 0.15, 0.20, 0.25, 0.30, 0.35]


def crecimiento_sostenible(r, k):
    """Crecimiento que se financia solo con utilidades (sin dividendos) en 2019."""
    a_s = r["ta"][3] / r["rev"][3]
    l_s = r["cl"][3] / r["rev"][3]
    m = k["mn"][3]
    return m / (a_s - l_s - m), r["ni"][3] / r["eq"][2]


if __name__ == "__main__":
    r = historico()
    k = indicadores(r)
    # 1) Verificar Exhibit 1: patrimonio = patrimonio anterior + utilidad (sin dividendos)
    for i in range(1, 4):
        print(ANIOS[i], "eq check", round(r["eq"][i] - (r["eq"][i - 1] + r["ni"][i]), 3),
              "| NFA check", round(r["nfa"][i] - (r["nfa"][i - 1] + r["capex"][i] - r["dep"][i]), 3))
    # 2) Verificar Exhibit 2
    for key, pub in EXH2.items():
        calc = k[key]
        diffs = [None if c is None else round(c - p, 4) for c, p in zip(calc, pub)]
        print(f"{key:6s}", [None if c is None else round(c, 4) for c in calc], "dif vs publicado:", diffs)
    for kk in ["tasa", "sga_pct", "ccc", "rotat", "cash_pct", "liq", "acida", "ktno", "ktno_pct"]:
        print(kk, [round(x, 4) for x in k[kk]])
    print("ni", r["ni"], "eq", [round(x, 1) for x in r["eq"]])
    for f in flujos(r):
        print({a: round(b, 1) for a, b in f.items()})
    for nombre, kw in ESCENARIOS:
        p = proyeccion(r, **kw)
        print(f"{nombre:45s} AFN0={p['afn0']:7.1f} deuda={p['debt']:6.1f} exc={p['excedente']:6.1f} "
              f"NI={p['ni']:5.1f} mn={p['mn']:.3%} roe={p['roe']:.3%} cfo={p['cfo']:7.1f} fcf={p['fcf']:7.1f} "
              f"D/E={p['de']:.2f} cob={p['cobertura'] or 0:.1f} D/EBITDA={p['deuda_ebitda']:.2f} liq={p['liq']:.2f} "
              f"ccc={p['ccc']:.0f} check={p['check']:.6f}")
    for g in CRECIMIENTOS:
        p = proyeccion(r, g=g)
        p2 = proyeccion(r, g=g, caja_obj=CAJA_2019)
        print(f"g={g:.0%} deuda={p['debt']:6.1f} exc={p['excedente']:6.1f} | sin caja obj deuda={p2['debt']:6.1f} exc={p2['excedente']:6.1f}")
    print("SGR", crecimiento_sostenible(r, k))
    base = proyeccion(r)
    print({a: (round(b, 3) if isinstance(b, float) else b) for a, b in base.items()})
