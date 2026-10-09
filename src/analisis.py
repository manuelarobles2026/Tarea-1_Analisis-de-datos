"""Estadísticas descriptivas, figuras de apoyo y reporte en Markdown."""
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

import config as C

SPAN_MIN = 10  # años mínimos entre primera y última observación para comparar cambios


def tabla_md(df: pd.DataFrame, indice: bool = False) -> str:
    d = df.reset_index() if indice else df
    filas = ["| " + " | ".join(map(str, d.columns)) + " |",
             "|" + "|".join(["---"] * d.shape[1]) + "|"]
    filas += ["| " + " | ".join(map(str, r)) + " |" for r in d.itertuples(index=False)]
    return "\n".join(filas)


def cambios_por_pais(df: pd.DataFrame) -> pd.DataFrame:
    """Cambio entre la primera y la última observación completa de cada país."""
    g = df.sort_values(C.ANIO).groupby(C.PAIS)
    ini, fin = g.first(), g.last()
    out = pd.DataFrame({
        "anio_inicial": ini[C.ANIO], "anio_final": fin[C.ANIO],
        "delta_log_gdp": fin[C.X] - ini[C.X], "delta_life_ladder": fin[C.Y] - ini[C.Y],
    })
    out["span"] = out["anio_final"] - out["anio_inicial"]
    return out[out["span"] >= SPAN_MIN]


def correlacion_intra(df: pd.DataFrame) -> float:
    """Correlación 'within': se resta a cada país su media (solo variación temporal)."""
    d = df[df.groupby(C.PAIS)[C.PAIS].transform("size") >= 2]
    dx = d[C.X] - d.groupby(C.PAIS)[C.X].transform("mean")
    dy = d[C.Y] - d.groupby(C.PAIS)[C.Y].transform("mean")
    return dx.corr(dy)


def figuras(resumen: pd.DataFrame, cambios: pd.DataFrame) -> None:
    C.DIR_FIGURAS.mkdir(parents=True, exist_ok=True)
    fig, ax1 = plt.subplots(figsize=(10, 5.5), dpi=120)
    ax1.plot(resumen[C.ANIO], resumen["corr_pearson_gdp_ladder"], "o-", color="#0072B2", label="Pearson")
    ax1.plot(resumen[C.ANIO], resumen["corr_spearman_gdp_ladder"], "s--", color="#009E73", label="Spearman")
    ax1.set_ylim(0, 1)
    ax1.set_ylabel("Correlación entre países (Log GDP vs Life Ladder)")
    ax1.set_xlabel("Año")
    ax2 = ax1.twinx()
    ax2.bar(resumen[C.ANIO], resumen["paises_graficados"], color="#DADCE0", alpha=0.6, zorder=0)
    ax2.set_ylabel("Países con datos completos")
    ax1.set_zorder(ax2.get_zorder() + 1)
    ax1.patch.set_visible(False)
    ax1.legend(loc="lower right")
    ax1.set_title("Correlación transversal por año y cobertura de la muestra")
    fig.tight_layout()
    fig.savefig(C.DIR_FIGURAS / "correlacion_por_anio.png")
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(9, 7), dpi=120)
    ax.axhline(0, color="#9AA0A6", lw=1)
    ax.axvline(0, color="#9AA0A6", lw=1)
    ax.scatter(cambios["delta_log_gdp"], cambios["delta_life_ladder"], color="#0072B2", alpha=0.6,
               edgecolor="white")
    for pais in C.PAISES_RESALTADOS:
        if pais in cambios.index:
            ax.annotate(pais, (cambios.loc[pais, "delta_log_gdp"], cambios.loc[pais, "delta_life_ladder"]),
                        xytext=(5, 5), textcoords="offset points", fontsize=9)
    ax.set_xlabel("Cambio en Log GDP per capita (primer → último año observado)")
    ax.set_ylabel("Cambio en Life Ladder")
    ax.set_title(f"Cambio dentro de cada país (países con ≥ {SPAN_MIN} años entre observaciones, n = {len(cambios)})")
    fig.tight_layout()
    fig.savefig(C.DIR_FIGURAS / "cambios_por_pais.png")
    plt.close(fig)


def reporte(info: dict, crudo: pd.DataFrame, limpio: pd.DataFrame, excluidas: pd.DataFrame,
            resumen: pd.DataFrame, cambios: pd.DataFrame, video: dict) -> str:
    r_pool = limpio[C.X].corr(limpio[C.Y])
    r_intra = correlacion_intra(limpio)
    sube_x = cambios["delta_log_gdp"] > 0
    sube_y = cambios["delta_life_ladder"] > 0
    cuadrantes = pd.crosstab(sube_x.map({True: "Log GDP subió", False: "Log GDP bajó/igual"}),
                             sube_y.map({True: "Life Ladder subió", False: "Life Ladder bajó/igual"}))
    cuadrantes.index.name = "Cambio país (≥ 10 años)"
    cuadrantes.columns.name = None
    r_cambios = cambios["delta_log_gdp"].corr(cambios["delta_life_ladder"])
    n_sube_x = int(sube_x.sum())
    n_ambos = int((sube_x & sube_y).sum())
    falt = info["faltantes"].copy()
    falt.index.name = "variable"
    falt_anio = info["faltantes_por_anio"]
    falt_anio = falt_anio[falt_anio.sum(axis=1) > 0]
    paises_sin = excluidas[C.PAIS].value_counts().head(10)
    nunca = sorted(set(crudo[C.PAIS]) - set(limpio[C.PAIS]))
    res_tab = resumen[[C.ANIO, "paises_en_base", "paises_graficados", "log_gdp_mediana",
                       "life_ladder_mediana", "healthy_life_exp_mediana", "corr_pearson_gdp_ladder"]]

    md = f"""# Reporte de análisis: ingreso y bienestar subjetivo (World Happiness Report)

**Pregunta:** ¿La evolución del ingreso de los países a través del tiempo está acompañada por cambios en sus niveles de bienestar?

> Todas las cifras son descriptivas. Una asociación o correlación no implica que el ingreso *cause* cambios en el bienestar (ni al revés); hay variables omitidas, causalidad inversa posible y cambios en la composición de la muestra.

## 1. Inspección de la base

- Observaciones (filas país-año): **{info['n_obs']}**
- Países: **{info['n_paises']}**
- Rango de años: **{info['anio_min']}–{info['anio_max']}**
- Duplicados país-año: **{info['duplicados']}**
- Variables ({len(info['variables'])}): {', '.join(f'`{v}`' for v in info['variables'])}

### Tipos de datos y valores faltantes

{tabla_md(falt, indice=True)}

### Observaciones por año (panel no balanceado)

{tabla_md(info['obs_por_anio'].rename('observaciones').to_frame(), indice=True)}

### Faltantes por año en las variables del gráfico

{tabla_md(falt_anio, indice=True)}

## 2. Decisiones de limpieza

1. **Codificación:** el CSV empieza con BOM UTF-8; se lee con `encoding="utf-8-sig"` para que la primera columna se llame `Country name`.
2. **Nombres de países:** se eliminan espacios sobrantes. No se renombran ni fusionan países.
3. **Duplicados país-año:** se verificó que no existen ({info['duplicados']}).
4. **Valores faltantes — análisis de casos completos por país-año:** un país se dibuja en un año solo si tiene `Log GDP per capita`, `Life Ladder` y `Healthy life expectancy at birth`. **No se imputa** (ni interpolación, ni media, ni arrastre del último valor), porque cualquier imputación inventaría posiciones en el gráfico.
   - Filas excluidas: **{len(excluidas)}** de {info['n_obs']} ({len(excluidas) / info['n_obs'] * 100:.1f} %). Base final: **{len(limpio)}** filas, **{limpio[C.PAIS].nunique()}** países.
   - Países que nunca tienen las tres variables y por eso no aparecen: {', '.join(nunca) if nunca else 'ninguno'}.
   - Países con más años excluidos: {', '.join(f'{p} ({n})' for p, n in paises_sin.items())}.
5. **Panel no balanceado:** el número de países cambia por año (muy bajo en {info['anio_min']} y menor en {info['anio_max']}). El video muestra *n* en cada frame para que no se confunda un cambio de composición con un cambio de los países.
6. Las variables restantes de la base no se usan; no se crean variables nuevas en el gráfico.

## 3. Resumen por año (detalle completo en `output/resumen_datos.csv`)

{tabla_md(res_tab.round(2))}

## 4. Hallazgos descriptivos

- **Entre países (transversal):** en todos los años los países con mayor Log GDP per capita tienden a reportar mayor Life Ladder. La correlación de Pearson anual va de {resumen['corr_pearson_gdp_ladder'].min():.2f} a {resumen['corr_pearson_gdp_ladder'].max():.2f}; agrupando todos los país-año es {r_pool:.2f}.
- **Dentro de cada país (en el tiempo):** al quitar a cada país su promedio, la correlación entre las desviaciones de ingreso y de bienestar es **{r_intra:.2f}**, mucho más débil que la transversal.
- **Cambios de largo plazo** (países con ≥ {SPAN_MIN} años entre su primera y última observación, n = {len(cambios)}): la correlación entre el cambio de Log GDP y el cambio de Life Ladder es {r_cambios:.2f}.

{tabla_md(cuadrantes, indice=True)}

**Respuesta a la pregunta (descriptiva):** el ingreso y el bienestar están fuertemente asociados *entre* países en cada año, pero la evolución del ingreso *dentro* de un país a lo largo de 2005–2020 se acompaña de cambios en Life Ladder de forma mucho menos consistente (asociación positiva débil a moderada). De los {n_sube_x} países cuyo Log GDP subió, {n_ambos} también aumentaron su Life Ladder y {n_sube_x - n_ambos} no; India y Afghanistan, por ejemplo, aumentaron su ingreso mientras su Life Ladder caía, mientras que China o Germany registraron aumentos en ambas variables. El periodo es corto, el Log GDP cambia poco año a año y otros factores no considerados pueden pesar. Estos resultados no permiten afirmar efectos causales.

Figuras de apoyo: `output/figures/correlacion_por_anio.png`, `output/figures/cambios_por_pais.png`, `output/figures/frame_*.png`.

## 5. Video

- Archivo: `output/video_final.mp4`
- Frames: {video.get('frames', 'n/d')} (1 frame = 1 año), {C.FPS} fps → duración {video.get('duracion', 'n/d')} s
- Resolución: {video.get('resolucion', 'n/d')}; códec {video.get('codec', 'n/d')}

## 6. Limitaciones

- Panel no balanceado; los promedios y correlaciones anuales mezclan cambios reales con cambios en qué países están en la muestra (sobre todo en {info['anio_min']} y {info['anio_max']}).
- Log GDP per capita está en escala logarítmica: distancias horizontales iguales representan cambios porcentuales iguales, no montos iguales.
- Life Ladder es una medida autorreportada basada en encuestas con error muestral; diferencias pequeñas entre años pueden no ser significativas.
- El tamaño de los círculos usa una escala lineal fija entre el mínimo y el máximo global de esperanza de vida saludable; no es proporcional al valor absoluto (exagera visualmente diferencias moderadas). La leyenda muestra valores de referencia.
- 2020 coincide con la pandemia de COVID-19 y tiene menos países encuestados.
- Ningún resultado es causal.
"""
    C.REPORTE.write_text(md, encoding="utf-8")
    return md
