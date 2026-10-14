# Reporte de análisis: ingreso y bienestar subjetivo (World Happiness Report)

**Pregunta:** ¿La evolución del ingreso de los países a través del tiempo está acompañada por cambios en sus niveles de bienestar?

> Todas las cifras son descriptivas. Una asociación o correlación no implica que el ingreso *cause* cambios en el bienestar (ni al revés); hay variables omitidas, causalidad inversa posible y cambios en la composición de la muestra.

## 1. Inspección de la base

- Observaciones (filas país-año): **1949**
- Países: **166**
- Rango de años: **2005–2020**
- Duplicados país-año: **0**
- Variables (11): `Country name`, `year`, `Life Ladder`, `Log GDP per capita`, `Social support`, `Healthy life expectancy at birth`, `Freedom to make life choices`, `Generosity`, `Perceptions of corruption`, `Positive affect`, `Negative affect`

### Tipos de datos y valores faltantes

| variable | faltantes | porcentaje | tipo |
|---|---|---|---|
| Country name | 0 | 0.0 | str |
| year | 0 | 0.0 | int64 |
| Life Ladder | 0 | 0.0 | float64 |
| Log GDP per capita | 36 | 1.85 | float64 |
| Social support | 13 | 0.67 | float64 |
| Healthy life expectancy at birth | 55 | 2.82 | float64 |
| Freedom to make life choices | 32 | 1.64 | float64 |
| Generosity | 89 | 4.57 | float64 |
| Perceptions of corruption | 110 | 5.64 | float64 |
| Positive affect | 22 | 1.13 | float64 |
| Negative affect | 16 | 0.82 | float64 |

### Observaciones por año (panel no balanceado)

| year | observaciones |
|---|---|
| 2005 | 27 |
| 2006 | 89 |
| 2007 | 102 |
| 2008 | 110 |
| 2009 | 114 |
| 2010 | 124 |
| 2011 | 146 |
| 2012 | 142 |
| 2013 | 137 |
| 2014 | 145 |
| 2015 | 143 |
| 2016 | 142 |
| 2017 | 147 |
| 2018 | 142 |
| 2019 | 144 |
| 2020 | 95 |

### Faltantes por año en las variables del gráfico

| year | Log GDP per capita | Life Ladder | Healthy life expectancy at birth |
|---|---|---|---|
| 2006 | 1 | 0 | 1 |
| 2007 | 0 | 0 | 1 |
| 2008 | 0 | 0 | 2 |
| 2009 | 1 | 0 | 3 |
| 2010 | 1 | 0 | 3 |
| 2011 | 1 | 0 | 5 |
| 2012 | 2 | 0 | 6 |
| 2013 | 1 | 0 | 4 |
| 2014 | 3 | 0 | 5 |
| 2015 | 3 | 0 | 4 |
| 2016 | 3 | 0 | 5 |
| 2017 | 1 | 0 | 4 |
| 2018 | 6 | 0 | 4 |
| 2019 | 6 | 0 | 5 |
| 2020 | 7 | 0 | 3 |

## 2. Decisiones de limpieza

1. **Codificación:** el CSV empieza con BOM UTF-8; se lee con `encoding="utf-8-sig"` para que la primera columna se llame `Country name`.
2. **Nombres de países:** se eliminan espacios sobrantes. No se renombran ni fusionan países.
3. **Duplicados país-año:** se verificó que no existen (0).
4. **Valores faltantes — análisis de casos completos por país-año:** un país se dibuja en un año solo si tiene `Log GDP per capita`, `Life Ladder` y `Healthy life expectancy at birth`. **No se imputa** (ni interpolación, ni media, ni arrastre del último valor), porque cualquier imputación inventaría posiciones en el gráfico.
   - Filas excluidas: **73** de 1949 (3.7 %). Base final: **1876** filas, **159** países.
   - Países que nunca tienen las tres variables y por eso no aparecen: Cuba, Hong Kong S.A.R. of China, Kosovo, North Cyprus, Somalia, Somaliland region, South Sudan.
   - Países con más años excluidos: Kosovo (14), Hong Kong S.A.R. of China (11), Taiwan Province of China (10), Palestinian Territories (9), North Cyprus (7), Somaliland region (4), South Sudan (4), Iran (3), Somalia (3), Venezuela (3).
5. **Panel no balanceado:** el número de países cambia por año (muy bajo en 2005 y menor en 2020). El video muestra *n* en cada frame para que no se confunda un cambio de composición con un cambio de los países.
6. Las variables restantes de la base no se usan; no se crean variables nuevas en el gráfico.

## 3. Resumen por año (detalle completo en `output/resumen_datos.csv`)

| year | paises_en_base | paises_graficados | log_gdp_mediana | life_ladder_mediana | healthy_life_exp_mediana | corr_pearson_gdp_ladder |
|---|---|---|---|---|---|---|
| 2005 | 27 | 27 | 10.46 | 6.62 | 67.0 | 0.68 |
| 2006 | 89 | 87 | 9.04 | 5.08 | 62.68 | 0.84 |
| 2007 | 102 | 101 | 9.27 | 5.22 | 63.78 | 0.77 |
| 2008 | 110 | 108 | 9.21 | 5.18 | 63.4 | 0.85 |
| 2009 | 114 | 111 | 9.33 | 5.35 | 64.35 | 0.73 |
| 2010 | 124 | 121 | 9.46 | 5.46 | 65.1 | 0.77 |
| 2011 | 146 | 141 | 9.33 | 5.22 | 64.66 | 0.77 |
| 2012 | 142 | 136 | 9.46 | 5.26 | 64.91 | 0.78 |
| 2013 | 137 | 133 | 9.44 | 5.24 | 65.12 | 0.78 |
| 2014 | 145 | 138 | 9.43 | 5.32 | 65.3 | 0.81 |
| 2015 | 143 | 137 | 9.46 | 5.34 | 65.6 | 0.82 |
| 2016 | 142 | 135 | 9.51 | 5.42 | 65.8 | 0.83 |
| 2017 | 147 | 142 | 9.5 | 5.57 | 65.9 | 0.8 |
| 2018 | 142 | 135 | 9.56 | 5.47 | 66.3 | 0.78 |
| 2019 | 144 | 136 | 9.59 | 5.64 | 66.65 | 0.76 |
| 2020 | 95 | 88 | 9.88 | 5.85 | 67.85 | 0.84 |

## 4. Hallazgos descriptivos

- **Entre países (transversal):** en todos los años los países con mayor Log GDP per capita tienden a reportar mayor Life Ladder. La correlación de Pearson anual va de 0.68 a 0.85; agrupando todos los país-año es 0.79.
- **Dentro de cada país (en el tiempo):** al quitar a cada país su promedio, la correlación entre las desviaciones de ingreso y de bienestar es **0.29**, mucho más débil que la transversal.
- **Cambios de largo plazo** (países con ≥ 10 años entre su primera y última observación, n = 135): la correlación entre el cambio de Log GDP y el cambio de Life Ladder es 0.34.

| Cambio país (≥ 10 años) | Life Ladder bajó/igual | Life Ladder subió |
|---|---|---|
| Log GDP bajó/igual | 12 | 6 |
| Log GDP subió | 37 | 80 |

**Respuesta a la pregunta (descriptiva):** el ingreso y el bienestar están fuertemente asociados *entre* países en cada año, pero la evolución del ingreso *dentro* de un país a lo largo de 2005–2020 se acompaña de cambios en Life Ladder de forma mucho menos consistente (asociación positiva débil a moderada). De los 117 países cuyo Log GDP subió, 80 también aumentaron su Life Ladder y 37 no; India y Afghanistan, por ejemplo, aumentaron su ingreso mientras su Life Ladder caía, mientras que China o Germany registraron aumentos en ambas variables. El periodo es corto, el Log GDP cambia poco año a año y otros factores no considerados pueden pesar. Estos resultados no permiten afirmar efectos causales.

Figuras de apoyo: `output/figures/correlacion_por_anio.png`, `output/figures/cambios_por_pais.png`, `output/figures/frame_*.png`.

## 5. Video

- Archivo: `output/video_final.mp4`
- Cada año es una imagen que se mantiene 2 s; el MP4 se escribe a 30 fps para compatibilidad con reproductores (960 fotogramas en total), duración 32.0 s
- Resolución: 1920x1080; códec h264

## 6. Limitaciones

- Panel no balanceado; los promedios y correlaciones anuales mezclan cambios reales con cambios en qué países están en la muestra (sobre todo en 2005 y 2020).
- Log GDP per capita está en escala logarítmica: distancias horizontales iguales representan cambios porcentuales iguales, no montos iguales.
- Life Ladder es una medida autorreportada basada en encuestas con error muestral; diferencias pequeñas entre años pueden no ser significativas.
- El tamaño de los círculos usa una escala lineal fija entre el mínimo y el máximo global de esperanza de vida saludable; no es proporcional al valor absoluto (exagera visualmente diferencias moderadas). La leyenda muestra valores de referencia.
- 2020 coincide con la pandemia de COVID-19 y tiene menos países encuestados.
- Ningún resultado es causal.
