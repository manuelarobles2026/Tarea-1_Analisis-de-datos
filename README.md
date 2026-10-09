# Ingreso y bienestar: animación tipo Hans Rosling (World Happiness Report)

**Pregunta:** ¿La evolución del ingreso de los países a través del tiempo está acompañada por cambios en sus niveles de bienestar?

## Qué se hizo

1. **Carga e inspección** de la base longitudinal del World Happiness Report ([fuente](https://raw.githubusercontent.com/Andres1984/Data-Analysis-with-R/refs/heads/master/Bases/world-happiness-report.csv)):
   1949 observaciones país-año, 166 países, años 2005–2020, 11 variables (1 texto, 1 entero, 9 numéricas). Sin duplicados país-año.
2. **Limpieza** documentada (ver abajo), sin imputación.
3. **Animación MP4** de burbujas:

   | Elemento | Variable |
   |---|---|
   | Eje X | `Log GDP per capita` |
   | Eje Y | `Life Ladder` |
   | Tamaño | `Healthy life expectancy at birth` |
   | Etiqueta | `Country name` (solo 8 países resaltados, para no saturar) |
   | Tiempo | `year` (1 frame = 1 año) |

   - Ejes y escala de tamaño **fijos** para todos los años.
   - El año aparece grande en la esquina superior izquierda junto con el número de países graficados y la correlación de Pearson entre países de ese año.
   - Los 8 países resaltados (Colombia, United States, China, India, Denmark, Venezuela, Afghanistan, Germany) llevan color, etiqueta y una **estela** con su trayectoria observada; los demás aparecen en gris sin etiqueta.
   - Si un país resaltado no tiene dato un año, se muestra su estela con la nota "(sin dato en AAAA)" en lugar de un círculo.
   - En los años con menos de 100 países (2005: 27; 2020: 88) aparece un aviso de muestra reducida.
   - El pie del video recuerda que se muestran asociaciones, no causalidad.
4. **Análisis descriptivo complementario**: correlación por año, correlación dentro de cada país (sin el promedio de cada país) y cambios de largo plazo por país.

## Archivos generados

| Archivo | Contenido |
|---|---|
| `output/video_final.mp4` | Animación final (1920×1080, H.264, 16 frames, 2 s por año, 32 s) |
| `output/resumen_datos.csv` | Por año: países en la base, graficados y excluidos; media, mediana, mín. y máx. de las 3 variables; correlaciones de Pearson y Spearman |
| `output/reporte_analisis.md` | Inspección, faltantes, decisiones de limpieza, hallazgos, verificación del video y limitaciones |
| `output/figures/frame_2005.png`, `frame_2012.png`, `frame_2020.png` | Frames de muestra del video |
| `output/figures/correlacion_por_anio.png` | Correlación transversal por año y número de países |
| `output/figures/cambios_por_pais.png` | Cambio en Log GDP vs. cambio en Life Ladder dentro de cada país |
| `data/world-happiness-report.csv` | Base original descargada |
| `data/world-happiness-clean.csv` | Base con casos completos en las 3 variables del gráfico |

## Estructura

```
data/       datos crudos y limpios
src/        código
  config.py     rutas, variables, países resaltados, parámetros del video
  datos.py      descarga, inspección, limpieza y resumen por año
  animacion.py  construcción y exportación de la animación
  analisis.py   estadísticas, figuras de apoyo y reporte
  main.py       ejecuta todo el pipeline y verifica el MP4
output/     video, resumen, reporte y figuras
```

## Decisiones de limpieza

- **Codificación:** el CSV trae BOM UTF-8; se lee con `encoding="utf-8-sig"`.
- **Nombres de países:** solo se eliminan espacios sobrantes; no se renombran ni fusionan países.
- **Duplicados país-año:** se verificó que no hay.
- **Valores faltantes:** en las variables del gráfico faltan `Log GDP per capita` (36) y `Healthy life expectancy at birth` (55); `Life Ladder` está completa. Se usan **casos completos por país-año**: un país se dibuja un año solo si tiene las tres variables. Se excluyen 73 filas (3,7 %); quedan 1876 filas y 159 países. 7 países nunca tienen las tres (p. ej. Kosovo, Hong Kong, Somalia) y no aparecen. **No se imputa** (ni interpolación ni arrastre de valores), porque eso inventaría posiciones en el gráfico. Las estelas solo unen años observados.
- No se usan las demás variables de la base y no se crean variables nuevas.

## Dependencias

- Python ≥ 3.10
- `pandas`, `numpy`, `matplotlib` (ver `requirements.txt`)
- `ffmpeg` instalado en el sistema (usado por matplotlib para escribir el MP4; `ffprobe` para verificarlo)

## Cómo ejecutar desde cero

```bash
pip install -r requirements.txt
python src/main.py                      # usa data/ si ya existe la base
python src/main.py --forzar-descarga    # vuelve a descargar la base
```

El script descarga la base, la inspecciona, limpia, genera el resumen, las figuras y el video, y verifica el MP4 con `ffprobe` (tamaño, códec, resolución, frames y duración).

## Hallazgos principales (descriptivos)

- **Entre países**, en todos los años, más ingreso se asocia con más Life Ladder (Pearson anual entre 0,68 y 0,85).
- **Dentro de cada país**, la asociación en el tiempo es mucho más débil (correlación sin el promedio de cada país = 0,29; correlación entre cambios de largo plazo = 0,34). De los 117 países con ingreso creciente y ≥ 10 años de datos, 80 también aumentaron su Life Ladder y 37 no (p. ej. India, Afghanistan).
- Ningún resultado permite afirmar causalidad.

## Limitaciones

- **Panel no balanceado:** 2005 tiene solo 27 países y 2020 tiene 88; los cambios en promedios o correlaciones pueden reflejar qué países entran en la muestra, no cambios reales.
- `Log GDP per capita` está en logaritmo: distancias horizontales iguales son cambios porcentuales iguales.
- `Life Ladder` es autorreportada y tiene error muestral; cambios pequeños entre años pueden no ser significativos.
- El tamaño de los círculos usa una escala lineal fija entre el mínimo y el máximo global de esperanza de vida saludable (no proporcional al valor absoluto), lo que acentúa las diferencias; la leyenda muestra valores de referencia.
- Solo 8 países llevan etiqueta; la selección es ilustrativa, no representativa.
- 2020 coincide con la pandemia de COVID-19.
- Correlaciones y trayectorias son descriptivas: no hay controles, y puede haber causalidad inversa o variables omitidas.
