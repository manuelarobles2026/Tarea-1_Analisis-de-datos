"""Configuración central del proyecto (rutas, variables y parámetros visuales)."""
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
DIR_DATOS = RAIZ / "data"
DIR_SALIDA = RAIZ / "output"
DIR_FIGURAS = DIR_SALIDA / "figures"

URL_DATOS = (
    "https://raw.githubusercontent.com/Andres1984/Data-Analysis-with-R/"
    "refs/heads/master/Bases/world-happiness-report.csv"
)
ARCHIVO_CRUDO = DIR_DATOS / "world-happiness-report.csv"
ARCHIVO_LIMPIO = DIR_DATOS / "world-happiness-clean.csv"

VIDEO = DIR_SALIDA / "video_final.mp4"
RESUMEN = DIR_SALIDA / "resumen_datos.csv"
REPORTE = DIR_SALIDA / "reporte_analisis.md"

# Variables de la visualización (nombres exactos de la base; no se crean variables nuevas)
PAIS = "Country name"
ANIO = "year"
X = "Log GDP per capita"
Y = "Life Ladder"
TAM = "Healthy life expectancy at birth"
VARS_GRAFICO = [X, Y, TAM]

# Países resaltados (con etiqueta y estela). Se eligen por cobertura temporal casi
# completa y por representar trayectorias distintas; el resto se muestra en gris sin etiqueta.
PAISES_RESALTADOS = {
    "Colombia": "#E69F00",
    "United States": "#0072B2",
    "China": "#D55E00",
    "India": "#CC79A7",
    "Denmark": "#009E73",
    "Venezuela": "#8C1C13",
    "Afghanistan": "#6A3D9A",
    "Germany": "#56B4E9",
}

# Posición de la etiqueta respecto al círculo (desplazamiento en puntos, alineación
# horizontal) para evitar choques entre países cercanos (Colombia/China, Denmark/Germany/US).
POS_ETIQUETA = {
    "Colombia": ((-14, 10), "right"),
    "China": ((14, 8), "left"),
    "Germany": ((-16, -18), "right"),
    "Denmark": ((-14, 14), "right"),
    "United States": ((16, -4), "left"),
}
POS_ETIQUETA_DEFECTO = ((12, 8), "left")

# Parámetros del video
SEG_POR_ANIO = 2   # cada año (un fotograma de la animación) se muestra 2 segundos
# El MP4 se escribe a una tasa estándar: ffmpeg repite cada año SEG_POR_ANIO * FPS_SALIDA veces.
# Una tasa de 0,5 fps es válida pero muchos reproductores se traban o no arrancan con ella.
FPS_SALIDA = 30
UMBRAL_MUESTRA = 100  # con menos países que esto se muestra un aviso de muestra reducida
ANCHO_IN, ALTO_IN, DPI = 16, 9, 120   # 1920 x 1080 px
