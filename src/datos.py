"""Descarga, inspección y limpieza de la base World Happiness Report."""
import urllib.request

import pandas as pd

import config as C


def descargar(forzar: bool = False) -> None:
    C.DIR_DATOS.mkdir(parents=True, exist_ok=True)
    if C.ARCHIVO_CRUDO.exists() and not forzar:
        print(f"[datos] Usando copia local: {C.ARCHIVO_CRUDO.relative_to(C.RAIZ)}")
        return
    print(f"[datos] Descargando {C.URL_DATOS}")
    urllib.request.urlretrieve(C.URL_DATOS, C.ARCHIVO_CRUDO)


def cargar() -> pd.DataFrame:
    # El archivo trae BOM UTF-8 al inicio; utf-8-sig evita que la primera columna
    # se llame "﻿Country name".
    return pd.read_csv(C.ARCHIVO_CRUDO, encoding="utf-8-sig")


def inspeccionar(df: pd.DataFrame) -> dict:
    faltantes = pd.DataFrame({
        "faltantes": df.isna().sum(),
        "porcentaje": (df.isna().mean() * 100).round(2),
        "tipo": df.dtypes.astype(str),
    })
    return {
        "n_obs": len(df),
        "n_paises": df[C.PAIS].nunique(),
        "anio_min": int(df[C.ANIO].min()),
        "anio_max": int(df[C.ANIO].max()),
        "variables": list(df.columns),
        "faltantes": faltantes,
        "duplicados": int(df.duplicated([C.PAIS, C.ANIO]).sum()),
        "obs_por_anio": df.groupby(C.ANIO).size(),
        "faltantes_por_anio": df.groupby(C.ANIO)[C.VARS_GRAFICO].apply(lambda g: g.isna().sum()),
    }


def limpiar(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Análisis de casos completos en las tres variables del gráfico.

    No se imputa: una fila (país-año) sin X, Y o tamaño no se dibuja ese año.
    Devuelve la base limpia y las filas excluidas (para documentarlas).
    """
    df = df.copy()
    df[C.PAIS] = df[C.PAIS].str.strip()
    df = df.drop_duplicates([C.PAIS, C.ANIO])
    completo = df[C.VARS_GRAFICO].notna().all(axis=1)
    limpio = df.loc[completo].sort_values([C.PAIS, C.ANIO]).reset_index(drop=True)
    excluidas = df.loc[~completo, [C.PAIS, C.ANIO] + C.VARS_GRAFICO]
    limpio.to_csv(C.ARCHIVO_LIMPIO, index=False)
    return limpio, excluidas


def resumen_por_anio(crudo: pd.DataFrame, limpio: pd.DataFrame) -> pd.DataFrame:
    g = limpio.groupby(C.ANIO)
    res = pd.DataFrame({
        "paises_en_base": crudo.groupby(C.ANIO)[C.PAIS].nunique(),
        "paises_graficados": g[C.PAIS].nunique(),
    })
    res["paises_excluidos_por_faltantes"] = res["paises_en_base"] - res["paises_graficados"]
    for var, corto in [(C.X, "log_gdp"), (C.Y, "life_ladder"), (C.TAM, "healthy_life_exp")]:
        res[f"{corto}_media"] = g[var].mean()
        res[f"{corto}_mediana"] = g[var].median()
        res[f"{corto}_min"] = g[var].min()
        res[f"{corto}_max"] = g[var].max()
    res["corr_pearson_gdp_ladder"] = g.apply(lambda d: d[C.X].corr(d[C.Y]))
    res["corr_spearman_gdp_ladder"] = g.apply(lambda d: d[C.X].rank().corr(d[C.Y].rank()))  # Spearman sin scipy
    return res.round(4).reset_index()
