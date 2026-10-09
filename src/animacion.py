"""Animación tipo Hans Rosling: ingreso (X) vs. bienestar (Y) por año."""
import matplotlib

matplotlib.use("Agg")
import matplotlib.patheffects as pe
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.animation import FFMpegWriter, FuncAnimation
from matplotlib.lines import Line2D

import config as C

GRIS = "#9AA0A6"
TEXTO = "#202124"
TEXTO_SUAVE = "#5F6368"

# Escala de tamaño FIJA para todos los años: el área del círculo se mapea linealmente
# desde el rango global de esperanza de vida saludable a [AREA_MIN, AREA_MAX] pt².
AREA_MIN, AREA_MAX = 25, 900


def _limites(df: pd.DataFrame) -> dict:
    """Límites constantes de ejes y de tamaño, calculados sobre todos los años."""
    return {
        "x": (np.floor(df[C.X].min() * 2) / 2 - 0.25, np.ceil(df[C.X].max() * 2) / 2 + 0.25),
        "y": (np.floor(df[C.Y].min() * 2) / 2 - 0.25, np.ceil(df[C.Y].max() * 2) / 2 + 0.25),
        "tam": (df[C.TAM].min(), df[C.TAM].max()),
    }


def _area(v, rango):
    lo, hi = rango
    return AREA_MIN + (np.asarray(v) - lo) / (hi - lo) * (AREA_MAX - AREA_MIN)


class Animador:
    def __init__(self, df: pd.DataFrame):
        self.df = df
        self.anios = sorted(df[C.ANIO].unique())
        self.lim = _limites(df)
        self.fig, self.ax = plt.subplots(figsize=(C.ANCHO_IN, C.ALTO_IN), dpi=C.DPI)
        self._fondo()
        self.dinamicos = []

    # ---------- elementos estáticos ----------
    def _fondo(self):
        fig, ax = self.fig, self.ax
        fig.patch.set_facecolor("white")
        fig.subplots_adjust(left=0.065, right=0.82, top=0.86, bottom=0.17)
        ax.set_xlim(*self.lim["x"])
        ax.set_ylim(*self.lim["y"])
        ax.set_xlabel("Log GDP per capita  (logaritmo del PIB per cápita, PPA)", fontsize=15, color=TEXTO, labelpad=10)
        ax.set_ylabel("Life Ladder  (bienestar subjetivo, escala 0–10)", fontsize=15, color=TEXTO, labelpad=10)
        ax.tick_params(labelsize=13, colors=TEXTO_SUAVE)
        ax.grid(color="#E8EAED", linewidth=1)
        ax.set_axisbelow(True)
        for s in ("top", "right"):
            ax.spines[s].set_visible(False)
        for s in ("left", "bottom"):
            ax.spines[s].set_color("#BDC1C6")

        fig.text(0.065, 0.945, "Ingreso y bienestar subjetivo de los países, 2005–2020",
                 fontsize=24, weight="bold", color=TEXTO)
        fig.text(0.065, 0.905,
                 "Cada círculo es un país en un año. Tamaño = Healthy life expectancy at birth. "
                 "Las estelas muestran la trayectoria observada de los países resaltados.",
                 fontsize=13.5, color=TEXTO_SUAVE)
        fig.text(0.065, 0.022,
                 "Fuente: World Happiness Report (base longitudinal). Se omite el país-año si falta alguna de las "
                 "tres variables (sin imputación); el número de países varía por año.\n"
                 "La figura muestra asociaciones descriptivas; no permite concluir relaciones causales.",
                 fontsize=11.5, color=TEXTO_SUAVE, linespacing=1.5)

        # Leyenda de tamaño (valores de referencia con la misma escala fija)
        refs = [50, 60, 70]
        handles = [plt.scatter([], [], s=_area(r, self.lim["tam"]), color=GRIS, alpha=0.5,
                               edgecolor="white") for r in refs]
        leg1 = ax.legend(handles, [f"{r} años" for r in refs], title="Esperanza de vida\nsaludable al nacer",
                         loc="upper left", bbox_to_anchor=(1.02, 1.0), frameon=False, fontsize=12.5,
                         title_fontsize=13, labelspacing=1.6, borderpad=0.8, handletextpad=1.5)
        leg1._legend_box.align = "left"
        ax.add_artist(leg1)
        # Leyenda de países resaltados
        hs = [Line2D([], [], marker="o", linestyle="-", color=col, markersize=10, linewidth=2)
              for col in C.PAISES_RESALTADOS.values()]
        hs.append(Line2D([], [], marker="o", linestyle="", color=GRIS, alpha=0.5, markersize=10))
        leg2 = ax.legend(hs, list(C.PAISES_RESALTADOS) + ["Otros países"], title="Países resaltados",
                         loc="upper left", bbox_to_anchor=(1.02, 0.60), frameon=False, fontsize=12.5,
                         title_fontsize=13)
        leg2._legend_box.align = "left"

    # ---------- elementos por año ----------
    def _limpiar(self):
        for a in self.dinamicos:
            a.remove()
        self.dinamicos = []

    def dibujar(self, anio: int):
        self._limpiar()
        ax, lim, df = self.ax, self.lim, self.df
        actual = df[df[C.ANIO] == anio].sort_values(C.TAM, ascending=False)

        n = actual[C.PAIS].nunique()
        r = actual[C.X].corr(actual[C.Y])
        aviso = (f"\nMuestra reducida (< {C.UMBRAL_MUESTRA} países): comparar con cautela"
                 if n < C.UMBRAL_MUESTRA else "")
        self.dinamicos.append(ax.text(
            0.02, 0.97, f"{anio}", transform=ax.transAxes, va="top", fontsize=34, weight="bold", color=TEXTO))
        self.dinamicos.append(ax.text(
            0.02, 0.86, f"Países con datos completos: {n}\nCorrelación de Pearson entre países: {r:.2f}{aviso}",
            transform=ax.transAxes, va="top", fontsize=13.5, color=TEXTO_SUAVE, linespacing=1.6,
            bbox=dict(boxstyle="round,pad=0.5", facecolor="#F8F9FA", edgecolor="#DADCE0")))

        otros = actual[~actual[C.PAIS].isin(C.PAISES_RESALTADOS)]
        self.dinamicos.append(ax.scatter(otros[C.X], otros[C.Y], s=_area(otros[C.TAM], lim["tam"]),
                                         color=GRIS, alpha=0.40, edgecolor="white", linewidth=0.6, zorder=2))

        contorno = [pe.withStroke(linewidth=3.5, foreground="white")]
        for pais, color in C.PAISES_RESALTADOS.items():
            hist = df[(df[C.PAIS] == pais) & (df[C.ANIO] <= anio)].sort_values(C.ANIO)
            if hist.empty:
                continue
            # Estela: solo puntos observados (años sin dato no se interpolan visualmente como datos)
            ln, = ax.plot(hist[C.X], hist[C.Y], color=color, linewidth=1.8, alpha=0.55, zorder=3)
            pts = ax.scatter(hist[C.X].iloc[:-1], hist[C.Y].iloc[:-1], s=14, color=color, alpha=0.55, zorder=3)
            self.dinamicos += [ln, pts]
            fila = hist[hist[C.ANIO] == anio]
            if fila.empty:   # país sin dato completo este año: estela + aviso, sin círculo
                ult = hist.iloc[-1]
                self.dinamicos.append(ax.annotate(
                    f"{pais} (sin dato en {anio})", (ult[C.X], ult[C.Y]), xytext=(10, -14),
                    textcoords="offset points", fontsize=11.5, style="italic", color=color, alpha=0.8,
                    zorder=5, path_effects=contorno))
                continue
            fx, fy, ft = fila[C.X].iloc[0], fila[C.Y].iloc[0], fila[C.TAM].iloc[0]
            self.dinamicos.append(ax.scatter([fx], [fy], s=_area(ft, lim["tam"]), color=color, alpha=0.85,
                                             edgecolor="#202124", linewidth=1.0, zorder=4))
            offset, ha = C.POS_ETIQUETA.get(pais, C.POS_ETIQUETA_DEFECTO)
            self.dinamicos.append(ax.annotate(pais, (fx, fy), xytext=offset, textcoords="offset points",
                                              ha=ha, fontsize=13, weight="bold", color=color, zorder=5,
                                              path_effects=contorno))
        return self.dinamicos

    def guardar_png(self, anio: int, ruta):
        self.dibujar(anio)
        self.fig.savefig(ruta, dpi=C.DPI)

    def exportar_video(self, ruta):
        anim = FuncAnimation(self.fig, lambda i: self.dibujar(self.anios[i]),
                             frames=len(self.anios), blit=False)
        writer = FFMpegWriter(fps=C.FPS, codec="libx264", bitrate=4000,
                              extra_args=["-pix_fmt", "yuv420p"],
                              metadata={"title": "Ingreso y bienestar subjetivo 2005-2020"})
        anim.save(ruta, writer=writer, dpi=C.DPI)
        plt.close(self.fig)
        return len(self.anios)
