"""Pipeline completo: python src/main.py [--forzar-descarga]"""
import json
import subprocess
import sys

import config as C
import analisis
import datos
from animacion import Animador


def verificar_video(ruta) -> dict:
    if not ruta.exists() or ruta.stat().st_size == 0:
        raise RuntimeError(f"El video no existe o está vacío: {ruta}")
    try:
        sal = subprocess.run(
            ["ffprobe", "-v", "error", "-count_frames", "-select_streams", "v:0", "-show_entries",
             "stream=codec_name,width,height,nb_read_frames:format=duration", "-of", "json", str(ruta)],
            capture_output=True, text=True, check=True)
    except (FileNotFoundError, subprocess.CalledProcessError) as e:
        print(f"[verificación] ffprobe no disponible o falló: {e}")
        return {"bytes": ruta.stat().st_size}
    info = json.loads(sal.stdout)
    st = info["streams"][0]
    return {
        "bytes": ruta.stat().st_size,
        "codec": st["codec_name"],
        "resolucion": f"{st['width']}x{st['height']}",
        "frames": int(st["nb_read_frames"]),
        "duracion": round(float(info["format"]["duration"]), 2),
    }


def main(forzar: bool = False) -> None:
    for d in (C.DIR_DATOS, C.DIR_SALIDA, C.DIR_FIGURAS):
        d.mkdir(parents=True, exist_ok=True)

    datos.descargar(forzar)
    crudo = datos.cargar()
    info = datos.inspeccionar(crudo)
    print(f"[inspección] {info['n_obs']} obs, {info['n_paises']} países, "
          f"{info['anio_min']}-{info['anio_max']}, {len(info['variables'])} variables")
    print(info["faltantes"].to_string())

    limpio, excluidas = datos.limpiar(crudo)
    print(f"[limpieza] filas excluidas por faltantes: {len(excluidas)}; base final: {len(limpio)}")
    resumen = datos.resumen_por_anio(crudo, limpio)
    resumen.to_csv(C.RESUMEN, index=False)

    cambios = analisis.cambios_por_pais(limpio)
    analisis.figuras(resumen, cambios)

    anim = Animador(limpio)
    for anio in (anim.anios[0], 2012, anim.anios[-1]):
        anim.guardar_png(anio, C.DIR_FIGURAS / f"frame_{anio}.png")
    print("[video] exportando...")
    anim.exportar_video(C.VIDEO)
    video = verificar_video(C.VIDEO)
    print(f"[verificación] {video}")

    analisis.reporte(info, crudo, limpio, excluidas, resumen, cambios, video)
    print(f"[listo] {C.VIDEO.relative_to(C.RAIZ)}, {C.RESUMEN.relative_to(C.RAIZ)}, {C.REPORTE.relative_to(C.RAIZ)}")


if __name__ == "__main__":
    main(forzar="--forzar-descarga" in sys.argv)
