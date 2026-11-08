# Radio One, Inc. (HBS 204-S05) – Valoración de las emisoras de Clear Channel

| Archivo | Contenido |
|---|---|
| `original/Radio_One_Solucion_Financiera.xlsx` | Solución original del grupo, sin cambios |
| `Radio_One_Solucion_Financiera_corregida.xlsx` | Versión corregida. **Naranja = celda modificada, verde = celda nueva.** La hoja *Registro de cambios* lista cada celda tocada (antes, después, motivo) |
| `corregir.py` | Script que aplica las correcciones al original (`python corregir.py original.xlsx salida.xlsx`) |

## Correcciones principales
1. **Año 2000 (t0):** sólo se incluye la inversión inicial en capital de trabajo; el flujo operativo de 2000 es del vendedor.
2. **Escudo fiscal de la amortización:** el flujo operativo paga impuestos completos sobre el EBITDA y el ahorro fiscal de US$90 M/año se valora aparte sólo para 2001-2015 (nota b, Anexo 9), en vez de perpetuarse en el valor terminal.
3. **Beta:** se reapalanca con la fórmula del Anexo 8 (beta de deuda 0,25) en lugar de Hamada.
4. Menores: mediana EBITDA 17,8x (no 18,2x), promedio del sector sin Radio One ni Hispanic, rango de oferta redefinido.

## Resultado (US$ miles)
| Concepto | Original | Corregido |
|---|---|---|
| DCF perpetuidad | 1.045.364 (16,8x FCE 2000) | 874.701 (14,1x) |
| DCF múltiplo terminal | 1.251.249 | 1.338.620 (21,5x) |
| Rango de oferta | 1.045.364 – 1.251.249 | 874.701 – 1.337.150 |
| Oferta de apertura sugerida | – | 1.243.860 (20x) |
