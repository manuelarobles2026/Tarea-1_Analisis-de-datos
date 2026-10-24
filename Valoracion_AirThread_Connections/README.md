# Valoración de AirThread Connections (HBS 4263)

Valoración de AirThread para su posible compra por American Cable Communications: valor intrínseco, costo de capital, sinergias y estrategia de negociación.

| Archivo | Contenido |
|---|---|
| `informe/informe_airthread.pdf` | Informe final (también en `.md` y `.docx`). |
| `modelo/modelo_airthread.xlsx` | Modelo con fórmulas: Supuestos, Historicos, Comparables, Costo_Capital, Deuda, Proyeccion, Sinergias, Valoracion, Escenarios, Sensibilidad, Verificacion y Control. |
| `informe/tabla_control.md` | Tabla de control "cifra del informe → celda del Excel" (la misma lista está en la hoja Control). |
| `verificacion/verificar.py` | Recalcula en Python, de forma independiente, 34 cifras clave y las compara con el Excel (`resultado_verificacion.txt`). |
| `modelo/construir_modelo.py` | Construye el Excel. Los datos del caso se escriben una sola vez como insumos (texto azul). |
| `informe/plantilla_informe.md`, `informe/generar_informe.py` | La plantilla del informe usa marcadores `[[Hoja!Celda|formato]]`; el generador los reemplaza con los valores del Excel. |

Para reconstruir todo (requiere LibreOffice, pandoc y Playwright):

```bash
bash construir_todo.sh /ruta/a/recalc.py
```

Los archivos originales del caso (PDF y Excel) no están en el repositorio porque tienen derechos de autor.

Resultado principal: el patrimonio vale 6.758 millones de US$ sin sinergias y 8.963 con sinergias. Se recomienda abrir en 6.083, buscar 7.084 y no pagar más de 7.409.
