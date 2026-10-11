# Caso Cafés Monte Bianco (HBS 204-S18): plan de ganancias 2020

| Archivo | Contenido |
|---|---|
| `Monte_Bianco_Analisis.xlsx` | Modelo con fórmulas: supuestos, anexos, diagnóstico 2019, mix de producción (66 combinaciones), EERR / balance / flujo 2020 (solo marca propia vs. mix óptimo), palancas de mejora y plan mensual |
| `Monte_Bianco_Informe.docx` | Informe escrito con el análisis y las recomendaciones |
| `scripts/` | Código que genera ambos archivos |

Convención del Excel: azul = dato del caso o supuesto editable, negro = fórmula, verde = vínculo a otra hoja, amarillo = supuesto clave.

## Regenerar

```bash
python scripts/construir_excel.py Monte_Bianco_Analisis.xlsx
# recalcular con LibreOffice (p. ej. recalc.py de la skill xlsx) para que las fórmulas tengan valores
python scripts/extraer_y_graficar.py Monte_Bianco_Analisis.xlsx build/
node scripts/construir_informe.js build/ Monte_Bianco_Informe.docx   # requiere el paquete npm `docx`
```

Los archivos originales del caso (PDF y Excel de HBS) no se incluyen por derechos de autor.
