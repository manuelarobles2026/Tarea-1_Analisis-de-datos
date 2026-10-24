#!/usr/bin/env bash
# Reconstruye el modelo, lo recalcula, lo verifica con Python y genera el informe.
# Uso: bash construir_todo.sh /ruta/a/recalc.py
set -euo pipefail
cd "$(dirname "$0")"
RECALC=${1:?"ruta a recalc.py (LibreOffice)"}
python3 modelo/construir_modelo.py modelo/modelo_airthread.xlsx
python3 "$RECALC" modelo/modelo_airthread.xlsx 120
python3 verificacion/verificar.py modelo/modelo_airthread.xlsx > verificacion/resultado_verificacion.txt
python3 "$RECALC" modelo/modelo_airthread.xlsx 120
if [ -f informe/generar_informe.py ]; then
  python3 informe/generar_informe.py control modelo/modelo_airthread.xlsx
  python3 "$RECALC" modelo/modelo_airthread.xlsx 120
  python3 informe/generar_informe.py informe modelo/modelo_airthread.xlsx
fi
