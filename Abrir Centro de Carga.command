#!/bin/bash
# Doble clic para abrir el Centro de Carga (macOS y Linux).
# La primera vez en macOS puede pedir permiso: clic derecho → Abrir.
cd "$(dirname "$0")" || exit 1
python3 centro_carga.py
