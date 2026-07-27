# Repositorio Inversiones

Contiene **dos cosas independientes**:

## 1 · `panel_control_TOP_P1/` — Control de Cierre QAQC

Entregable semanal de Besalco Montajes: panel HTML + PPT ejecutiva del cierre QAQC de tres
proyectos (MASA, Desaladora, Talabre).

**Es el trabajo activo.** Antes de tocar nada ahí, leer
[`panel_control_TOP_P1/CLAUDE.md`](panel_control_TOP_P1/CLAUDE.md): tiene qué archivos pedir
cada semana, los comandos, las reglas de homologación entre proyectos (validadas por el
usuario, no cambiarlas sin avisar) y las trampas ya descubiertas en los archivos de origen.

Resumen de un vistazo:

```bash
cd panel_control_TOP_P1
python3 actualizar.py  <Estatus_Resumen_General_QAQC.xlsx>              # MASA
python3 desaladora.py  --reporte <REPORTE_GERENCIAL_*> --punch <Listado_Puntos_Punch_*>
python3 talabre.py     --status  <TalabreSTATUS_PEC>   --dt    <TalabreCuadro_DT>
node gen_ppt.js        # SIEMPRE al final: regenera la PPT y la embebe en el panel
```

## 2 · `src/`, `config/`, `reports/` — Monitor de cartera BTG

Motor de reglas en Python para monitorear cartera de acciones (CL/US) y detectar
oportunidades. Ver [`README.md`](README.md). No tiene relación con el panel QAQC.

---

## Entorno

- **Rama de trabajo:** `claude/top-p1-control-panel-gl1gva`.
- **Chromium para Playwright:** `/opt/pw-browsers/chromium-1194/chrome-linux/chrome`
  (el `executable_path` por defecto apunta a una versión que no está instalada).
- **LibreOffice no abre PPTX** en este contenedor: falta el filtro de Impress. Verificar las
  presentaciones leyendo su XML con `python-pptx`, no intentando convertirlas.
