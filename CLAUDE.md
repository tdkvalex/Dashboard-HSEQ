# Repositorio Inversiones

Contiene **dos cosas independientes**:

## La actualización semanal, en un comando

```bash
python3 actualizar_semana.py --entrada <carpeta con los archivos del corte>
python3 verificar_suite.py
```

`actualizar_semana.py` reconoce cada archivo **por su contenido** —los nombres cambian—,
corre los tres módulos en el orden correcto (los `gen_ppt.js` después de su script de Python
y `armar_suite.js` al final) y **se detiene** si algún archivo viene más viejo que el ya
cargado o si falta la planilla de NC externas. Lo que no llega no se toca: ese módulo conserva
su corte anterior.

`verificar_suite.py` revisa las dos PPT y la suite en tres anchos antes de enviar.

Todo el procedimiento del lunes está en [`CADA_LUNES.md`](CADA_LUNES.md): qué pedir, a quién,
las trampas conocidas y qué mirar antes de enviar. Los comandos sueltos de cada módulo, que
siguen sirviendo para reprocesar uno solo, están más abajo.

## 0 · `suite_qaqc/` — Suite QAQC (consola ejecutiva)

Reúne los dashboards de calidad en un solo archivo con portada que cruza **proyectos ×
módulos**: Protocolos, Cierre QAQC y No Conformidades. Los tres cubren los mismos tres
proyectos: Desaladora (P2416), Talabre (P2407) y Arqueros/MASA (P2342).

Ver [`suite_qaqc/README.md`](suite_qaqc/README.md). Punto de entrada:

```bash
cd suite_qaqc
cp ../panel_control_TOP_P1/index.html modulos/cierre_qaqc.html   # tras actualizar el módulo
node armar_suite.js
```

Cada módulo va aislado en su propio iframe, así que se integran **sin reescribir ninguno**.
Detalle crítico: los módulos se empaquetan con un centinela en lugar de `</script`; escaparlos
como `<\/script>` **rompe Protocolos**, que ya trae esa secuencia propia. Está explicado en el
README.

## 1 · `modulo_nc/` — Control de No Conformidades

Módulo de la suite. **Dos** Excel de origen: el registro principal (`Data_NCR.xlsx`, hoja
«Observaciones») y la planilla de NC que el cliente MASA levanta en Arqueros.
Ver [`modulo_nc/README.md`](modulo_nc/README.md).

```bash
cd modulo_nc
python3 no_conformidades.py --data <Data_NCR.xlsx> --externas <Data_NCR*externas.xlsx>
node gen_ppt.js        # SIEMPRE al final: regenera la PPT y la embebe en el panel
```

El panel abre con **«Levantados en la última semana»** (cuántos, de qué disciplina, internos o
externos, y el detalle uno a uno). Esa ventana se cuenta contra `--hoy`, que por defecto es el
día en que se corre: si se reprocesa un corte pasado hay que pasar `--hoy AAAA-MM-DD`, o la
semana sale vacía.

**Son DOS archivos.** El registro principal trae, para Arqueros, solo lo que emite Besalco;
las NC que **el cliente MASA le levanta** vienen en su propia planilla (`--externas`, hoja
«Disposición NC-Externas»). Sin ella Arqueros se ve con 1 abierta y 98,9% de cierre en vez de
36 y 78%. El script avisa si no se pasa.

**El atraso se mide contra el plazo de respuesta, no contra una fecha comprometida.** Hay
**10 días para responder una NC desde que se emite**; del día 11 en adelante corre atraso
(`PLAZO_RESPUESTA` en `no_conformidades.py`). Se cuenta sobre la fecha de emisión porque
**ninguna de las dos fuentes trae una fecha comprometida de cierre**: las columnas de fecha del
Excel solo se llenan al cerrar. Al corte, 50 de las 53 abiertas están fuera de plazo.
**Las opciones de mejora no entran al módulo** (`TIPOS_EXCLUIDOS`): son propuestas, no hallazgos
que corregir. El panel declara cuántas quedaron fuera.
*No volver a dar el atraso por «no calculable»: eso fue un error de criterio ya corregido.*

## 2 · `panel_control_TOP_P1/` — Control de Cierre QAQC

Entregable semanal de Besalco Montajes: panel HTML + PPT ejecutiva del cierre QAQC de tres
proyectos: **P2416 · Desaladora**, **P2407 · Talabre** y **P2342 · Arqueros** —código, nombre
y cliente exactamente como los escribe el módulo de Protocolos, que es la referencia de
identidad de toda la suite.

**Es el trabajo activo.** Antes de tocar nada ahí, leer
[`panel_control_TOP_P1/CLAUDE.md`](panel_control_TOP_P1/CLAUDE.md): tiene qué archivos pedir
cada semana, los comandos, las reglas de homologación entre proyectos (validadas por el
usuario, no cambiarlas sin avisar) y las trampas ya descubiertas en los archivos de origen.

Resumen de un vistazo:

```bash
cd panel_control_TOP_P1
python3 desaladora.py  --reporte <REPORTE_GERENCIAL_*> --punch <Listado_Puntos_Punch_*>
python3 talabre.py     --status  <STATUS_SUBSISTEMAS_TALABRE>  --dt <Detalle_de_TerminacionesBesalco>
python3 actualizar.py  <Estatus_Resumen_General_QAQC.xlsx>   # Arqueros (cliente MASA)
node gen_ppt.js        # SIEMPRE al final: regenera la PPT y la embebe en el panel
```

## 3 · `src/`, `config/`, `reports/` — Monitor de cartera BTG

Motor de reglas en Python para monitorear cartera de acciones (CL/US) y detectar
oportunidades. Ver [`README.md`](README.md). No tiene relación con el panel QAQC.

---

## Entorno

- **Rama de trabajo:** `claude/top-p1-control-panel-gl1gva`.
- **Chromium para Playwright:** `/opt/pw-browsers/chromium-1194/chrome-linux/chrome`
  (el `executable_path` por defecto apunta a una versión que no está instalada).
- **LibreOffice no abre PPTX** en este contenedor: falta el filtro de Impress. Verificar las
  presentaciones leyendo su XML con `python-pptx`, no intentando convertirlas.
