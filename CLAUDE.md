# Dashboard HSEQ — Besalco Montajes

Repositorio de **uso exclusivo para dashboards y paneles de control HSEQ**
(Seguridad y Salud Ocupacional, Medio Ambiente, Calidad/QAQC). Nada que no sea
un tablero HSEQ o lo que lo alimenta entra aquí.

Hoy contiene la **Suite QAQC**, el entregable semanal de calidad. Vino del
repositorio `Inversiones`, donde convivía con un monitor de cartera de acciones
sin ninguna relación; se trajo **con todo su historial**, así que los commits que
explican por qué existe cada regla siguen disponibles con `git log`.

> **Este repositorio es público.** El material trae nombres de cliente, códigos de
> contrato, detalle de no conformidades y costos en UF. Es una decisión tomada a
> conciencia (10-08-2026), no un descuido.

## La actualización semanal

Sin terminal: doble clic en **«Abrir Centro de Carga»** (`.command` / `.bat`) levanta
[`centro_carga.py`](centro_carga.py), una página local donde se arrastran los archivos
del corte, se ve qué llegó y qué falta, y se descarga la Suite ya armada. **La carga
está bloqueada tras la contraseña de administrador**: de ella se guarda solo el hash
en `.centro_carga.json`, que no se versiona. La página
**no calcula nada**: le pasa los archivos a `actualizar_semana.py` y muestra su salida,
así que no existe una segunda implementación de las reglas que pueda divergir. Nunca se
empaqueta dentro del entregable.

Por terminal es lo mismo:

```bash
python3 actualizar_semana.py --entrada <carpeta con los archivos del corte>
python3 verificar_suite.py
python3 auditoria_datos.py <la misma carpeta>
```

`actualizar_semana.py` reconoce cada archivo **por su contenido** —los nombres cambian—,
corre los tres módulos en el orden correcto (los `gen_ppt.js` después de su script de Python
y `armar_suite.js` al final) y **se detiene** si algún archivo viene más viejo que el ya
cargado o si falta la planilla de NC externas. Lo que no llega no se toca: ese módulo conserva
su corte anterior.

`verificar_suite.py` revisa las dos PPT y la suite en tres anchos antes de enviar.
`auditoria_datos.py` cruza la misma cifra en todas las capas —Excel, JSON, portada y PPT—.

Los ocho documentos que entran, cómo se leen y en qué orden se inyectan están dibujados en
[`CADA_LUNES.md`](CADA_LUNES.md) §3. Un corte completo demora unos **6 segundos**: los Excel
se abren desde [`qaqc_excel.py`](qaqc_excel.py), que descarta los estilos con nombre —el
REPORTE_GERENCIAL arrastra 52.000, 11 de sus 12 MB— y lee los nombres de hoja del ZIP en vez
de abrir el libro entero solo para reconocerlo.

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
cp <dashboard_protocolos_del_corte>.html modulos/protocolos.html   # el único que se copia
node armar_suite.js
```

**Cierre QAQC y No Conformidades NO se copian a `modulos/`.** El generador los toma en vivo
de `panel_control_TOP_P1/index.html` y `modulo_nc/index.html`. Si alguien deja una copia en
`modulos/`, esa copia **gana** y el módulo queda congelado en ese corte para siempre, sin que
nada avise. Protocolos sí se copia: llega de otro equipo y no se genera aquí.

**El dashboard de Protocolos SÍ se versiona** (`suite_qaqc/modulos/protocolos.html`).
Es el único insumo que no se puede regenerar desde aquí —llega armado del equipo de
Protocolos— y por eso es el único que se guarda: si se pierde, no hay corte.

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
externos, y el detalle uno a uno). La ventana es la **semana calendario de lunes a domingo que
ya cerró** al corte —con `--hoy` el lunes 03-08 informa lunes 27-07 → domingo 02-08—, no los
últimos 7 días corridos. `--hoy` por defecto es el día en que se corre: si se reprocesa un
corte pasado hay que pasarlo, o se informa la semana equivocada.

**Son DOS archivos.** El registro principal trae, para Arqueros, solo lo que emite Besalco;
las NC que **el cliente MASA le levanta** vienen en su propia planilla (`--externas`, hoja
«Disposición NC-Externas»). Sin ella Arqueros se ve con 1 abierta y 98,9% de cierre en vez de
36 y 78%. El script avisa si no se pasa. Desde el corte 10-08-2026 esa planilla llega como
`Log_Control_NC_MASA_*.xlsm`, un libro con varias hojas; se sigue leyendo la misma.
*La numeración de MASA tiene huecos (al 10-08 faltan 13 números entre el 1 y el 84): **no son
registros perdidos**. La hoja «Data Externas» del mismo libro lista los mismos 71 con
numeración correlativa, así que los huecos son de la numeración del cliente.*

**En Arqueros, las NC del cliente se cuentan SOLO desde el log de MASA.** El registro
principal también trae alguna marcada «Externa Cliente», pero es una copia que se llena
aparte y llega a contradecir al log: al corte 10-08, `Calidad - 0063` figuraba **Cerrada**
en el registro y **«Listo para revisión»** en el log, siendo la misma NC. Manda el log, que
es el documento que se revisa con el cliente. La constante es
`PROYECTO_CLIENTE_SOLO_DEL_LOG` en `no_conformidades.py`; el script declara cuántas descartó
y el JSON las guarda en `control.clienteSoloDelLog`. *Pedido por el usuario (10-08-2026).*

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

---

## Entorno

- **Chromium para Playwright:** `/opt/pw-browsers/chromium-1194/chrome-linux/chrome`
  (el `executable_path` por defecto apunta a una versión que no está instalada).
- **LibreOffice no abre PPTX** en este contenedor: falta el filtro de Impress. Verificar las
  presentaciones leyendo su XML con `python-pptx`, no intentando convertirlas.
