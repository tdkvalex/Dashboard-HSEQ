# Control de Cierre QAQC — runbook

Panel HTML + PPT ejecutiva del cierre QAQC de **tres proyectos** de Besalco Montajes.
Entregable semanal. Este archivo es la memoria operativa: qué archivos pedir, cómo
procesarlos y qué **no** volver a romper.

---

## 1 · Archivos que hay que pedir cada semana

| Proyecto | Archivos | Hojas que se leen |
|---|---|---|
| **Desaladora** | `REPORTE_GERENCIAL_AAAAMMDD.xlsx` **+** `Listado_Puntos_Punch_Consolidado_AAAA_MM_DD.xlsx` | `REPORTE GERENCIAL`, `Resumen  general` · `LISTADO PUNCH ITEMS` |
| **Talabre** | `STATUS_SUBSISTEMAS_TALABRE.xlsx` **+** `Detalle_de_TerminacionesBesalco.xlsx` | `STATUS` (y `RESUMEN`/`DT` si vienen, como contraste) · `DT` |
| **Arqueros** (cliente MASA) | `Estatus_Resumen_General_QAQC.xlsx` | `BD Caminatas-CTOP`, `BD Detalles Terminación` |

Talabre cambió de archivos en el corte 03-08-2026 (antes: `TalabreSTATUS_PEC.xlsx` +
`TalabreCuadro_DT.xlsx`). El `Detalle_de_TerminacionesBesalco.xlsx` es el **registro completo**
del proyecto (2.024 DT, no solo los 480 pendientes) y cuadra exacto con la hoja STATUS, cosa
que el archivo anterior no lograba. `talabre.py` resuelve las columnas **por nombre de
encabezado** y detecta solo la fila donde empieza, así que soporta los dos formatos.
**OJO con `Programa_de_Caminatas.xlsm`**: trae una hoja «DT» pero es una foto vieja
(emisión hasta 19-06); no usarla como fuente.

**La identidad del proyecto sale del módulo de Protocolos**, que es la referencia de toda la
suite: mismo orden, mismo código, mismo nombre y mismo cliente, escritos igual.

| Código | Nombre | Cliente |
|---|---|---|
| `P2416` | Desaladora | ANTOFAGASTA MINERALS |
| `P2407` | Talabre | CODELCO |
| `P2342` | Arqueros | MASA |

La pestaña se rotula **`P2416 · Desaladora`** con el cliente debajo, y la versaleta se deja
solo en la pestaña corporativa, que no es un proyecto. En gráficos y encabezados de tabla va
el nombre solo, que es lo que cabe: para eso está el helper `etiq()`.
El id interno y el Excel de Arqueros siguen llamándose MASA; lo que cambia es cómo se muestra.

Desaladora y Talabre **necesitan sus dos archivos**: uno trae subsistemas/caminatas/carpetas
y el otro el detalle ítem a ítem. Con uno solo el proyecto queda incompleto.

Si falta un proyecto entero, no pasa nada: su pestaña queda deshabilitada y la PPT sale sin
sus láminas. El resumen corporativo se recalcula con los que haya.

---

## 2 · Comandos

```bash
python3 desaladora.py --reporte /ruta/REPORTE_GERENCIAL_*.xlsx \
                      --punch   /ruta/Listado_Puntos_Punch_Consolidado_*.xlsx
python3 talabre.py    --status  /ruta/STATUS_SUBSISTEMAS_TALABRE.xlsx \
                      --dt      /ruta/Detalle_de_TerminacionesBesalco.xlsx
python3 actualizar.py /ruta/Estatus_Resumen_General_QAQC.xlsx   # Arqueros
node gen_ppt.js        # SIEMPRE al final
```

`desaladora.py` y `talabre.py` aceptan `--hoy AAAA-MM-DD` para fijar la fecha con la que se
calculan los atrasos. **Usar la fecha del corte**, no la del día en que se corre, si se está
reprocesando una semana pasada.

`gen_ppt.js` va al final porque: regenera la PPT con los tres JSON, la **embebe** en
`index.html` (botón «Descargar Informe») y embebe los logos. Correrlo antes de actualizar un
proyecto deja el informe descargable desfasado respecto de las pestañas.

Cada script imprime un resumen y una sección **⚠ AVISOS**. Leerlos siempre: ahí aparecen los
valores nuevos que no reconoce y los cruces que dejaron de cuadrar.

---

## 3 · Reglas de homologación — NO cambiar sin avisar al usuario

Costaron trabajo establecerlas y el usuario las validó. Están explicadas en el panel
(desplegable «Cómo se homologaron los tres proyectos») y en el README.

- **Caminatas: se identifican por NÚMERO, nunca por el porcentaje que representan.**
  El «80%» de Arqueros no es el «80%» de otro proyecto. Arqueros: «80%» = Caminata 1, «100%» = Caminata 2.
- **Caminata vigente** = la que cada proyecto declara como su indicador:
  Arqueros → Caminata 2 · Talabre → Caminata 2 · **Desaladora → depende del tipo de subsistema**
  (Operables por la 2, Facility por la 1, según declara su hoja «Resumen general»).
- **Desaladora: los «Componente» NO son subsistemas del universo.** Son entregas parciales de
  un subsistema que ya está contado (`0587-ESL-201 Comp 1` cuelga de `0587-ESL-201`), así que
  sumarlos duplica alcance. El universo son Operables + Facility = **95**, igual que declara
  la hoja «Resumen general» del propio archivo. Los componentes se informan aparte, con su
  estado, en el bloque `componentes` del JSON y en «Control de calidad del dato».
  *Lo pidió Juan Vergara, jefe de calidad del proyecto (30-07-2026), y el usuario lo aprobó.*
  La constante es `FUERA_DEL_UNIVERSO` en `desaladora.py`.
- **Detalles de terminación = punch list = DT.** Mismo concepto, se suman.
- **El % de cierre de detalles se mide sobre P1+P2, no sobre todas las prioridades.**
  Los P3 y P4 no se abordan mientras el cliente no formalice su ejecución, así que
  incluirlos mide algo que nadie gestiona. El detalle por prioridad sigue completo en la
  pestaña de cada proyecto. *Pedido por Mauricio Rocha (31-07-2026), aprobado por el usuario.*
  La constante es `PRIO_GESTION` en `index.html` y se replica en `gen_ppt.js`.
- **Caminatas: el semáforo mide sobre lo EXIGIBLE, no sobre el universo.** Una caminata con
  fecha agendada todavía por llegar **no es un incumplimiento al corte**: se descuenta del
  denominador (`exigible = total − agendadas en plazo`) y se informa aparte como
  **«programadas en plazo»** — descontarlas no puede significar esconderlas. Se guardan las
  dos cifras (`pct` sobre el universo y `pctExigible`) porque miden cosas distintas y el panel
  muestra ambas. *Pedido por Mauricio Rocha (31-07-2026), aprobado por el usuario.*
  La agendada cuya **fecha ya pasó** sin marcarse realizada queda **«Vencida»**: sigue siendo
  exigible (se agendó y no se hizo) y el panel la declara aparte cuando existe.
- **Talabre: los pozos se agrupan en POZOS, pero LÍNEAS IMPULSIÓN y PQS son áreas propias.**
  Las fuentes nuevas abren los pozos uno a uno (PBO-15…, PBBR-01…, PBN-01, TB-01/02): se
  reagrupan bajo POZOS, que es como siempre se reportó. La versión anterior, que derivaba el
  área desde el *sistema*, metía también LÍNEAS IMPULSIÓN (151 DT) y PQS (145 DT) dentro de
  POZOS; desde el corte 03-08-2026 cada una reporta lo suyo, igual que en la hoja STATUS.
  «ADUCCIÓN» se rotula «IMPULSIÓN ADUCCIÓN». La función es `area_canon()` en `talabre.py`.
- **«Sin base al corte»**: un área sin ningún detalle levantado **y** sin ninguna caminata
  vigente realizada no se puntúa. No hay nada medido, y pintarla del mismo rojo que un frente
  con atrasos reales desinforma. Al corte aplica a Cerro Verde/PS2 en Talabre.
- **Prioridades al dígito**: `P1A→P1`, `P2B→P2`, `P3C→P3`. `P0` (Desaladora) y `P4` (Talabre)
  se mantienen aparte.
- **«Vencido» (Arqueros) = «Atrasado» (Desaladora, Talabre)**: ítem **abierto** cuya fecha
  comprometida ya venció. Los abiertos **sin** fecha comprometida **no** se cuentan como
  atrasados; se informan aparte.
- **Disciplinas**: `CANERIAS→Piping` · `ELECTRICOS→Eléctrica` · `CIVIL` y
  `MOVIMIENTO DE TIERRAS→Obras Civiles` · `INSTRUMENTACION→Instrumentación y Control`.
- **Las carpetas NO se consolidan en una cifra única.** Desaladora y Arqueros miden *estado de
  aprobación*; Talabre mide *% de avance (PEC)*. Promediarlas daría un número sin significado.
  El resumen corporativo muestra la métrica nativa de cada uno y suma solo los dos comparables.

---

## 4 · Trampas ya descubiertas

### Columnas de origen (índices 0-based, ya resueltos en los scripts)
- **Punch de Desaladora** (`LISTADO PUNCH ITEMS`, encabezado fila 8, datos fila 9+):
  `1` Nº caminata · `5` Sistema/Facility · `6` Subsistema · `7` Disciplina · `9` Categoría ·
  `15` Fecha Requerida Cierre · `22` STATUS. *Hubo un desfase de una columna en la primera
  versión: la categoría tomaba los nombres de «Observación por».*
- **REPORTE GERENCIAL** (encabezado fila 10, datos fila 11+): `2` subsistema · `3` tipo ·
  `4` zona · `6/8/10` caminatas 1/2/3 · `32` estatus del certificado · `34` tarjeta verde.
- **Talabre**: `talabre.py` ya **no usa índices fijos** — resuelve las columnas por nombre de
  encabezado y detecta la fila donde empieza (el registro de DT lo movió de la fila 1 a la 9,
  y la hoja STATUS insertó tres columnas de protocolos que corrieron DT P1/P2 de I/J a L/M).
  Cosas del formato nuevo que el script ya absorbe: prioridades como dígito suelto (`1`…`4`,
  no `P1`); sin columna STATUS ni DIAS ATRASO (se calculan con la regla homologada); la
  columna ÁREA de STATUS viene con celdas combinadas (el valor se arrastra hacia abajo);
  ya no hay hoja RESUMEN (se pierde ese contraste y el script avisa); la hoja «DT» del
  archivo de STATUS es el extracto de pendientes y se usa como contraste de los abiertos.

### Discrepancias reales entre archivos (no son bugs, se informan en pantalla)
- **Desaladora**: el reporte gerencial y el punch list **no cuadran** (945/441 vs 960/426;
  P0 41 vs 20). **Manda el punch list**, que es la fuente ítem a ítem.
- **Talabre**: con los archivos viejos la hoja STATUS y el registro de DT **no cuadraban**
  (P1 29 vs 22 · P2 296 vs 265) y mandaba el registro. Con los nuevos (corte 03-08-2026)
  **cuadran exacto** (P1 29 · P2 230 · 480 abiertos = extracto de pendientes): si vuelven a
  divergir, investigar antes de publicar.
- Ambos scripts **contrastan su cálculo contra los totales que declara el propio archivo** y
  avisan si dejan de cuadrar. Si aparece un aviso nuevo, investigar antes de publicar.

### Cálculo y presentación
- **Usar `pct1()`, no `round()`.** El `round()` de Python usa redondeo bancario
  (`round(51.25,1)` = 51.2) y el panel calcula en JS con `Math.round` (51.3). `pct1()` iguala
  ambos para que panel, JSON y PPT muestren la misma cifra.
- **Talabre: el corte es `meta.hoy`**, la fecha de referencia de los atrasos. `meta.ultimaAgenda`
  guarda la caminata agendada más lejana, que es una **fecha futura** y no debe usarse como corte
  (hacía que el informe se descargara con fecha 10-08 en vez de 27-07).
- **Desaladora: el reporte gerencial trae 97 filas pero el universo es 95.** Las dos de
  diferencia son componentes (ver regla de homologación). `subsistemas.filasReporte` guarda el
  conteo bruto para que el cruce contra el archivo siga siendo trazable. El script verifica
  que esos componentes **no tengan punch asociado** —si algún día lo tienen, avisa—, porque
  entonces excluirlos escondería trabajo pendiente.

---

## 5 · PPT — formato fijo (17 láminas)

**El mazo sigue el mismo orden que las pestañas:** 1 portada Desaladora · 2-5 Desaladora ·
6 portada Talabre · 7-10 Talabre · 11 portada Arqueros · 12-17 Arqueros.
Los comentarios de sección **no llevan número de lámina**: se pudren al reordenar y la
numeración real la pone `nSlide`.

- **Logo** en todas: grande (`x9.14 y0.46 w3.88 h0.84`) en portadas y láminas azules; menor
  (`x9.97 y0.24 w3.03 h0.65`) en contenido.
- **Etiqueta del proyecto** al pie de cada lámina de contenido (`x2.0 y7.09 w10.67`, alineada a
  la derecha, 8,5 pt). Las portadas no la llevan.
- **Numeración correlativa automática** (`nSlide`), no escribirla a mano.
- **Etiquetas de gráficos**: constante `ETIQ` — blanco, negrita, 11 pt. Los colores de segmento
  deben dar **≥4,5:1 con texto blanco**: usar `goodD`, `blue`, `blueM`, `warnD`, `crit`,
  `neutral`. **Nunca** `track`, `blueL`, `warn` ni `good` como segmento con etiqueta encima:
  daban 1,3:1 a 3,4:1 y el número quedaba invisible.
- **Valores 0**: envolver en `z()` para que no se dibujen. PowerPoint los imprime pegados al
  eje, pisando la etiqueta vecina.
- El nombre de contrato de Arqueros está en `PROY_MASA` (arriba de `gen_ppt.js`).
- **`let s` se declara una sola vez**, antes del bloque de Desaladora: los tres proyectos la
  reasignan y el orden del mazo lo fija el orden de los bloques, no el de las variables.

---

## 6 · index.html — bloques marcados

Los scripts reemplazan solo su bloque; el resto del archivo se edita a mano.

| Marca | La escribe | Contiene |
|---|---|---|
| `DATOS:INICIO/FIN` | `actualizar.py` | `const D` — Arqueros (su Excel se llama MASA) |
| `DESALADORA:INICIO/FIN` | `desaladora.py` | `const DES` |
| `TALABRE:INICIO/FIN` | `talabre.py` | `const TAL` |
| `LOGO:INICIO/FIN` | `gen_ppt.js` | `const LOGOS` — logos en base64 |
| `PPT:INICIO/FIN` | `gen_ppt.js` | `const PPTX` — la PPT en base64 |

El panel es **un solo archivo portable** (~2,5 MB): logos y PPT van embebidos para que se pueda
enviar por correo tal cual. **No romper esto** agregando referencias a archivos externos.

El **resumen corporativo se calcula en el navegador** desde `D`, `DES` y `TAL` (funciones
`homologar()` y `renderCORP()`). No hay un cuarto JSON: si cambia un proyecto, el consolidado
cambia solo.

---

## 7 · Cómo verificar antes de entregar

**LibreOffice de este contenedor no abre PPTX** (falta el filtro de Impress): `soffice
--convert-to pdf` falla con «source file could not be loaded» para cualquier pptx. Verificar
leyendo el XML con `python-pptx` / `unzip`, no intentando renderizar.

Chequeos que conviene correr:

```bash
# La PPT: numeración, logo, etiqueta de proyecto, desbordes de lámina
python3 -c "from pptx import Presentation; ..."   # ver historial de la sesión

# Contraste de las etiquetas de los gráficos (leyendo ppt/charts/chart*.xml)
unzip -q -o Panel_Control_TOP_P1.pptx -d /tmp/x && grep -o 'srgbClr val="[0-9A-F]*"' ...
```

El panel se revisa con Playwright y **Chromium en
`/opt/pw-browsers/chromium-1194/chrome-linux/chrome`** (el `executable_path` por defecto de
Playwright apunta a una versión que no está instalada). Verificar en las 4 pestañas:
sin errores de consola, `scrollWidth == viewport` (sin desbordes) y que la descarga entregue
un PPTX que abra.

`box-sizing: border-box` está acotado a `.viz-root`; sin él el padding del `.wrap` se suma al
100% y la página desborda ~52 px.

---

## 8 · Identidad visual

Misma que el **dashboard corporativo de Protocolos**: base oscura `#080d18`, superficies
`#0f1623`/`#141d2e`, bordes `#1d2d44`, acento dorado `#F8BD19`, **Rajdhani** para cifras y
titulares e **IBM Plex Sans** para el cuerpo. Oscuro por defecto; el modo claro oscurece el
dorado para que contraste sobre blanco.

La rampa del heatmap va de azul profundo (pocos) a azul brillante (muchos), con el texto
invertido (`k >= 3 ? '#0b0b0b' : '#e8eeff'`) para que contraste en ambos extremos.
