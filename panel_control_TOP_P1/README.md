# Control de Cierre QAQC — Gestión de Calidad

Entregable ejecutivo multiproyecto para medir el estatus de **armado y entrega de carpetas**,
el **avance de caminatas** y los **detalles de terminación en condición P1**, con resumen
general y desglose por área/zona y disciplina.

El panel abre en una pestaña de **resumen corporativo** que consolida los tres proyectos, y
tiene además **una pestaña por proyecto** con su detalle. Cada proyecto conserva sus propios
conceptos —no se fuerza una equivalencia que no existe— y solo se consolida lo que sí es
comparable.

| Pestaña | Proyecto | Cliente | Corte | Fuente |
|---|---|---|---|---|
| **Resumen corporativo** | Los 3 proyectos | Besalco Montajes | — | se calcula desde los otros tres |
| **P2416 · Desaladora** | PV3 Ampliación Desaladora | ANTOFAGASTA MINERALS | 27-07-2026 | `REPORTE_GERENCIAL_*.xlsx` + `Listado_Puntos_Punch_Consolidado_*.xlsx` |
| **P2407 · Talabre** | Captación de agua, pozos y estaciones de bombeo | CODELCO | 03-08-2026 | `STATUS_SUBSISTEMAS_TALABRE.xlsx` + `Detalle_de_TerminacionesBesalco.xlsx` |
| **P2342 · Arqueros** | Contrato Electromecánico Planta Concentradora y de Espesado | MASA | 26-07-2026 | `Estatus_Resumen_General_QAQC.xlsx` |

**La identidad del proyecto —orden, código, nombre y cliente— sale del módulo de Protocolos**,
que es la referencia de toda la suite. La pestaña se rotula `P2416 · Desaladora` con el cliente
debajo; en gráficos y encabezados de tabla va el nombre solo, que es lo que cabe.
*Arqueros* es el proyecto y *MASA* el cliente; el Excel de origen y el id interno siguen
llamándose MASA.

---

## 🔄 Actualización semanal

**Desaladora** (2 comandos):

```bash
python3 desaladora.py \
    --reporte /ruta/REPORTE_GERENCIAL_AAAAMMDD.xlsx \
    --punch   /ruta/Listado_Puntos_Punch_Consolidado_AAAA_MM_DD.xlsx
node gen_ppt.js
```

**Talabre** (2 comandos):

```bash
python3 talabre.py \
    --status /ruta/STATUS_SUBSISTEMAS_TALABRE.xlsx \
    --dt     /ruta/Detalle_de_TerminacionesBesalco.xlsx
node gen_ppt.js
```

`talabre.py` resuelve las columnas **por nombre de encabezado** (y detecta la fila donde
empieza), así que tolera los cambios de layout de estos archivos; soporta también el formato
anterior (`TalabreSTATUS_PEC` + `TalabreCuadro_DT`). El `Programa_de_Caminatas.xlsm` **no** es
fuente: su hoja «DT» es una foto vieja del registro.

**Arqueros** (2 comandos):

```bash
python3 actualizar.py /ruta/al/Estatus_Resumen_General_QAQC.xlsx   # datos + panel HTML
node gen_ppt.js                                                    # PPT ejecutiva
```

`desaladora.py` y `talabre.py` aceptan `--hoy AAAA-MM-DD` para fijar la fecha con la que se
calculan los atrasos (por defecto usan el día en que se corren). `gen_ppt.js` regenera la PPT
completa leyendo los tres JSON, así que conviene correrlo al final, después de actualizar los
proyectos que corresponda. Si falta el JSON de un proyecto, su pestaña queda deshabilitada y la
PPT sale sin sus láminas, sin romperse.

Eso es todo. `actualizar.py` recalcula **todos** los indicadores desde el Excel, inyecta los
datos en `index.html` y registra el corte en `historial.json`; `gen_ppt.js` reconstruye la
PPT leyendo esos mismos datos. **Ninguna cifra ni texto se escribe a mano** — los titulares,
subtítulos, notas al pie y la "lectura del período" se redactan solos a partir de los números.

### Qué revisar antes de publicar

`actualizar.py` imprime un resumen, la **variación contra el corte anterior** y, si aparece un
estado nuevo en el Excel que no reconoce (por ejemplo un «Listo para cerrar» inédito), lo
lista bajo **⚠ AVISOS** en vez de clasificarlo en silencio. Si aparece un aviso, ajusta el
diccionario correspondiente al inicio del script (`CAMINATA_MAP`, `CTOP_MAP`, `DT_MAP`) y
vuelve a correrlo.

El script tolera variantes de escritura (Observada/Observado, Rechazada/Rechazado, con o sin
tilde), así que un cambio de redacción en la planilla no rompe el tablero.

---

## Contenido de la carpeta

| Archivo | Qué es |
|---|---|
| `index.html` | Panel interactivo con **una pestaña por proyecto**, en **un solo archivo portable** (se puede enviar por correo tal cual). Lleva la PPT del corte embebida y descargable con el botón **Descargar Informe**. Modo claro/oscuro, tooltips y tablas de detalle. |
| `Panel_Control_TOP_P1.pptx` | Presentación ejecutiva de 17 láminas: 5 de Desaladora + 5 de Talabre + 7 de Arqueros. |
| `actualizar.py` | **Punto de entrada de Arqueros.** Lee el Excel, recalcula todo y actualiza el panel. |
| `desaladora.py` | **Punto de entrada de Desaladora.** Lee el reporte gerencial y el punch list, valida el cruce entre ambos y actualiza el panel. |
| `talabre.py` | **Punto de entrada de Talabre.** Lee la hoja STATUS y el registro de DT (columnas por nombre), calcula el estado con la regla homologada y actualiza el panel. |
| `gen_ppt.js` | Regenera la PPT completa desde los datos ya calculados de los tres proyectos. |
| `estatus_datos.json` | Datos consolidados de Arqueros. |
| `datos_desaladora.json` | Datos consolidados de Desaladora. |
| `datos_talabre.json` | Datos consolidados de Talabre. |
| `historial.json` | Un registro por corte de Arqueros; alimenta la variación semana a semana. |

---

## Arqueros — indicadores al corte 26-07-2026

| Frente | Indicador | Valor | vs. 16-07 |
|---|---|---|---|
| Caminatas | 80% realizadas | **135/135 · 100%** | = |
| Caminatas | 100% realizadas | **90/135 · 66,7%** | = |
| Carpetas TOP | Entregadas | **27/135 · 20,0%** | +1 |
| Carpetas TOP | Aprobadas | **0** | = |
| Carpetas TOP | Rechazadas | **12** | +4 |
| Detalles P1 | Cerrados | **1.148/1.988 · 57,7%** | +1 |
| Detalles P1 | Vencidos | **405** | +4 |

**Semáforo:** 2000 *Al día* · 3000 *Crítico* · 4000 *Atención* · 8000 *Crítico*.

**Lectura:** la programación de caminatas se destrabó (los «por programar» cayeron de 36 a 1),
pero la **ejecución no se movió**: las realizadas al 100% siguen en 90. En carpetas TOP los
rechazos crecen más rápido que las entregas y aún no hay ninguna aprobada. El cierre de P1
está prácticamente detenido (+1 en el período) mientras la deuda vencida sube a 405.
El área **3000** concentra 707 de los 830 P1 abiertos (85,2%).

---

## Desaladora — indicadores al corte 27-07-2026

**95 subsistemas** (80 Operables · 15 Facility) en dos zonas:
Planta Desaladora (74) y Estación de Bombeo 2 (21).

El reporte gerencial trae **97 filas**: las otras dos son **componentes ya entregados**
(`0587-ESL-201 Comp 1` y `4515-ESL-201-Comp 1`), entregas parciales de un subsistema que ya
está contado. No suman al universo —sumarlos duplicaría alcance— y se informan aparte en
«Control de calidad del dato». El universo de 95 es el que declara la propia hoja
«Resumen general» del archivo: 80 operables y 15 facility.

| Frente | Indicador | Valor |
|---|---|---|
| Caminatas | Caminata 2 — Operables | **41/80 · 51,3%** |
| Caminatas | Caminata 1 — Facility | **10/15 · 66,7%** |
| Carpetas | Entregadas al cliente | **37/95 · 38,9%** |
| Carpetas | Aprobadas / Rechazadas / En preparación | **20 / 9 / 14** |
| Punch P1 | Cerrados | **592/829 · 71,4%** |
| Punch P1 | Abiertos (de ellos, atrasados) | **237 (178)** |
| Punch total | Cerrados | **960/1.386 · 69,3%** |
| Punch total | Abiertos (de ellos, atrasados) | **426 (332)** |

**Semáforo:** Planta Desaladora *Crítico* · EB-2 *Atención*.

**Lectura:** el punch cierra a buen ritmo (69,3%), pero **el 78% de lo que sigue abierto ya
está atrasado** (332 de 426). El atraso es reciente —mediana de 22 días, máximo 78— así que
todavía es recuperable. **Piping** es el cuello de botella real: con solo 153 P1 levantados
acumula 97 abiertos (37% de cierre), mientras Eléctrica, con 587 levantados, ya cerró el 86%.
En carpetas, de las 58 que aún no llegan al cliente, 14 están en preparación y 44 sin iniciar:
el freno está en el armado interno, no en la revisión del mandante.

### Criterios propios de Desaladora

- **Caminatas:** se identifican por **número**, nunca por su porcentaje, porque el avance que
  representa cada caminata cambia entre proyectos. Cada tipo de subsistema se mide con la suya:
  **Operables por la Caminata 2** y **Facility por la Caminata 1**, tal como lo declara la hoja
  «Resumen general» del propio reporte. El script **contrasta su cálculo contra esos totales
  declarados** y avisa si dejan de cuadrar.
- **Atrasado** = ítem **abierto** cuya *Fecha Requerida Cierre* ya venció a la fecha de
  referencia. Los abiertos **sin** fecha requerida (29) no se dan por atrasados: se informan
  aparte para no inflar el indicador.
- **Categorías:** `P1A → P1`, `P2B → P2`, `P3C → P3`, y `P0` se mantiene aparte, de modo que
  las cifras sean comparables con los otros proyectos.
- **Carpeta entregada** = certificado de entrega *En revisión*, *Aprobada* o *Rechazada*.
  «En preparación» (Proceso BSMT) **no** cuenta: todavía es interna y no salió a cliente.

### Control de calidad del dato

El panel trae una sección **«Control de calidad del dato»** que deja a la vista el cruce entre
los dos archivos de origen. Al corte actual:

- Las caminatas **cuadran exactamente** con lo que declara el reporte gerencial
  (Operables 41/80 y Facility 10/15).
- El **conteo de punch difiere** entre los dos archivos: el reporte gerencial totaliza
  945 cerrados / 441 abiertos, y el punch list 960 / 426. También difieren las categorías
  (P0: 41 vs 20). **El panel usa el punch list**, que es la fuente ítem a ítem; la diferencia
  queda documentada en pantalla en vez de resolverse en silencio.
- 45 de los 95 subsistemas no tienen ningún punch levantado.
- Los 2 componentes que quedan fuera del universo **no tienen ningún punch asociado**:
  excluirlos no saca de la vista ningún pendiente. El script lo verifica y avisa si algún
  día lo tuvieran.

---

## Talabre — indicadores al corte 03-08-2026

**218 subsistemas** en 9 áreas (Pozos 110, EBMS 25, Piscinas y Sedimentador 22, EBMP 20,
Líneas Impulsión 14, PQS 13, Sala GAR 8, Cerro Verde/PS2 3, Impulsión Aducción 3).

Desde este corte el proyecto entrega **el registro completo de DT**
(`Detalle_de_TerminacionesBesalco.xlsx`: 2.024 registros desde 09-2025) en lugar del extracto
anterior, y por primera vez **la hoja STATUS y el registro cuadran exacto**.

| Frente | Indicador | Valor |
|---|---|---|
| Caminatas | Caminata 1 — sobre lo exigible | **205/208 · 98,6%** (94,0% del universo) |
| Caminatas | C1: programadas en plazo / agenda vencida / sin programar | **10 / 2 / 1** |
| Caminatas | Caminata 2 — sobre lo exigible | **192/193 · 99,5%** (88,1% del universo) |
| Caminatas | C2: programadas en plazo / sin programar | **25 / 1** |
| Carpetas | Avance PEC promedio | **92,7%** |
| Carpetas | Sobre 95% / bajo 80% | **181 / 9** |
| DT P1 | Cerrados | **323/352 · 91,8%** |
| DT P1+P2 | Cerrados | **1.467/1.726 · 85,0%** — 298 P3/P4 quedan fuera |
| DT P1+P2 | Abiertos (de ellos, atrasados) | **259 (156)** |

**Semáforo:** Piscinas y Sedimentador *Crítico* · Pozos, Impulsión Aducción y Sala GAR
*Al día* · EBMP, EBMS, Líneas Impulsión y PQS en *Atención* · Cerro Verde/PS2 *Sin base al
corte* (no tiene ningún detalle levantado ni ninguna caminata vigente realizada).

**El avance de caminata se mide sobre lo exigible**, descontando las agendadas a fecha futura
—que no son incumplimiento al corte— e informándolas aparte como «programadas en plazo». La
agendada cuya fecha **ya pasó** sin marcarse realizada queda **«con agenda vencida»**: sigue
siendo exigible, porque descontarla escondería un incumplimiento real.

**Lectura:** el perfil de Talabre es **el inverso al de Desaladora**. Aquí P1 está
prácticamente cerrado (91,8%) y el volumen pendiente está en las prioridades menores: 230 DT
abiertos en P2, 171 en P3 y **50 en P4 sin ni un solo cierre**. Las caminatas están al día
sobre lo exigible (98,6% y 99,5%): lo que falta ya tiene fecha dentro de agosto, con el grueso
de las agendas en **Piscinas y Sedimentador** (12) y las 3 de Cerro Verde/PS2. El atraso
pendiente es de mediana 31 días con 21 ítems sobre los 90. Las carpetas avanzan bien
—promedio 92,7%— con solo 9 subsistemas bajo 80%.

### Criterios propios de Talabre

- **Estado del DT:** el registro nuevo **no trae columna STATUS**, así que `talabre.py`
  **calcula** el estado con la regla homologada de los tres proyectos: *Cerrado* si hay fecha
  de cierre; *Atrasado* si está abierto y la fecha de compromiso ya venció; *Abierto* el
  resto. (Si el archivo volviera a traer su propio STATUS, manda él y el script verifica
  registro por registro que respete esa misma regla, como hacía antes.)
- Los **218 DT abiertos sin fecha de compromiso** no se dan por atrasados; se informan aparte.
- **Caminatas:** `SI` = realizada · fecha aún por llegar = programada en plazo · fecha ya
  pasada = **agenda vencida** · `NO` o vacío = sin programar.
- **Carpetas:** Talabre **no maneja estados de aprobación** como los otros dos proyectos: lleva
  un **% de avance por subsistema (PEC)**. Se reporta como avance y por tramos, sin traducirlo
  a «entregadas», para no inventar un concepto que el proyecto no usa.
- **Disciplinas:** `CANERIAS → Piping` · `CIVIL` y `MOVIMIENTO DE TIERRAS → Obras Civiles` ·
  `ELECTRICOS → Eléctrica` · `INSTRUMENTACION → Instrumentación y Control`.
- **Áreas:** las dos fuentes abren los pozos uno a uno (PBO-xx, PBBR-xx, PBN-01, TB-01…): se
  agrupan en **POZOS**, que es como siempre se reportó. **LÍNEAS IMPULSIÓN y PQS reportan lo
  suyo** (151 y 145 DT): la versión anterior, que derivaba el área desde el sistema, las metía
  dentro de POZOS. «ADUCCIÓN» se rotula «IMPULSIÓN ADUCCIÓN».
- **Prioridades:** el registro nuevo las trae como dígito suelto (`1`…`4`); se leen igual que
  `P1`…`P4`.

### Control de calidad del dato

- **Cuadran:** DT abiertos contra el extracto de pendientes del archivo de STATUS
  (**480 = 480**) y el conteo por subsistema de la hoja STATUS contra el registro
  (**P1 29 = 29 · P2 230 = 230**) — con los archivos anteriores esto no cuadraba.
- **Se pierde:** el archivo nuevo ya no trae la hoja «RESUMEN», así que el contraste contra
  los totales que declaraba el proyecto ya no es posible; el script lo avisa.

---

## Definiciones y criterios (Arqueros)

- **Abierto** = *Trabajo requerido* + *Iniciado* + *Trabajo no aceptado*.
- **En trámite** = *Listo para revisión* o *Listo para cerrar* — **no** cuenta como cerrado.
- **Vencido** = ítem abierto con *Fecha de vencimiento* anterior a la fecha de corte.
  Ojo: al avanzar la fecha de corte, ítems que estaban en plazo pueden pasar a vencidos sin
  que nadie haya hecho nada — parte del alza semanal se explica así.
- **% cierre P1** = Cerrados / Total P1 levantados.
- **Carpeta TOP entregada** = con estatus *En revisión*, *Observada*, *Rechazada* o *Aprobada*.
- **Semáforo:** *Crítico* si caminata 100% < 60% **o** cierre P1 < 50% · *Atención* si algún
  frente queda entre 60–90% o restan carpetas por entregar · *Al día* si caminatas ≥ 90% y P1 ≥ 90%.

> El criterio del semáforo lo aplica el script de forma uniforme. Por eso el área **8000**
> figura como *Crítico* (caminata 100% = 50%, bajo el umbral de 60%) aunque su cierre de P1
> sea alto: en la versión manual previa estaba marcada como *Atención*, lo que era
> inconsistente con el criterio declarado.

## Resumen corporativo — qué se consolida y qué no

La primera pestaña compara los tres proyectos. La homologación se calcula en el propio panel
desde los datos de cada proyecto, así que **no hay cifras duplicadas**: si cambia un proyecto,
el consolidado cambia solo.

| Concepto | Desaladora | Talabre | Arqueros | Cómo se consolida |
|---|---|---|---|---|
| **Caminatas** | «80%» y «100%» | Caminata 1, 2 y 3 | CAMINATA 1 y 2 | Por **número**, nunca por el porcentaje que representan. Se compara la *caminata vigente* que cada proyecto declara como indicador |
| **Detalles** | Detalles de terminación | Punch list | DT | Son lo mismo: **se suman** |
| **Prioridad** | P1/P2/P3 | P0/P1A/P2B/P3C | P1/P2/P3/P4 | Al dígito: P1A→P1, P2B→P2, P3C→P3. P0 y P4 aparte |
| **Atraso** | Vencido | Atrasado | Atrasado | Mismo concepto — abierto con su fecha comprometida cumplida: **se suman** |
| **Carpetas** | Estatus CTOP | Certificado de entrega | % PEC (avance) | **No se consolidan** (ver abajo) |
| **Disciplinas** | Piping, Eléctrica… | Piping, Eléctrica… | CANERIAS, ELECTRICOS… | Se unifican los sinónimos |

**Las carpetas no se suman a propósito.** Desaladora y Arqueros miden *estado de aprobación* ante
el cliente; Talabre mide *porcentaje de avance* de la carpeta. Un promedio entre ambas cosas no
significaría nada, así que la pestaña muestra la métrica nativa de cada proyecto, etiquetada, y
solo consolida a los dos que sí son comparables (66 de 232 carpetas entregadas, 28,4%).

**Caminata vigente** es la que cada proyecto usa como su indicador: Arqueros la Caminata 2 (su
«100%»), Talabre la Caminata 2, y Desaladora una por tipo de subsistema (Operables por la 2,
Facility por la 1), sumando lo realizado de cada tipo.

### Consolidado al corte actual

| Indicador | Valor |
|---|---|
| Subsistemas en control | **448** (Desaladora 95 · Talabre 218 · Arqueros 135) |
| Caminata vigente realizada | **333/448 · 74,3%** |
| Cierre de detalles P1+P2 | **3.701/5.846 · 63,3%** — 527 P3/P4 quedan fuera |
| Detalles de terminación levantados | **6.373** |
| Cierre de detalles (todas las prioridades) | **3.819/6.373 · 59,9%** |
| Cierre de detalles P1 | **2.063/3.169 · 65,1%** |
| Detalles abiertos / de ellos atrasados | **2.469 / 1.238 (50,1%)** |

**Semáforo:** Desaladora *Atención* · Talabre *Al día* · Arqueros *Crítico*.

## Descargar el informe desde el panel

El panel trae un botón **📊 Descargar Informe** en la cabecera que entrega la PPT del corte
**sin necesidad de servidor ni de archivos sueltos**: `gen_ppt.js` la genera y la deja embebida
dentro del propio `index.html`, así que el panel se puede enviar por correo y quien lo reciba
descarga la presentación desde ahí.

El archivo se descarga como `Panel_CRP_DDMMAA.pptx`, con el corte más reciente de los tres
proyectos, para que informes de semanas distintas no se pisen en la carpeta de descargas.

Esto suma unos 2,3 MB al panel (queda en ~2,5 MB). Si todavía no se ha corrido `gen_ppt.js`,
el botón aparece deshabilitado explicando qué falta, en vez de fallar al hacer clic.

## Formato de la PPT (fijo)

La presentación tiene **17 láminas** con esta estructura, que se mantiene corte a corte:

| Láminas | Contenido |
|---|---|
| 1 | Portada **Proyecto 01 · Desaladora** |
| 2–5 | Desaladora: resumen, frentes 1 y 2, frente 3, foco de gestión |
| 6 | Portada **Proyecto 02 · Talabre** |
| 7–10 | Talabre: resumen, frentes 1 y 2, frente 3, foco de gestión |
| 11 | Portada **Proyecto 03 · Arqueros** |
| 12–17 | Arqueros: resumen, semáforo, caminatas, carpetas TOP, detalles P1, foco de gestión |

Es el mismo orden de las pestañas del panel y de los otros dos módulos de la suite.

Reglas de formato que `gen_ppt.js` aplica solo:

- **Logo corporativo** en todas las láminas, desde `besalco_logo.png`: grande
  (3,88 × 0,84 pulg. en `x9.14 y0.46`) en portadas y láminas de fondo azul; menor
  (3,03 × 0,65 en `x9.97 y0.24`) en las de contenido. Si el archivo no está, la PPT
  se genera igual, sin logo.
- **Código del proyecto** en el kicker de cada portada (`PROYECTO 01 · P2416`).
- **Etiqueta del proyecto** al pie de cada lámina de contenido (`PROYECTO DESALADORA`,
  `PROYECTO TALABRE`, `PROYECTO ARQUEROS`), alineada a la derecha en 8,5 pt. Las
  portadas no la llevan porque ya muestran el nombre en grande.
- **Numeración correlativa** automática: al agregar o quitar láminas no hay que
  renumerar a mano.
- **Etiquetas de los gráficos**: número blanco en negrita a 11 pt sobre segmentos
  cuyo color garantiza contraste ≥ 4,5:1. Los valores 0 **no** se dibujan, para que
  no quede un «0» pegado al eje pisando la etiqueta vecina.

> El nombre de contrato que aparece en la portada de Arqueros está en `PROY_MASA`, al
> inicio de `gen_ppt.js`; si el JSON llega a traer `meta.proyecto`, se usa ese.

## Identidad visual

El panel usa el **mismo sistema de diseño que el dashboard corporativo de Protocolos** de
Besalco Montajes: base oscura (`#080d18`), acento dorado corporativo (`#F8BD19`), borde
superior en degradado azul→dorado, tipografía **Rajdhani** para cifras y titulares e **IBM Plex
Sans** para el cuerpo. Las pestañas, tarjetas KPI (con su barra de acento superior), cabeceras
de bloque con filete dorado y tablas replican los componentes de ese dashboard, de modo que los
dos tableros se lean como un mismo producto.

El **modo oscuro es el predeterminado**. Se mantiene un modo claro —con el dorado oscurecido
para que contraste sobre blanco— para imprimir o proyectar sobre fondo claro.

No hay dependencias externas: el panel sigue siendo **un solo archivo portable** y las
tipografías degradan a la fuente del sistema si no están instaladas.

## Nota de accesibilidad (colores)

La paleta fue verificada con un validador de contraste y daltonismo. Los estados de las
carpetas TOP usan colores de **estado** (verde/ámbar/azul/rojo), que bajo daltonismo severo
no se distinguen solo por tono; por eso la identidad **nunca depende del color**: cada gráfico
lleva leyenda, valores numéricos impresos sobre cada segmento y una tabla de detalle.

## Nota de reconciliación

Un bloque intermedio del Excel muestra «Total CTOP = 134» (excluye un alcance sin código de
subsistema), mientras que su indicador de cumplimiento usa el universo completo de **135**.
Este panel adopta 135 como base única para los tres frentes; por eso puede haber diferencias
de una décima respecto de ese bloque intermedio.

## Logo

El logo de Besalco Montajes va **arriba a la izquierda**, sin recuadro, igual que en el
dashboard de Protocolos. La carpeta guarda dos versiones:

| Archivo | Uso |
|---|---|
| `besalco_logo.png` | Versión de color — modo claro del panel y láminas de la PPT |
| `besalco_logo_blanco.png` | Versión blanca — modo oscuro del panel (el predeterminado) |

`gen_ppt.js` los **embebe en `index.html`** como data URI, así el panel se puede enviar por
correo sin adjuntar la imagen aparte, y el panel elige la versión que corresponde al tema
activo. Si algún día cambia el logo, basta reemplazar los PNG y volver a correr `gen_ppt.js`.
La versión blanca se obtiene de la de color conservando su silueta:

```bash
python3 -c "from PIL import Image; im=Image.open('besalco_logo.png').convert('RGBA'); \
a=im.split()[3]; Image.merge('RGBA',(a.point(lambda _:255),)*3+(a,)).save('besalco_logo_blanco.png')"
```

Si no hay ningún logo, el espacio se oculta solo, sin romper el diseño.
