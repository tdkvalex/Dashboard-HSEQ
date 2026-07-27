# Panel de Control — Carpetas · Caminatas · Detalles de Terminación P1

Entregable ejecutivo multiproyecto para medir el estatus de **armado y entrega de carpetas**,
el **avance de caminatas** y los **detalles de terminación en condición P1**, con resumen
general y desglose por área/zona y disciplina.

El panel tiene una **pestaña por proyecto**. Cada uno conserva sus propios conceptos —no se
fuerza una equivalencia que no existe— y solo se homologa lo que sí es comparable.

| Pestaña | Proyecto | Cliente | Corte | Fuente |
|---|---|---|---|---|
| MASA | Planta Concentradora | MASA | 26-07-2026 | `Estatus_Resumen_General_QAQC.xlsx` |
| Desaladora | PV3 Ampliación Desaladora | Minera Los Pelambres (AMSA) | 27-07-2026 | `REPORTE_GERENCIAL_*.xlsx` + `Listado_Puntos_Punch_Consolidado_*.xlsx` |
| Talabre | Captación de agua, pozos y estaciones de bombeo | Codelco | 27-07-2026 | `TalabreSTATUS_PEC.xlsx` + `TalabreCuadro_DT.xlsx` |

---

## 🔄 Actualización semanal

**MASA** (2 comandos):

```bash
python3 actualizar.py /ruta/al/Estatus_Resumen_General_QAQC.xlsx   # datos + panel HTML
node gen_ppt.js                                                    # PPT ejecutiva
```

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
    --status /ruta/TalabreSTATUS_PEC.xlsx \
    --dt     /ruta/TalabreCuadro_DT.xlsx
node gen_ppt.js
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
| `index.html` | Panel interactivo con **una pestaña por proyecto**, en **un solo archivo portable** (se puede enviar por correo tal cual). Modo claro/oscuro, tooltips y tablas de detalle. |
| `Panel_Control_TOP_P1.pptx` | Presentación ejecutiva de 17 láminas: 7 de MASA + 5 de Desaladora + 5 de Talabre. |
| `actualizar.py` | **Punto de entrada de MASA.** Lee el Excel, recalcula todo y actualiza el panel. |
| `desaladora.py` | **Punto de entrada de Desaladora.** Lee el reporte gerencial y el punch list, valida el cruce entre ambos y actualiza el panel. |
| `talabre.py` | **Punto de entrada de Talabre.** Lee la hoja STATUS y el registro de DT, verifica la regla de atraso y actualiza el panel. |
| `gen_ppt.js` | Regenera la PPT completa desde los datos ya calculados de los tres proyectos. |
| `estatus_datos.json` | Datos consolidados de MASA. |
| `datos_desaladora.json` | Datos consolidados de Desaladora. |
| `datos_talabre.json` | Datos consolidados de Talabre. |
| `historial.json` | Un registro por corte de MASA; alimenta la variación semana a semana. |

---

## MASA — indicadores al corte 26-07-2026

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

**97 subsistemas** (80 Operables · 15 Facility · 2 Componentes) en dos zonas:
Planta Desaladora (75) y Estación de Bombeo 2 (22).

| Frente | Indicador | Valor |
|---|---|---|
| Caminatas | Caminata 2 — Operables | **41/80 · 51,3%** |
| Caminatas | Caminata 1 — Facility | **10/15 · 66,7%** |
| Carpetas | Entregadas al cliente | **39/97 · 40,2%** |
| Carpetas | Aprobadas / Rechazadas / En preparación | **22 / 9 / 14** |
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
- 47 de los 97 subsistemas no tienen ningún punch levantado.

---

## Talabre — indicadores al corte 27-07-2026

**219 subsistemas** en 9 áreas (Pozos 110, EBMS 26, Piscinas y Sedimentador 22, EBMP 20,
Líneas Impulsión 14, PQS 13, Sala GAR 8, Cerro Verde/PS2 3, Impulsión Aducción 3).

| Frente | Indicador | Valor |
|---|---|---|
| Caminatas | Caminata 1 realizada | **199/219 · 90,9%** |
| Caminatas | Caminata 2 realizada | **170/219 · 77,6%** |
| Carpetas | Avance PEC promedio | **91,1%** |
| Carpetas | Sobre 95% / bajo 80% | **159 / 10** |
| DT P1 | Cerrados | **313/335 · 93,4%** |
| DT total | Cerrados | **1.446/1.951 · 74,1%** |
| DT total | Abiertos (de ellos, atrasados) | **505 (206)** |

**Semáforo:** Cerro Verde/PS2 y Piscinas y Sedimentador *Críticos* · Impulsión Aducción y
Sala GAR *Al día* · el resto en *Atención*.

**Lectura:** el perfil de Talabre es **el inverso al de Desaladora**. Aquí P1 está
prácticamente cerrado (93,4%) y el volumen pendiente está en las prioridades menores: 265 DT
abiertos en P2, 168 en P3 y **50 en P4 sin ni un solo cierre**. El atraso, en cambio, es más
viejo que en los otros proyectos: mediana de 36 días y 29 ítems sobre los 90 días. En
caminatas, la 1 está casi completa pero la 2 va en 77,6%, concentrando el rezago en
**Piscinas y Sedimentador** (5/22) y **Cerro Verde/PS2** (0/3). Las carpetas avanzan bien
—promedio 91,1%— con solo 10 subsistemas bajo 80%.

### Criterios propios de Talabre

- **Estado del DT:** Talabre **ya trae calculada** la columna *STATUS* (Cerrado / Abierto /
  Atrasado). El panel la usa tal cual, pero `talabre.py` **verifica registro por registro** que
  respete la misma regla que el resto de los proyectos (atrasado = sin fecha de cierre y con
  fecha de compromiso vencida) y avisa si alguno no cuadra. **Al corte actual los 1.951
  registros la respetan.**
- Los **215 DT abiertos sin fecha de compromiso** no se dan por atrasados; se informan aparte.
- **Caminatas:** `SI` = realizada · una fecha = agendada para esa fecha · `NO` = sin programar.
  La hoja RESUMEN publica las caminatas «programadas», que **suman las realizadas más las
  agendadas de la semana** (209 y 172). Este panel las informa por separado, porque una
  caminata agendada no es una caminata hecha.
- **Carpetas:** Talabre **no maneja estados de aprobación** como los otros dos proyectos: lleva
  un **% de avance por subsistema (PEC)**. Se reporta como avance y por tramos, sin traducirlo
  a «entregadas», para no inventar un concepto que el proyecto no usa. El archivo declara
  además 0 carpetas en revisión del cliente.
- **Disciplinas:** `CANERIAS → Piping` · `CIVIL` y `MOVIMIENTO DE TIERRAS → Obras Civiles` ·
  `ELECTRICOS → Eléctrica` · `INSTRUMENTACION → Instrumentación y Control`.
- **Áreas:** el registro de DT usa códigos más finos que la hoja STATUS (PBO-xx, PBBR-xx,
  PBN-01, TB-01…). Todos pertenecen al sistema «Pozos Barrera Hidráulica» y se agrupan en
  **POZOS**, igual que en STATUS.

### Control de calidad del dato

- **Cuadran:** total de subsistemas (219) y carpetas sobre 95% (159).
- **Difiere:** el RESUMEN declara **206** carpetas sobre 80% y el detalle da **209**.
- **Difiere:** los DT abiertos por subsistema de la hoja STATUS no coinciden con el registro
  de DT (P1: 29 vs 22 · P2: 296 vs 265). **El panel usa el registro de DT**, que es la fuente
  ítem a ítem.

---

## Definiciones y criterios (MASA)

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

El panel muestra el logo de Besalco Montajes si existe un archivo `besalco_logo.png` en esta
carpeta; si no está, el espacio se oculta automáticamente sin romper el diseño.
