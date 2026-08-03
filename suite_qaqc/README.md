# Suite QAQC — Besalco Montajes

Un solo archivo (`Suite_QAQC.html`) que reúne los dashboards de calidad en una consola
ejecutiva: una portada que cruza **proyectos × módulos** y, dentro, cada dashboard completo.

| Módulo | Qué mide | Escala | Estado |
|---|---|---|---|
| **Protocolos** | Registros constructivos | % pendiente — **menos es mejor** | activo |
| **Cierre QAQC** | Carpetas · caminatas · detalles de terminación | % cierre — **más es mejor** | activo |
| **No Conformidades** | Hallazgos de calidad y su corrección | % cerrado — **más es mejor** | activo |

Cierre QAQC y No Conformidades traen además su **PPT ejecutiva embebida**: el botón «Descargar
Informe» de cada módulo entrega el mazo del corte (17 y 16 láminas) sin salir de la suite.

Los tres cubren **los mismos tres proyectos**, y los tres los nombran igual: el código, el
nombre y el cliente salen del módulo de **Protocolos**, que es la referencia de identidad.

| Código | Nombre | Cliente |
|---|---|---|
| `P2416` | Desaladora | ANTOFAGASTA MINERALS |
| `P2407` | Talabre | CODELCO |
| `P2342` | Arqueros | MASA |

Se rotulan **`P2416 · Desaladora`** en la pestaña de cada módulo, en la matriz de la portada y
en las portadas de las dos PPT. La tabla vive en `armar_suite.js` (`PROYECTOS`), en
`panel_control_TOP_P1/index.html` (`PROY`) y en `modulo_nc/no_conformidades.py` (`PROYECTOS`).

---

## Actualización

```bash
# 1 · actualizar el módulo Cierre QAQC (ver panel_control_TOP_P1/CLAUDE.md)
cd ../panel_control_TOP_P1
python3 desaladora.py --reporte <REPORTE_GERENCIAL_*> --punch <Listado_Puntos_Punch_*>
python3 talabre.py    --status  <STATUS_SUBSISTEMAS_TALABRE> --dt <Detalle_de_TerminacionesBesalco>
python3 actualizar.py <Estatus_Resumen_General_QAQC.xlsx>   # Arqueros
node gen_ppt.js

# 2 · actualizar el módulo No Conformidades (ver modulo_nc/README.md)
cd ../modulo_nc
python3 no_conformidades.py --data <Data_NCR.xlsx> --externas <Data_NCR*externas.xlsx>
node gen_ppt.js

# 3 · dejar el dashboard de Protocolos actualizado y rearmar la suite
cd ../suite_qaqc
cp <dashboard_protocolos_actualizado>.html modulos/protocolos.html
node armar_suite.js
```

En los dos módulos el `node gen_ppt.js` va **al final**: regenera la PPT y la embebe en el
`index.html` del módulo, que es lo que después empaqueta la suite. Correrlo antes deja el
informe descargable desfasado respecto de las pestañas.

`armar_suite.js` hace todo lo demás: extrae el KPI de cada proyecto en cada módulo, escribe
`kpis_suite.json` y ensambla `Suite_QAQC.html` con la portada y los módulos dentro.

El módulo Cierre QAQC **no hay que copiarlo**: si no está en `modulos/`, el generador lo toma
directo de `../panel_control_TOP_P1/index.html`. Solo Protocolos hay que dejarlo, porque lo
genera otro equipo.

Si falta un módulo, avisa y la suite sale sin él: su tarjeta queda apagada, la matriz muestra
«Sin módulo» y los demás siguen funcionando.

---

## Cómo está armada

**Cada módulo vive en su propio documento**, montado en un `iframe` desde un `Blob`. Eso da
aislamiento total de CSS y JS: ni el `*{margin:0}` global de Protocolos ni sus variables
(`const D`, `render()`, `$`…) pueden chocar con las del otro módulo ni con las del shell.
Fue la decisión clave: permite integrarlos **sin reescribir ninguno**.

**Los módulos se montan al abrir su pestaña**, no al cargar la página. Por eso la portada abre
en ~0,4 s aunque el archivo pese 5,6 MB.

**Dos franjas y nada más.** Arriba, la de la suite (41 px) con el logo, los módulos y el
corte. Debajo, la del módulo (48 px) con su título centrado y, en la misma línea, su corte y
sus botones a la derecha.

Al empaquetar, cada módulo recibe un `<style>` que normaliza su cabecera:

- **oculta su logo** — el corporativo va una sola vez, en la franja de la suite;
- **centra su título respecto del ancho completo** (posición absoluta, no dentro del hueco que
  dejan los botones, que miden distinto en cada módulo) y lo lleva a la misma tipografía en
  todos, para que no salte de tamaño ni de lugar al cambiar de módulo;
- **ancla el corte y los botones a la derecha** con `margin-left:auto`, que funciona sin
  importar cuántos elementos queden visibles;
- **fija la altura en 48 px** — con `min-height` la franja quedaba 2 px desalineada entre un
  módulo y otro, porque sus botones miden distinto.

El título lleva `pointer-events:none` para no tapar los botones que quedan debajo. Las
pestañas de proyecto de cada dashboard quedan intactas.

**Ancho del contenido.** Protocolos ocupa todo el ancho disponible; Cierre QAQC y No
Conformidades usan `max-width: clamp(1680px, 88vw, 2160px)` en su `.wrap` — el mismo valor en
los dos, para que no salten de tamaño al cambiar de módulo. Hasta ~1900 px se comportan como
antes; de ahí en adelante se ensanchan hasta 2160 px, dejando siempre unos 200 px de margen.
El tope fijo anterior (1680 px) desperdiciaba casi 470 px por lado en una pantalla de 2560.

**El corte de la franja de la suite solo se muestra en el Resumen** (`body[data-vista]` +
la regla de `.hd-r`): dentro de un módulo lo declara el módulo, que es quien sabe a qué fecha
está lo que se ve. Es la única forma de que no aparezca dos veces ni se contradiga cuando el
módulo cambia de corte al cambiar de pestaña.

**La portada no lee los módulos en vivo**: `armar_suite.js` extrae los KPI en el momento de
armar y los deja en un bloque `KPIS`. Así la portada no depende de que los iframes estén
montados ni de comunicación entre documentos.

### «Detalle por proyecto» — cada columna dice de qué módulo es

Es la tabla de cifras crudas y mezcla columnas de los tres módulos. Sin rotularlo, «KPI» o
«atrasados» no significan nada: ¿atrasados de protocolos, de detalles, de no conformidades?
Por eso lleva **una fila de cabecera que agrupa por módulo** —con su color, el mismo de su
pestaña, y su escala— y una línea vertical donde empieza cada grupo. Los títulos van completos
(`Protocolos en falta`, `Detalles atrasados`, `Hallazgos abiertos`) para que la columna se
entienda aunque se lea sola.

**Los dos porcentajes no se dividen sobre la columna de al lado**, y eso hay que decirlo o
alguien va a sacar la calculadora y no le va a dar:

- `% pendiente` se calcula sobre la **base del KPI**, no sobre el universo: el remanente (`S`)
  suma al universo pero no puntúa. Talabre: 133 de **11.550**, no de 13.898.
- `% cierre de P1` mide **solo P1**, mientras que `Detalles levantados/cerrados/atrasados` son
  de **todas las prioridades**. Desaladora: 592/829 P1, no 960/1.386.

El denominador real de cada uno va en el `title` de la celda —se ve al pasar el cursor— y las
dos excepciones están además en la nota al pie de la portada.

### El detalle que hay que respetar

Los módulos van dentro de `<script type="text/plain">`. Lo único que corta ese bloque es la
secuencia `</script`, así que se reemplaza por un centinela (`@@CIERRE_SCRIPT_a7f3@@`) que el
shell restaura al montar.

**No se puede escapar como `<\/script`**: los módulos ya traen esa secuencia dentro de sus
propias cadenas JS (Protocolos tiene una), y al restaurar se convertirían en `</script` reales,
cerrando el bloque antes de tiempo y rompiendo el módulo. `armar_suite.js` verifica que el
centinela no exista en el origen y falla ruidosamente si algún día aparece.

---

## Archivos

| Archivo | Qué es |
|---|---|
| `Suite_QAQC.html` | **El entregable.** Un solo archivo, ~8,2 MB, sin dependencias externas |
| `armar_suite.js` | Extrae los KPI y ensambla la suite. **Punto de entrada** |
| `plantilla_suite.html` | Shell ejecutivo: portada, navegación y montaje de módulos |
| `modulos/protocolos.html` | Dashboard de Protocolos, tal cual |
| `modulos/cierre_qaqc.html` | Opcional: si no está, se usa `panel_control_TOP_P1/index.html` |
| `modulos/no_conformidades.html` | Opcional: si no está, se usa `modulo_nc/index.html` |
| `kpis_suite.json` | KPI extraídos, para consultar sin abrir la suite |
| `besalco_logo*.png` | Logo corporativo (versión blanca para fondo oscuro) |

---

## Consolidado al corte

| Módulo | Corte | Indicador | Valor |
|---|---|---|---|
| Protocolos | 03-08-2026 | % pendiente | **0,4%** — 163 en falta de 44.878 |
| Cierre QAQC | 03-08-2026 | % detalles cerrados | **59,9%** — 1.238 atrasados |
| No Conformidades | 03-08-2026 | % hallazgos cerrados | **89,2%** — 54 abiertos de 500 |

El corte de Cierre QAQC es el más reciente de sus tres proyectos (Talabre, 03-08); Desaladora
sigue al 27-07 y Arqueros al 26-07 dentro del módulo.

**Cada módulo se actualiza por su lado**, así que no siempre quedan a la misma fecha.

**El corte se muestra una sola vez, y siempre el que corresponde a lo que se está mirando.**
En el **Resumen**, la franja de la suite muestra el corte más reciente de los tres módulos y,
si difieren, agrega el aviso **«cortes mixtos»** con el detalle al pasar el cursor. **Al abrir
un módulo esa franja se oculta** y manda el corte del propio módulo. Antes se veían las dos
franjas con la fecha repetida y, peor, contradiciéndose: la de arriba decía 03-08 mientras la
pestaña de Desaladora, abajo, decía 27-07.

Dentro de Cierre QAQC el corte **cambia por pestaña** (Desaladora 27-07 · Talabre 03-08 ·
Arqueros 26-07), porque cada proyecto entrega sus archivos cuando puede. Su pestaña
corporativa, que no es un proyecto, muestra el más reciente con el mismo aviso «cortes
mixtos», para que ese 03-08 no se lea como la fecha de los tres.

**Lo que muestra el cruce:** la documentación va muy por delante del cierre físico. Protocolos
está prácticamente al día mientras el cierre de detalles P1 arrastra 594 ítems atrasados. El
cuello de botella no está en el papel: está en terreno.

---

## Validación al corte actual

Cada cifra de la portada se contrastó contra lo que muestra el propio módulo:

| Control | Suite | Módulo |
|---|---|---|
| Protocolos · universo | 62.728 | 62.728 |
| Protocolos · KPI corporativo | 0,4% | 0,4% |
| Protocolos · estado (P) | 36 | 36 |
| Protocolos · estado (AP) | 127 | 127 |
| Protocolos · estado (AE) | 5.426 | 5.426 |
| Protocolos · estado (C) | 39.289 | 39.289 |
| Cierre QAQC · P1 | 65,1% | 65,1% |
| Cierre QAQC · detalles | 6.373 · 1.238 atrasados | 6.373 · 1.238 |

### Cómo se calcula el KPI de Protocolos

Se replica la lógica del propio dashboard, para que las cifras digan lo mismo en los dos lados:

```
Universo      = S + C + P + AP + AE + rv       (62.728)
En falta      = AP + P                         (163)
Base del KPI  = AP + P + AE + C                (44.878 = universo − S − rv)
KPI           = En falta / Base = % pendiente  (0,4% — menos es mejor)
```

Dos detalles que hay que respetar, o las cifras se van:

- **`AP` ya viene neto de `rv`** (protocolos en revisión del cliente), así que **no** se le
  vuelve a restar. Pero el `rv` **sí suma al universo**: si no se devuelve, el universo queda
  corto — eran los 200 de diferencia contra el módulo.
- **`S` (remanente) queda fuera del KPI** pero dentro del universo. Reportar la base del KPI
  como si fuera el universo fue el error original.
- En un nodo con hijos **no** se suman sus propios estados, solo los de los hijos (así lo hace
  el `sumCh()` del dashboard); el `rv` del padre sí se acumula.

**Encuadre de los módulos:** título a y=45, alto 22 px, centrado en el mismo eje, cabecera de
48 px y controles terminando en x=1478 — idénticos en los dos módulos.

**Rendimiento:** portada 0,37 s · Protocolos 0,70 s · Cierre QAQC 0,60 s · volver a un módulo
ya montado 0,19 s · 18 MB de memoria JS.
