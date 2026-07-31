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

Los tres cubren **los mismos tres proyectos**, con distinto nombre en cada módulo:

| Proyecto | Cliente | En Protocolos | En Cierre QAQC |
|---|---|---|---|
| Desaladora | Antofagasta Minerals | P2416 | Desaladora |
| Talabre | Codelco | P2407 | Talabre |
| Arqueros | MASA | P2342 | Arqueros |

---

## Actualización

```bash
# 1 · actualizar el módulo Cierre QAQC (ver panel_control_TOP_P1/CLAUDE.md)
cd ../panel_control_TOP_P1
python3 desaladora.py --reporte <REPORTE_GERENCIAL_*> --punch <Listado_Puntos_Punch_*>
python3 talabre.py    --status  <TalabreSTATUS_PEC>   --dt    <TalabreCuadro_DT>
python3 actualizar.py <Estatus_Resumen_General_QAQC.xlsx>   # Arqueros
node gen_ppt.js

# 2 · actualizar el módulo No Conformidades (ver modulo_nc/README.md)
cd ../modulo_nc
python3 no_conformidades.py --data <Data_NCR.xlsx>
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

**La portada no lee los módulos en vivo**: `armar_suite.js` extrae los KPI en el momento de
armar y los deja en un bloque `KPIS`. Así la portada no depende de que los iframes estén
montados ni de comunicación entre documentos.

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
| `Suite_QAQC.html` | **El entregable.** Un solo archivo, ~8,1 MB, sin dependencias externas |
| `armar_suite.js` | Extrae los KPI y ensambla la suite. **Punto de entrada** |
| `plantilla_suite.html` | Shell ejecutivo: portada, navegación y montaje de módulos |
| `modulos/protocolos.html` | Dashboard de Protocolos, tal cual |
| `modulos/cierre_qaqc.html` | Opcional: si no está, se usa `panel_control_TOP_P1/index.html` |
| `modulos/no_conformidades.html` | Opcional: si no está, se usa `modulo_nc/index.html` |
| `kpis_suite.json` | KPI extraídos, para consultar sin abrir la suite |
| `besalco_logo*.png` | Logo corporativo (versión blanca para fondo oscuro) |

---

## Consolidado al corte

| Módulo | Indicador | Valor |
|---|---|---|
| Protocolos | % pendiente | **0,3%** — 120 en falta de 44.234 |
| Cierre QAQC | % detalles cerrados | **59,1%** — 1.286 atrasados |
| No Conformidades | % hallazgos cerrados | **93%** — 30 abiertos de 428 |

**Lo que muestra el cruce:** la documentación va muy por delante del cierre físico. Protocolos
está prácticamente al día mientras el cierre de detalles P1 arrastra 596 ítems atrasados. El
cuello de botella no está en el papel: está en terreno.

---

## Validación al corte actual

Cada cifra de la portada se contrastó contra lo que muestra el propio módulo:

| Control | Suite | Módulo |
|---|---|---|
| Protocolos · universo | 61.264 | 61.264 |
| Protocolos · KPI corporativo | 0,3% | 0,3% |
| Protocolos · estado (P) | 11 | 11 |
| Protocolos · estado (AP) | 109 | 109 |
| Protocolos · estado (AE) | 5.420 | 5.420 |
| Protocolos · estado (C) | 38.694 | 38.694 |
| Cierre QAQC · P1 | 65,1% | 65,1% |
| Cierre QAQC · detalles | 6.300 · 1.286 atrasados | 6.300 · 1.286 |

### Cómo se calcula el KPI de Protocolos

Se replica la lógica del propio dashboard, para que las cifras digan lo mismo en los dos lados:

```
Universo      = S + C + P + AP + AE + rv       (61.264)
En falta      = AP + P                         (120)
Base del KPI  = AP + P + AE + C                (44.234 = universo − S − rv)
KPI           = En falta / Base = % pendiente  (0,3% — menos es mejor)
```

Dos detalles que hay que respetar, o las cifras se van:

- **`AP` ya viene neto de `rv`** (protocolos en revisión del cliente), así que **no** se le
  vuelve a restar. Pero el `rv` **sí suma al universo**: si no se devuelve, el universo queda
  corto — eran los 148 de diferencia contra el módulo.
- **`S` (remanente) queda fuera del KPI** pero dentro del universo. Reportar la base del KPI
  como si fuera el universo fue el error original.
- En un nodo con hijos **no** se suman sus propios estados, solo los de los hijos (así lo hace
  el `sumCh()` del dashboard); el `rv` del padre sí se acumula.

**Encuadre de los módulos:** título a y=45, alto 22 px, centrado en el mismo eje, cabecera de
48 px y controles terminando en x=1478 — idénticos en los dos módulos.

**Rendimiento:** portada 0,37 s · Protocolos 0,70 s · Cierre QAQC 0,60 s · volver a un módulo
ya montado 0,19 s · 18 MB de memoria JS.
