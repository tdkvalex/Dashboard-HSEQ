# Control de No Conformidades — módulo de la Suite QAQC

Dashboard de los hallazgos de calidad de los tres proyectos **y de Oficina Central**.

**Corte 03-08-2026 · 508 registros** desde 10-02-2023, de **dos fuentes**.

| Frente | Levantados | Cerrados | Abiertos | % cierre | Mediana de cierre |
|---|---|---|---|---|---|
| Desaladora | 95 | 92 | 3 | 96,8% | 19 días |
| Talabre | 241 | 226 | 15 | 93,8% | 23 días |
| **Arqueros** | **164** | **128** | **36** | **78,0%** | **131 días** |
| Oficina Central | 8 | 8 | 0 | 100% | 77 días |
| **Total** | **508** | **454** | **54** | **89,4%** | **30 días** |

**Semáforo:** Arqueros y Talabre *Crítico* y *Atención* · el resto *Al día*.

### Dos fuentes, y por qué

El registro principal (`Data_NCR*.xlsx`, hoja «Observaciones») trae, **para Arqueros, solo lo
que emite Besalco**: internas del contrato y contra subcontratos. Las que **el cliente MASA
levanta contra Besalco** viven en su propia planilla de control de recepción y respuesta
(`Data_NCR*externas.xlsx`, hoja «Disposición NC-Externas») y se cargan con `--externas`.

Sin esa segunda fuente Arqueros se veía con **1 hallazgo abierto y 98,9% de cierre**. Con ella:
**36 abiertos y 78,0%**, y el frente pasa de *Al día* a **Crítico**. Son 70 registros —35
cerrados y 35 abiertos— recibidos entre 30-04-2025 y 09-06-2026.

| Arqueros por vía de emisión | Levantados | Abiertos | % cierre |
|---|---|---|---|
| Interna (Besalco) | 53 | 1 | 98,1% |
| **Externa · Cliente** | **71** | **35** | **50,7%** |
| Externa · Subcontrato | 40 | 0 | 100% |

La autodetección de Arqueros cae de 56,4% a **32,3%**: no cambió lo que Besalco detecta,
apareció lo que le levantan.

### Tres vías de emisión, nunca dos

**La externa del cliente y la del subcontrato no se suman en ningún lugar del panel ni de la
PPT.** No son lo mismo: la del cliente **compromete el contrato**, la del subcontrato la emite
Besalco y la absorbe internamente. Juntarlas en un solo «externas» escondía justo la delicada
—en Arqueros, «35 externas abiertas» eran **35 del cliente y 0 de subcontrato**—.

Van separadas en el KPI de levantados, en el de abiertas, en «lo que sigue abierto por quién lo
levantó», en el gráfico de origen, en los cortes de la semana, en el semáforo corporativo (tres
columnas de abiertos y la fila `internas / cliente / subcontrato`), en las dos tablas de detalle
y en las láminas 3 y 5 de la PPT. El color es el mismo en todas partes: **azul** lo interno,
**ámbar** lo del cliente, **naranja** lo del subcontrato. El chip del cliente lleva borde lleno;
el del subcontrato, punteado.

### «Estado de cada vía» — qué pasó con lo que entró por cada una

Bajo el gráfico de origen, una tabla abre cada vía en **levantadas · cerradas · % cierre ·
abiertas · atrasadas · más de 180 días · antigüedad mediana**. El gráfico dice por dónde entra
el hallazgo; la tabla, qué pasó después.

**La columna «Atrasadas» dice `n/c`, no `0`.** Con este archivo el atraso no es calculable
(ninguna abierta trae fecha comprometida, ver más abajo) y escribir 0 se leería como que no hay
atraso. En su lugar la severidad de lo abierto la dan las dos columnas siguientes, que sí se
pueden calcular. Si algún día el Excel trae la fecha comprometida, la columna muestra el número
real sin tocar nada: depende de `control.atrasoCalculable`.

Al corte, lo que muestra: **las 49 abiertas del cliente cierran al 83%** contra 96,9% las
internas y 100% las de subcontrato —de las que no queda ninguna abierta—, y las 14 que superan
los 180 días son todas del cliente. En Arqueros la brecha es la mayor: 50,7% contra 98,1%.

Las cifras salen de `resumir()` en `no_conformidades.py`, que expone `cliente`, `subcontrato`,
`abiertasCliente`, `abiertasSubcontrato` y `abiertasInternas` además de los viejos
`internas`/`externas`. Si se agrega una vista nueva, usar esas y no `externas`.

**Lo que esa planilla no trae:** disciplina, responsable ni costo. Los 70 registros entran como
«Sin especialidad» y no suman al costo declarado. El panel lo declara en «Control de calidad
del dato» en vez de dejar que el hueco se lea como dato perdido.

**Trampa del archivo:** la «Fecha MASA» de las revisiones abiertas trae la fecha de hoy —es una
fórmula que cuenta días de espera—, así que solo se lee como fecha de cierre cuando el status
de esa misma revisión es «Aprobado».

### Levantados en la última semana (28-07 → 03-08): 1

| Fecha | Frente | N° | Tipo | Origen | Disciplina | Responsable |
|---|---|---|---|---|---|---|
| 28-07 | Talabre | 241 | No Conformidad | Externa · Cliente | CALIDAD | Daniel Ramirez B. |

**Talabre cerró 13 hallazgos en la semana**, entre ellos 6 internas de Adquisiciones que
superaban los 100 días —una llevaba 385—. Sus abiertas bajan de 26 a 15 y pasa de *Crítico* a
*Atención*. De las 5 internas que quedan abiertas en todo el módulo, **3 son de Expeditoría en
Talabre** —392, 328 y 224 días, todas «Iniciado» y del mismo responsable—.

---

## Actualización

```bash
python3 no_conformidades.py --data /ruta/Data_NCR.xlsx \
                            --externas /ruta/Data_NCR*externas.xlsx
node gen_ppt.js                       # SIEMPRE después del script de Python
cd ../suite_qaqc && node armar_suite.js
```

**`--externas` no es opcional en la práctica**: sin él Arqueros queda sin las NC que le levanta
el cliente y el script avisa. Ver «Dos fuentes» más arriba.

`no_conformidades.py` acepta `--hoy AAAA-MM-DD` para fijar la fecha de referencia. Escribe
`datos_nc.json`, inyecta los datos en `index.html` y avisa de todo lo que no pueda clasificar.

`gen_ppt.js` va **después** porque regenera la PPT desde `datos_nc.json`, la **embebe** en
`index.html` (botón «Descargar Informe») y embebe los logos. Correrlo antes deja el informe
descargable desfasado respecto de las pestañas. Necesita `npm install` una sola vez.

Para la suite no hay que copiar nada: el generador lee `modulo_nc/index.html` directamente.

---

## La PPT ejecutiva — 16 láminas

Mismo formato y metodología que la del módulo de Cierre QAQC
(`panel_control_TOP_P1/gen_ppt.js`), para que el mazo se lea igual venga del módulo que venga.

| Láminas | Contenido |
|---|---|
| 1 | Portada del consolidado |
| 2 | Resumen ejecutivo — 4 tarjetas + lectura del período |
| 3 | Semáforo por frente (los 3 proyectos + Oficina Central + total de obra) |
| 4 | Levantados en la última semana — ritmo de 8 semanas (5 al detalle + 3 acumuladas) y detalle uno a uno |
| 5 | Origen del hallazgo — interna, cliente y subcontrato · estado de cada vía · autodetección por frente |
| 6 | Antigüedad de lo abierto + por qué el atraso no es calculable + velocidad de cierre |
| 7 | Foco de gestión |
| 8-16 | Portada + 2 láminas por proyecto (estado · pendiente y novedades) |

Reglas del formato, iguales a las del otro módulo:

- **Logo** en todas: grande (`x9.14 y0.46`) en portadas y láminas azules, menor (`x9.97 y0.24`)
  en contenido. Se busca en esta carpeta y, si no está, en `panel_control_TOP_P1/`.
- **Etiqueta del proyecto** al pie de cada lámina de contenido; las portadas no la llevan.
- **Numeración correlativa automática** (`nSlide`), no escrita a mano.
- **Etiquetas de gráficos**: constante `ETIQ` — blanco, negrita, 11 pt. Los colores de segmento
  dan **≥4,5:1 con texto blanco** (el peor es `warnD`, 4,68:1). El cobre original daba 3,4:1 y
  el número quedaba invisible: para eso está `copperD`.
- **Valores 0** envueltos en `z()` para que no se dibujen.
- **Nunca se recorta en silencio**: si una semana trae más hallazgos de los que caben en la
  tabla, la lámina dice cuántos se listan y cuántos hay.

La PPT viaja **embebida en base64 dentro de `index.html`**, así el panel sigue siendo un solo
archivo: se envía por correo y quien lo reciba descarga el informe sin servidor. El archivo
descargado se llama `Panel_NC_DDMMAA.pptx` con la fecha del corte, para no pisar informes de
semanas distintas.

**LibreOffice de este contenedor no abre PPTX** (falta el filtro de Impress). Para verificar,
leer el XML con `python-pptx` / `unzip`, no intentar renderizar.

---

## Lo que hay que saber del dato

### El atraso NO es calculable con este archivo

La regla del proyecto es: **NC abierta cuya fecha de cierre comprometida ya venció**. Está
implementada en el script, pero **ninguna de las 19 NC abiertas tiene fecha**: la columna
«Fecha De Cierre» solo se llena **cuando la NC se cierra**, así que es la fecha de cierre real,
no un compromiso.

El panel **no reporta 0 atrasadas** —eso se leería como un buen resultado—, sino que lo declara
abiertamente en «Control de calidad del dato» y usa en su lugar la **antigüedad**: días desde
que se levantó cada hallazgo abierto. Al corte: mediana 55 días, y **1 lleva más de un año**.

> Para activar el indicador de atraso basta agregar al Excel una columna de **fecha
> comprometida de cierre**. El script ya la calcularía sin cambios.

### Otras definiciones

- **Interna** = la levanta Besalco (`Interna BSMT`) · **Externa · Cliente** = la levanta el
  mandante y compromete el contrato · **Externa · Subcontrato** = la emite Besalco contra un
  subcontratista y la absorbe internamente. **Las dos externas nunca se suman** (ver «Tres vías
  de emisión» más arriba). Al corte: 162 internas · 289 del cliente · 48 de subcontrato ·
  9 sin clasificar; de las 54 abiertas, **49 son del cliente**, 5 internas y ninguna de
  subcontrato.
- **Autodetección** = qué parte de los hallazgos clasificados los levanta Besalco. Global
  **37,8%**; por frente, Desaladora 59,6% · Arqueros 56,4% · **Talabre 22,0%**.
- **La semana** = lo levantado en los últimos 7 días (`creada > hoy − 7 días`). Cada pestaña
  muestra la suya: la corporativa los cuatro frentes y cada proyecto solo los propios.
- **El ritmo** cubre 8 semanas, pero solo las **5 más recientes van una a una**; las 3
  anteriores se suman en una fila, «5 semanas o más». Ocho filas ocupaban media pantalla sin
  aportar lectura. Esa fila **no es una semana**: va en cursiva, con una línea punteada al
  costado y con el número de semanas que suma en su rótulo, para que su barra no se compare
  con la de una sola. Las constantes son `SEM_VENTANA` y `SEM_DETALLE` en
  `no_conformidades.py`; el panel y la PPT se adaptan solas al cambiarlas.
- **Abierta** = todo estatus distinto de «Cerrado»: *Iniciado*, *Listo para revisión* y
  *No aceptado*. Se conserva el estatus original, porque «No aceptado» no se gestiona igual
  que «Iniciado».
- **Tiempo de cierre** = días entre la fecha de creación y la de cierre, solo sobre las
  cerradas. Es el indicador de qué tan rápido reacciona cada proyecto.
- **Oficina Central** genera NC pero no es una obra: aparece en el resumen corporativo y en el
  consolidado «Los 3 proyectos» se excluye, para que sea comparable con los otros módulos.
- **182 de 438 hallazgos no declaran costo**, así que las 10.783 UF son un piso, no el costo real.

---

## Lo que muestran los datos

**Arqueros concentra el pendiente:** 36 de las 54 NC abiertas, y **35 de esas 36 las levantó el
cliente** — ninguna es de subcontrato. Talabre sigue siendo el frente con más hallazgos del
mandante que propios (188 contra 53 internas), pero es el que más se movió en la semana:
cerró 13 y bajó de 26 a 15 abiertas.

**La deuda antigua que queda es de Expeditoría:** 3 de las 5 internas abiertas son de esa
especialidad en Talabre —392, 328 y 224 días, todas «Iniciado» y del mismo responsable—.
Adquisiciones, que arrastraba 6 sobre 100 días, las cerró todas.

**Arqueros cierra 6 veces más lento:** 139 días de mediana contra 23 de Talabre y 19 de
Desaladora, pese a tener solo 1 NC abierta. Cierra todo, pero tarde.

**Obras Civiles domina** con 168 hallazgos de 438 (38%) y 10 abiertos.

**El volumen se disparó en 2025** (293 hallazgos contra 18 en 2024), y 2026 ya lleva 124.

---

## Archivos

| Archivo | Qué es |
|---|---|
| `index.html` | El dashboard, un solo archivo portable con 4 pestañas y la PPT embebida |
| `no_conformidades.py` | **Punto de entrada.** Procesa el Excel y actualiza el dashboard |
| `gen_ppt.js` | Genera la PPT y la embebe en el dashboard. Va después del script de Python |
| `datos_nc.json` | Datos consolidados; la suite lee de aquí su KPI |
| `Panel_No_Conformidades.pptx` | La PPT del corte, también disponible como archivo suelto |
| `package.json` | La única dependencia: `pptxgenjs`. `npm install` una vez |
| `besalco_logo*.png` | Logo corporativo. No versionados: se usan los de `panel_control_TOP_P1/` |

### Bloques marcados de `index.html`

Los scripts reemplazan solo su bloque; el resto del archivo se edita a mano.

| Marca | La escribe | Contiene |
|---|---|---|
| `NC:INICIO/FIN` | `no_conformidades.py` | `const NC` — todos los datos |
| `LOGO:INICIO/FIN` | `gen_ppt.js` | `const LOGOS` — logos en base64 |
| `PPT:INICIO/FIN` | `gen_ppt.js` | `const PPTX` — la PPT en base64 |

**Ancho del contenido:** `.wrap` usa `max-width: clamp(1680px, 88vw, 2160px)`, **el mismo valor
que `panel_control_TOP_P1/index.html`**. Si se cambia en uno hay que cambiarlo en el otro, o
los dos módulos saltan de tamaño al cambiar de pestaña dentro de la suite.
