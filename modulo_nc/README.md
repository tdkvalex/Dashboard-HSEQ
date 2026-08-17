# Control de No Conformidades — módulo de la Suite QAQC

Dashboard de los hallazgos de calidad de los tres proyectos **y de Oficina Central**.

**Corte 03-08-2026 · 493 registros** desde 10-02-2023, de **dos fuentes**.

| Frente | Levantados | Cerrados | Abiertos | Fuera de plazo | % cierre | Mediana de cierre |
|---|---|---|---|---|---|---|
| Desaladora | 82 | 80 | 2 | 1 | 97,6% | 19 días |
| Talabre | 240 | 225 | 15 | 13 | 93,8% | 23 días |
| **Arqueros** | **163** | **127** | **36** | **36** | **77,9%** | **129 días** |
| Oficina Central | 8 | 8 | 0 | 0 | 100% | 77 días |
| **Total** | **493** | **440** | **53** | **50** | **89,2%** | **30 días** |

### Las opciones de mejora quedan fuera

**No son hallazgos que haya que corregir, son propuestas.** Sumarlas infla el universo y
ensucia el % de cierre con algo que nadie está obligado a cerrar.

**La regla rige para los cuatro frentes y para las dos fuentes**, no solo para el registro
principal: se filtran al leer, antes de construir nada, así que ninguna vista, ni la PPT, ni la
portada de la suite las cuenta. En el módulo no queda ningún tipo «Opción de Mejora».

El Excel trae 507 registros y el módulo reporta **493**: quedan fuera **14 opciones
de mejora** —13 de Desaladora · 1 de Talabre; Arqueros y Oficina Central no tenían—, de las cuales solo 1 seguía abierta.

No se borran en silencio: el panel las declara en «Control de calidad del dato» con el
contraste 507 → 493 y el desglose por frente, y el script las informa en sus avisos.
La constante es `TIPOS_EXCLUIDOS` en `no_conformidades.py`; para volver a incluirlas basta
vaciarla, y para excluir otro tipo basta agregarlo.

**Semáforo:** Arqueros y Talabre *Crítico* y *Atención* · el resto *Al día*.

### Dos fuentes, y por qué

El registro principal (`Data_NCR*.xlsx`, hoja «Observaciones») trae, **para Arqueros, solo lo
que emite Besalco**: internas del contrato y contra subcontratos. Las que **el cliente MASA
levanta contra Besalco** viven en su propia planilla de control de recepción y respuesta
(`Data_NCR*externas.xlsx`, hoja «Disposición NC-Externas») y se cargan con `--externas`.

Sin esa segunda fuente Arqueros se veía con **1 hallazgo abierto y 98,9% de cierre**. Con ella:
**36 abiertos y 77,9%**, y el frente pasa de *Al día* a **Crítico**. Son 70 registros —35
cerrados y 35 abiertos— recibidos entre 30-04-2025 y 09-06-2026.

| Arqueros por vía de emisión | Levantados | Abiertos | % cierre |
|---|---|---|---|
| Interna (Besalco) | 53 | 1 | 98,1% |
| **Externa · Cliente** | **70** | **35** | **50,0%** |
| Externa · Subcontrato | 40 | 0 | 100,0% |

La autodetección de Arqueros cae de 56,4% a **32,5%**: no cambió lo que Besalco detecta,
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
abiertas · atrasadas · promedio en respuesta**. El gráfico dice por dónde entra el
hallazgo; la tabla, qué pasó después.

- **Atrasada** = abierta que pasó el plazo de respuesta (ver más abajo).
- **Promedio en respuesta** = días entre que se emite y se cierra, **solo sobre las ya
  cerradas**. Es lo que tardó de verdad cada vía contra los 10 días de plazo. Las que siguen
  abiertas no entran, así que la cifra es optimista: la antigüedad de lo abierto va en su propio
  bloque. Es `promedioCierre` en `resumir()`; la mediana se guarda aparte (`medianaCierre`)
  porque un hallazgo de 400 días mueve el promedio y no la mediana.

La tabla usa `.tbl.compacta` (6 px de aire por celda en vez de 10): comparte fila con «Estado de
cierre», así que con el padding normal se salía de su media columna. Entra completa desde
~1.900 px de ancho; por debajo se desplaza dentro de su caja. Si se quiere que entre siempre,
sacarla del `.grid2` y dejar el bloque a ancho completo.

Al corte: **las 48 abiertas del cliente cierran al 82,7%** contra 96,9% las internas y 100% las de
subcontrato —de esas no queda ninguna abierta—. De las 50 fuera de plazo, **45 son del cliente**.
Y el promedio de respuesta va en **69 días contra los 10 de plazo**: 50 las del cliente,
64 las internas y 179 las de subcontrato. En Arqueros la brecha es la mayor: 50,0% de cierre
contra 98,1%, sus 36 abiertas están todas fuera de plazo y responde en 127 días promedio.

Las cifras salen de `resumir()` en `no_conformidades.py`, que expone `cliente`, `subcontrato`,
`abiertasCliente`, `abiertasSubcontrato` y `abiertasInternas` además de los viejos
`internas`/`externas`. Si se agrega una vista nueva, usar esas y no `externas`.

**Lo que esa planilla no trae:** disciplina, responsable ni costo. Los 70 registros entran como
«Sin especialidad» y no suman al costo declarado. El panel lo declara en «Control de calidad
del dato» en vez de dejar que el hueco se lea como dato perdido.

**Trampa del archivo:** la «Fecha MASA» de las revisiones abiertas trae la fecha de hoy —es una
fórmula que cuenta días de espera—, así que solo se lee como fecha de cierre cuando el status
de esa misma revisión es «Aprobado».

### Levantados en la última semana (27-07 → 02-08): 3

**La semana es la calendario, de lunes a domingo, que ya cerró al corte** — no los últimos 7
días corridos. El informe se arma el lunes y habla de la semana que terminó: con corte el lunes
03-08 la ventana es lunes 27-07 → domingo 02-08. Con la ventana móvil anterior (28-07 → 03-08)
se colaban los hallazgos del lunes del propio informe y se perdían los del lunes anterior: la
semana quedaba partida en dos y no coincidía con la que se revisa en la reunión.

| Fecha | Frente | N° | Tipo | Origen | Disciplina | Responsable |
|---|---|---|---|---|---|---|
| 27-07 | Desaladora | 95 | No Conformidad | Externa · Cliente | MECÁNICA | Wilson Jara |
| 27-07 | Talabre | 240 | Producto No Conforme | Externa · Cliente | OO.CC | Sebastián Ramos |
| 28-07 | Talabre | 241 | No Conformidad | Externa · Cliente | CALIDAD | Daniel Ramirez B. |

**Talabre cerró 13 hallazgos en la semana**, entre ellos 6 internas de Adquisiciones que
superaban los 100 días —una llevaba 385—. Sus abiertas bajan de 26 a 15 y pasa de *Crítico* a
*Atención*. De las 5 internas que quedan abiertas en todo el módulo, **3 son de Expeditoría en
Talabre** —392, 328 y 224 días, todas «Iniciado» y del mismo responsable—.

---

## Actualización

Desde la raíz del repositorio, `python3 actualizar_semana.py --entrada <carpeta>` corre este
módulo y los otros dos en orden. Ver [`CADA_LUNES.md`](../CADA_LUNES.md). Para reprocesar solo
No Conformidades:

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

## La PPT ejecutiva — 17 láminas

Mismo formato y metodología que la del módulo de Cierre QAQC
(`panel_control_TOP_P1/gen_ppt.js`), para que el mazo se lea igual venga del módulo que venga.

| Láminas | Contenido |
|---|---|
| 1 | Portada del consolidado |
| 2 | Resumen ejecutivo — 4 tarjetas + lectura del período |
| 3 | Semáforo por frente (los 3 proyectos + Oficina Central + total de obra) |
| 4 | Levantados en la última semana — ritmo de 8 semanas (5 al detalle + 3 acumuladas) y detalle uno a uno |
| 5 | Origen del hallazgo — interna, cliente y subcontrato · estado de cada vía · autodetección por frente |
| 6 | Antigüedad de lo abierto + fuera de plazo de respuesta + velocidad de cierre |
| 7 | **Costo mensual por disciplina** — matriz disciplina × mes, cobertura del dato y los tres estados |
| 8 | Foco de gestión |
| 9-17 | Portada + 2 láminas por proyecto (estado · pendiente y novedades) |

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

### El atraso se mide contra el plazo de respuesta

**La regla del proyecto: hay 10 días para responder una NC desde que se emite; del día 11 en
adelante corre atraso.** Los días de atraso son la antigüedad menos el plazo. La constante es
`PLAZO_RESPUESTA` en `no_conformidades.py`: si el plazo cambia, se cambia ahí y el panel, la PPT
y la suite se recalculan solos.

**Se cuenta sobre la fecha de emisión, no sobre una fecha comprometida de cierre.** Ninguna de
las dos fuentes trae esa fecha: en el registro principal las tres columnas de fecha
—«Fecha De Cierre» (dos veces) y «Fecha Revisión Y Cierre»— **solo se llenan al cerrar**, y las
19 abiertas las tienen todas vacías; la planilla del cliente tampoco declara plazo. La fecha de
emisión, en cambio, está en las dos, así que el atraso **sí es calculable**.

> Una versión anterior del módulo daba el atraso por no calculable y lo reemplazaba por la
> antigüedad. Era un error de criterio, no de dato: el plazo se cuenta desde la emisión.

Al corte: **50 de las 53 abiertas están fuera de plazo**. Solo 3 siguen dentro —dos de
7 días y una de 6—. Por frente: Arqueros 36 de 36, Talabre 13 de 15, Desaladora 2 de 3. Por vía:
46 de las 49 del cliente y las 5 internas. El atraso va de 3 a **385 días**.

La antigüedad se sigue reportando y no es redundante: dice **qué tan grave** es cada atraso.
Pasarse una semana del plazo no es lo mismo que llevar un año abierto, y de las 51 atrasadas,
17 superan los 180 días y 3 el año.

### Otras definiciones

- **Interna** = la levanta Besalco (`Interna BSMT`) · **Externa · Cliente** = la levanta el
  mandante y compromete el contrato · **Externa · Subcontrato** = la emite Besalco contra un
  subcontratista y la absorbe internamente. **Las dos externas nunca se suman** (ver «Tres vías
  de emisión» más arriba). Al corte: 162 internas · 289 del cliente · 48 de subcontrato ·
  de las 53 abiertas, **48 son del cliente**, 5 internas y ninguna de
  subcontrato.
- **Autodetección** = qué parte de los hallazgos clasificados los levanta Besalco. Global
  **37,8%**; por frente, Desaladora 59,6% · Arqueros 56,4% · **Talabre 22,0%**.
- **La semana** = la semana calendario **de lunes a domingo que ya cerró** al corte, no los
  últimos 7 días corridos (`semana_cerrada()`). Cada pestaña muestra la suya: la corporativa
  los cuatro frentes y cada proyecto solo los propios. Las ventanas de contexto —30 y 90 días—
  también se cuentan desde ese domingo, no desde el corte, o incluirían días que la semana
  informada deja fuera.
- **El ritmo** cubre 8 semanas lunes-domingo contadas hacia atrás desde ese mismo domingo, sin
  huecos ni solapes, pero solo las **5 más recientes van una a una**; las 3
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
### El costo: tres estados, no dos

La columna «Costo De La No Conformidad» admite tres situaciones que **no significan lo mismo**
y que el módulo publica por separado (`control.costoConMonto`, `costoEnCero`, `sinCosto`):

| Estado | Qué dice | Al corte 17-08-2026 |
|---|---|---|
| Declara un monto | costó tanto | **108** |
| Declara **0** | afirma que no costó | **143** |
| Casilla vacía | **no se sabe** | **247** |

Juntar los dos últimos en un solo «sin costo» haría leer como gratis lo que solo está sin
medir. Por eso **las 10.613 UF son un piso del costo real, nunca el costo real**, y el panel
y la PPT lo dicen en la misma lámina donde va la cifra, no en una nota al pie.

### Costo mensual por disciplina

**El monto se imputa al mes en que se LEVANTÓ la NC**, no al mes en que se cerró: la fuente no
registra una fecha propia del costo, así que esa es la única imputación defendible — y además
hace que la serie se lea contra «hallazgos por mes», que usa el mismo criterio. Va en
`costoMes` de cada frente, con el desglose por disciplina de cada mes.

Se agrega redondeando a **dos decimales**, no a uno: la columna de origen no trae más de dos,
así que a dos no se pierde nada y **la suma de los meses da exactamente el costo del frente**.
A un decimal cada mes se redondeaba por su cuenta y la columna terminaba sumando un décimo más
que el total, con la tabla del panel contradiciendo a la tarjeta del KPI. `auditoria_datos.py`
comprueba las dos cuadraturas —los meses contra el frente, y las disciplinas contra su mes— y
recuenta los tres estados desde el Excel.

Se presenta como **matriz disciplina × mes**, no como barras apiladas. La razón es concreta: un
solo mes de 4.382 UF —casi todo una NC de CALIDAD— dejaba a los otros veintiún meses en una raya
de un píxel, y la lectura que se pide, «cuánto puso cada disciplina en cada mes», había que
adivinarla por color. En la matriz el número está escrito y el color solo ordena la magnitud, en
**tramos fijos** (<50 · 50–199 · 200–499 · 500–1.499 · ≥1.500 UF) y no en cuantiles, para que un
mes no cambie de tono porque cambió otro. Además ocupa una fila por disciplina —doce— en vez de
una por mes —veintidós—.

La rampa arranca en el segundo escalón de `--seq-*`: con el primero, el tramo más bajo quedaba
del mismo tono que una celda vacía y no se distinguía «cobró poco» de «no cobró». La tinta de
cada celda se elige calculando el contraste contra las dos disponibles y quedándose con la mejor
—no con un umbral de luminancia fijo, que dejaba el tramo más brillante en 2,2:1—; así el panel
da ≥4,9:1 en modo oscuro y en modo claro, donde la rampa va al revés.

La última fila, **«NC con monto declarado»**, no es decorativa: un mes sin celdas es «sin
declarar», nunca «sin costo». Al corte hay 4 meses seguidos así (may-26 a ago-26): 45 hallazgos
levantados y ninguno con monto.

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
