# Cada lunes — actualización de la Suite QAQC

Tres comandos. El resto es pedir bien los archivos.

```bash
python3 actualizar_semana.py --entrada ~/Descargas/corte_10-08
python3 verificar_suite.py
python3 auditoria_datos.py ~/Descargas/corte_10-08
```

Se envía `suite_qaqc/Suite_QAQC.html`. Es un solo archivo: va por correo tal cual.

---

## 1 · Qué pedir, a quién

Todo a una misma carpeta. **Los nombres de archivo dan lo mismo**: el script
reconoce cada uno por su contenido, no por cómo se llame.

| Proyecto | A quién | Qué se pide | Cómo se reconoce |
|---|---|---|---|
| **P2416 · Desaladora** | Juan Vergara | `REPORTE_GERENCIAL_*.xlsx` **y** `Listado_Puntos_Punch_Consolidado_*.xlsx` | hojas «REPORTE GERENCIAL» y «LISTADO PUNCH ITEMS» |
| **P2407 · Talabre** | Mauricio Rocha | `STATUS_SUBSISTEMAS_TALABRE.xlsx` **y** `Detalle_de_TerminacionesBesalco.xlsx` | hoja «STATUS» · hoja «DT» con más de 20 columnas |
| **P2342 · Arqueros** | José Muñoz | `Estatus_Resumen_General_QAQC.xlsx` | hojas «BD Caminatas-CTOP» + «BD Detalles Terminación» |
| **No Conformidades** | — | `Data_NCR*.xlsx` **y** la planilla de NC externas del cliente MASA | hoja «Observaciones» · hoja «Disposición NC-Externas» |
| **Protocolos** | el equipo del dashboard | el `.html` del dashboard actualizado | contiene `PROJECTS` y `NODE_HISTORY` |

**Desaladora, Talabre y No Conformidades necesitan SUS DOS archivos.** Con uno
solo el script no procesa ese módulo y lo dice.

**Lo que no llega no se toca.** Si Desaladora no mandó nada, su pestaña conserva
el corte anterior y el resto se actualiza igual. No hay que hacer nada especial.

---

## 2 · Las tres trampas que ya costaron un corte

**El `Programa_de_Caminatas.xlsm` de Talabre no es fuente.** Trae una hoja «DT»
que parece el registro completo pero es una foto vieja. El script compara la
emisión más nueva de cada candidato contra la que ya está cargada y **se detiene**
si el archivo viene más viejo. Si aparece ese aviso, pedir el archivo correcto.

**La planilla de NC externas no es opcional.** Sin ella Arqueros pierde las 36 NC
que le levanta MASA y aparece con 98,9% de cierre en vez de 78%. También detiene
la corrida.

**El dashboard de Protocolos es el único que no se puede regenerar aquí.** Llega
armado de otro equipo y vive solo en `suite_qaqc/modulos/protocolos.html`, que no
se versiona (pesa 3 MB y no es nuestro). **Guarda una copia fuera del repositorio
cada vez que llegue uno nuevo.** Si se pierde, la suite no se puede armar hasta
que ese equipo lo reenvíe — y desde este corte el generador se detiene en vez de
publicar Protocolos en 0 %, que se lee como «sin avance» y no como «sin dato».

**Si se reprocesa un corte pasado, va `--corte`.** La ventana de «levantados en la
última semana» y los días de atraso se cuentan contra esa fecha; sin el parámetro
se usa el día en que se corre.

```bash
python3 actualizar_semana.py --entrada ~/Descargas/corte_03-08 --corte 2026-08-03
```

Cuando el script se detiene por un aviso no procesó nada: se corrige y se repite.
Si el aviso está entendido y se quiere seguir igual, se agrega `--igual`.

---

## 3 · Cómo circula un archivo, de la carpeta a la suite

Ocho documentos entran; uno sale. Nadie escribe una cifra a mano en ninguna parte.

```
CARPETA DEL CORTE                 SE LEE ASÍ                 PRODUCE            SE INYECTA EN
─────────────────────────────────────────────────────────────────────────────────────────────
REPORTE_GERENCIAL_*.xlsx    ─┐  hojas «REPORTE GERENCIAL»
Listado_Puntos_Punch_*.xlsx ─┴─▶ y «LISTADO PUNCH ITEMS»  ──▶ datos_desaladora  ─┐
                                 columnas por POSICIÓN                           │
STATUS_SUBSISTEMAS_*.xlsx   ─┐  hojas «STATUS» y «DT»                            │
Detalle_de_Terminaciones*   ─┴─▶ columnas por NOMBRE       ──▶ datos_talabre    ─┼─▶ panel_control_TOP_P1/
                                                                                 │      index.html
Estatus_Resumen_General_*   ───▶ «BD Caminatas-CTOP» +                            │   (3 bloques marcados)
                                 «BD Detalles Terminación»──▶ estatus_datos     ─┘        │
                                 columnas por POSICIÓN                                    ▼
                                                                                   node gen_ppt.js
Data_NCR*.xlsx              ─┐  «Observaciones»                                    embebe la PPT
NC externas del cliente     ─┴─▶ «Disposición NC-Externas» ──▶ datos_nc.json ──▶ modulo_nc/index.html
                                                                                          │
dashboard_protocolos.html   ───▶ se copia tal cual ─────────▶ modulos/protocolos.html      │
                                                                                          ▼
                                                            node armar_suite.js ──▶ Suite_QAQC.html
```

**El reconocimiento es por contenido, no por nombre.** `actualizar_semana.py` mira los
nombres de hoja de cada archivo de la carpeta y deduce su rol; solo cuando dos archivos
podrían ser el mismo rol —las dos hojas «DT» de Talabre— baja a mirar encabezados y ancho.
Por eso da igual cómo se llamen los archivos, y por eso un archivo que no encaja se informa
en vez de procesarse mal.

**El orden importa en dos puntos, y el script lo respeta solo:**

1. `gen_ppt.js` va **después** de los scripts de Python de su módulo. Regenera la PPT desde
   los JSON y la embebe en el `index.html`; correrlo antes deja el informe descargable
   desfasado respecto de las pestañas.
2. `armar_suite.js` va **al final de todo**. Empaqueta los `index.html` ya terminados.

**Nada se pisa entre módulos.** Cada script reemplaza solo su bloque marcado del
`index.html`; el resto del archivo no se toca. Los tres proyectos de Cierre QAQC escriben
bloques distintos del mismo archivo, así que corren **en serie**, no en paralelo.

**Lo que no llega, no se toca.** Cada módulo conserva su corte anterior, y por eso los tres
proyectos suelen quedar a fechas distintas: el panel lo declara con el aviso «cortes mixtos».

**Cuánto demora.** Un corte completo son unos **6 segundos** —el reconocimiento de la
carpeta, menos de 1—. Si alguna vez se dispara, casi siempre es un Excel que llegó con
`styles.xml` inflado: `qaqc_excel.py` los descarta al vuelo y lo explica ahí mismo.

---

## 4 · Antes de enviar

`verificar_suite.py` revisa solo lo que se olvida a mano y **sale con error si
algo falla**:

- las dos PPT: numeración correlativa, logo en todas, ningún texto montado sobre
  el pie, etiquetas de gráfico legibles (≥4,5:1 contra el número blanco);
- la suite en 1366, 1920 y 2560 px: los tres módulos montan, se recorren todas
  sus pestañas, no hay desbordes horizontales ni errores de consola;
- que el botón «Descargar Informe» entregue un `.pptx` de verdad;
- que los detalles sumen lo mismo en los proyectos y en el consolidado.

Con `--rapido` corre solo a 1920 px, para una pasada corta.

`auditoria_datos.py` revisa lo otro: que **la misma cifra diga lo mismo en todas
las capas** —Excel, JSON de cada módulo, portada de la suite y texto de las PPT—,
que las partes sumen el total, que la semana informada vaya de lunes a domingo y
que el ritmo no tenga huecos. Pasándole la carpeta del corte agrega el cruce
contra el Excel de origen, que es el único capaz de pillar un error de lectura.
También sale con error si algo falla.

Lo que **sí** hay que mirar a ojo:

1. **Los AVISOS de cada script.** Ahí aparecen los valores que no supo clasificar
   y los cruces que dejaron de cuadrar. Si sale uno nuevo, investigar antes de
   enviar: casi siempre es que el archivo de origen cambió.
2. **El corte de la franja superior.** Si los módulos quedaron a fechas distintas
   aparece el aviso «cortes mixtos» con el detalle al pasar el cursor.
3. **Una pasada por las pestañas**, sobre todo si algún proyecto cambió de formato.

---

## 5 · Si algo se rompe

| Síntoma | Causa probable |
|---|---|
| «no vino — se conserva el corte anterior» en un archivo que sí mandaste | el archivo no tiene las hojas esperadas: ábrelo y compara con la tabla de arriba |
| Un script avisa de columnas o valores no reconocidos | el proyecto cambió el formato del Excel. `talabre.py` resuelve columnas por nombre y aguanta; los demás usan posiciones fijas documentadas en `panel_control_TOP_P1/CLAUDE.md` §4 |
| «Falta … — la suite NO se escribió» | falta el HTML de ese módulo (casi siempre `modulos/protocolos.html`). Repónlo; la suite anterior sigue intacta. Armarla incompleta a propósito: `node armar_suite.js --sin-modulo` |
| Un módulo se quedó pegado en un corte viejo | hay una copia en `suite_qaqc/modulos/` que le gana al archivo en vivo. Solo `protocolos.html` debe estar ahí: borra `cierre_qaqc.html` o `no_conformidades.html` si aparecieron |
| La PPT no descarga desde el panel | se corrió `gen_ppt.js` antes que el script de Python. Repetir en orden — `actualizar_semana.py` lo hace solo |

El detalle de cada módulo está en su propio README:
[`panel_control_TOP_P1/`](panel_control_TOP_P1/README.md) ·
[`modulo_nc/`](modulo_nc/README.md) · [`suite_qaqc/`](suite_qaqc/README.md).
Las reglas de homologación entre proyectos —validadas con los jefes de calidad y
que **no se cambian sin avisar**— están en
[`panel_control_TOP_P1/CLAUDE.md`](panel_control_TOP_P1/CLAUDE.md) §3.
