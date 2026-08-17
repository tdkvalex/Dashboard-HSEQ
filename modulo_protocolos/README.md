# Protocolos — P2342 Arqueros (cliente MASA)

Carga las matrices de disciplina del proyecto y escribe el bloque `P2342` del
dashboard de Protocolos (`../suite_qaqc/modulos/protocolos.html`), su historial
por nodo y el punto semanal del KPI. **No toca Desaladora ni Talabre.**

```bash
cd modulo_protocolos
python3 protocolos_masa.py --matrices <carpeta con las matrices>
python3 protocolos_masa.py --matrices <carpeta> --seco             # no escribe
python3 protocolos_masa.py --matrices <carpeta> --verificar-cache  # ver abajo
```

`--corte AAAA-MM-DD` fija la fecha del punto semanal; por defecto es el lunes de
la semana en curso. Reprocesando un corte pasado hay que pasarlo, o el historial
queda con la fecha equivocada.

---

## 1 · Los archivos

Son **siete matrices, una por disciplina**. Se reconocen por contenido —los
nombres de archivo cambian—: el título de la hoja «Matriz de Protocolos» trae el
código del documento, `BSMT-P23/42-85-<SIGLA>-MTZ-01`, que identifica disciplina
**y** proyecto. Si el código dice otro proyecto, el archivo no se carga.

| Sigla | Matriz | Columnas que aporta al panel |
|---|---|---|
| `OOCC` | Obras Civiles | TOPO, OOCC |
| `ESTR` | Estructuras Metálicas | TOPO, ESTR |
| `MECA` | Equipos Mecánicos | TOPO, MECA, PCOM |
| `PIPN` | Piping | TOPO, PIPN |
| `ELEC` | Eléctricos | TOPO, EEII, PCOM |
| `INS` | Instrumentación | INST, EEII, PCOM |
| `COM` | Comunicación | EEII, PCOM, PREC |

**Se usa la hoja `KPI-BSMT`, nunca `KPI-MASA`.** La segunda es el alcance del
cliente; sólo Obras Civiles trae las dos. Si a una matriz le falta `KPI-BSMT`,
el script lo dice y **no** cae a `KPI-MASA`: hay que pedir la matriz correcta.

El proyecto se carga completo o no se carga: si falta una disciplina el script se
detiene. Media carga dejaría el KPI corporativo mintiendo sin que nada avise.

---

## 2 · Por qué se recalcula en vez de leer la hoja

`KPI-BSMT` no se lee: se **recalcula** aplicando sus propias fórmulas sobre la
matriz. Sus valores están guardados de la última vez que Excel recalculó, y eso
puede ser hace semanas.

Al corte 17-08-2026, **11 de 17 columnas** venían desfasadas. En casi todas la
diferencia era de 1 a 25 registros, pero en **Obras Civiles** las fórmulas
dinámicas (`UNIQUE`/`FILTER`) quedaron guardadas como `1` en cada columna —el
desborde colapsado— y la hoja declaraba **16 cerrados donde hay 4.748**.
Publicado tal cual, Arqueros aparecía desplomándose sin que hubiera pasado nada.

Donde la caché está sana el recálculo da **exactamente** su mismo número: eso es
lo que lo valida. `--verificar-cache` imprime la comparación columna por columna.

### Cómo cuenta el libro

La unidad **no es la misma para todos los estados**, y así viene del origen:

| Estado | Fórmula del libro | Unidad |
|---|---|---|
| `C`, `AP`, `AE` | `COUNTA(UNIQUE(FILTER(PROTOCOLO N°; ESTATUS="x")))` | protocolo distinto |
| `S`, `P`, `N` | `COUNTIF(ESTATUS;"x")` | celda elemento×protocolo |

El universo que arma la hoja suma las dos cosas. Es una inconsistencia del
formato de origen, no un error de lectura: se replica tal cual para que el panel
muestre lo mismo que MASA y Besalco revisan en la reunión. Cambiarlo por cuenta
propia haría que el panel y la planilla del cliente dejaran de coincidir.

Cada columna del KPI suma además la matriz de protocolos **y** la de precom del
mismo libro (`Matriz PreCom`, `Matriz de Precom`, `Matriz Precom-EEII`),
filtrando por la sigla que va en la fila 2 sobre cada columna de ESTATUS.

---

## 3 · Nodos sin matriz

**La línea BT/MT no llega desde el 29-06-2026.** Sus dos nodos —`TOPO-ELECBT-2342`
(98 protocolos) y `ELEC-BT-2342` (72)— conservan su último valor conocido y **no
reciben punto nuevo de historial**: inventarle continuidad a un dato que nadie
actualizó es peor que dejarlo quieto, y así la variación semanal no compara
contra una semana que no existió. El script lo declara en AVISOS cada vez que
corre, y `auditoria_datos.py` lo informa aparte.

Además, sus valores llevan 13 cortes idénticos (desde el 23-03-2026): esa parte
del alcance no se ha movido nunca. Conviene confirmar con el proyecto si sigue
vigente.

Dos nodos entraron nuevos el 17-08-2026, con alcance que las matrices declaran y
el panel no mostraba: `TOPO-ELEC-2342` (topografía dentro de la matriz Eléctrica,
8 protocolos) y `PCOM-COMU-2342` (precom de Comunicaciones, 150).

---

## 4 · Qué se escribe en el panel

| Dónde | Qué |
|---|---|
| `PROJECTS.P2342` | el árbol completo, con `updated` y `lastUpload` |
| `NODE_HISTORY.P2342` | un punto por nodo con la fecha del corte |
| `KPI_HISTORY` | el punto semanal: KPI de los tres proyectos + CORP + desglose de Arqueros |

El KPI de cada proyecto y el corporativo se **releen del panel ya escrito**
sumando las hojas del árbol, no se arrastran del punto anterior: así el
corporativo no se queda pegado cuando cambia un solo proyecto. El desglose por
disciplina de Desaladora y Talabre sí se copia del punto anterior, hasta que
entreguen sus matrices.

`PPT_PREV_DISC` **no se toca**: el panel lo declara como respaldo de último
recurso y las variaciones se calculan desde `NODE_HISTORY`.

Los KPI van a **dos decimales**, igual que el resto del historial. Con uno solo,
los de este proyecto —todos por debajo de 1%— se aplastan contra 0,0 y la
variación semanal deja de verse.

---

## 5 · Después de cargar

```bash
cd ../suite_qaqc && node armar_suite.js
cd .. && python3 verificar_suite.py
python3 auditoria_datos.py <carpeta del corte>
```

La sección 4 de la auditoría cruza el panel contra las matrices recalculadas,
contra su propio historial y contra el KPI guardado. Es el cruce que habría
pillado lo de Obras Civiles antes de enviar.
