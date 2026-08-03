# Cada lunes — actualización de la Suite QAQC

Dos comandos. El resto es pedir bien los archivos.

```bash
python3 actualizar_semana.py --entrada ~/Descargas/corte_10-08
python3 verificar_suite.py
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

**Si se reprocesa un corte pasado, va `--corte`.** La ventana de «levantados en la
última semana» y los días de atraso se cuentan contra esa fecha; sin el parámetro
se usa el día en que se corre.

```bash
python3 actualizar_semana.py --entrada ~/Descargas/corte_03-08 --corte 2026-08-03
```

Cuando el script se detiene por un aviso no procesó nada: se corrige y se repite.
Si el aviso está entendido y se quiere seguir igual, se agrega `--igual`.

---

## 3 · Antes de enviar

`verificar_suite.py` revisa solo lo que se olvida a mano y **sale con error si
algo falla**:

- las dos PPT: numeración correlativa, logo en todas, ningún texto montado sobre
  el pie, etiquetas de gráfico legibles (≥4,5:1 contra el número blanco);
- la suite en 1366, 1920 y 2560 px: los tres módulos montan, se recorren todas
  sus pestañas, no hay desbordes horizontales ni errores de consola;
- que el botón «Descargar Informe» entregue un `.pptx` de verdad;
- que los detalles sumen lo mismo en los proyectos y en el consolidado.

Con `--rapido` corre solo a 1920 px, para una pasada corta.

Lo que **sí** hay que mirar a ojo:

1. **Los AVISOS de cada script.** Ahí aparecen los valores que no supo clasificar
   y los cruces que dejaron de cuadrar. Si sale uno nuevo, investigar antes de
   enviar: casi siempre es que el archivo de origen cambió.
2. **El corte de la franja superior.** Si los módulos quedaron a fechas distintas
   aparece el aviso «cortes mixtos» con el detalle al pasar el cursor.
3. **Una pasada por las pestañas**, sobre todo si algún proyecto cambió de formato.

---

## 4 · Si algo se rompe

| Síntoma | Causa probable |
|---|---|
| «no vino — se conserva el corte anterior» en un archivo que sí mandaste | el archivo no tiene las hojas esperadas: ábrelo y compara con la tabla de arriba |
| Un script avisa de columnas o valores no reconocidos | el proyecto cambió el formato del Excel. `talabre.py` resuelve columnas por nombre y aguanta; los demás usan posiciones fijas documentadas en `panel_control_TOP_P1/CLAUDE.md` §4 |
| La suite sale sin un módulo | falta su JSON o su HTML; el generador avisa y arma la suite con los que haya |
| La PPT no descarga desde el panel | se corrió `gen_ppt.js` antes que el script de Python. Repetir en orden — `actualizar_semana.py` lo hace solo |

El detalle de cada módulo está en su propio README:
[`panel_control_TOP_P1/`](panel_control_TOP_P1/README.md) ·
[`modulo_nc/`](modulo_nc/README.md) · [`suite_qaqc/`](suite_qaqc/README.md).
Las reglas de homologación entre proyectos —validadas con los jefes de calidad y
que **no se cambian sin avisar**— están en
[`panel_control_TOP_P1/CLAUDE.md`](panel_control_TOP_P1/CLAUDE.md) §3.
