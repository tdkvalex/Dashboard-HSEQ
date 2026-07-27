# Panel de Control — Carpetas TOP · Caminatas · Detalles de Terminación P1

Entregable ejecutivo para medir el estatus de **armado y entrega de carpetas TOP**,
el **avance de caminatas (80% / 100%)** y los **detalles de construcción en condición P1**
de la planta concentradora, con resumen general y desglose por **área** y **disciplina**.

**Corte vigente:** 26-07-2026
**Fuente:** `Estatus_Resumen_General_QAQC.xlsx` — hojas «BD Caminatas-CTOP» (135 subsistemas)
y «BD Detalles Terminación» (2.963 registros).

---

## 🔄 Actualización semanal (lunes) — 2 comandos

```bash
python3 actualizar.py /ruta/al/Estatus_Resumen_General_QAQC.xlsx   # datos + panel HTML
node gen_ppt.js                                                    # PPT ejecutiva
```

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
| `index.html` | Panel interactivo, **un solo archivo portable** (se puede enviar por correo tal cual). Modo claro/oscuro, tooltips, variación semanal y tablas de detalle. |
| `Panel_Control_TOP_P1.pptx` | Presentación ejecutiva de 7 láminas. |
| `actualizar.py` | Lee el Excel, recalcula todo y actualiza el panel. **Punto de entrada semanal.** |
| `gen_ppt.js` | Regenera la PPT desde los datos ya calculados. |
| `estatus_datos.json` | Datos consolidados del corte vigente. |
| `historial.json` | Un registro por corte; alimenta la variación semana a semana. |

---

## Indicadores al corte 26-07-2026

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

## Definiciones y criterios

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
