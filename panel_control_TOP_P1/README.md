# Panel de Control — Carpetas TOP · Caminatas · Detalles de Terminación P1

Entregable ejecutivo para medir el estatus de **armado y entrega de carpetas TOP**,
el **avance de caminatas (80% / 100%)** y los **detalles de construcción en condición P1**
de la planta concentradora, con resumen general y desglose por **área** y **disciplina**.

**Corte de datos:** 16-jul-2026
**Fuente:** `Estatus.xlsx` — hojas «BD Caminatas-CTOP» (135 subsistemas) y «BD Detalles Terminación» (2.963 registros).

## Contenido de la carpeta

| Archivo | Qué es |
|---|---|
| `Panel_Control_TOP_P1.pptx` | Presentación ejecutiva (7 láminas): portada, resumen, semáforo por área y una lámina por frente + foco de gestión. |
| `index.html` | Panel de control interactivo (un solo archivo, sin dependencias externas). Se abre en cualquier navegador; incluye modo claro/oscuro, tooltips y tablas de detalle desplegables. |
| `estatus_datos.json` | Datos consolidados y agregados que alimentan el panel, calculados desde el Excel. |

## Indicadores principales (corte 16-jul-2026)

| Frente | Indicador | Valor |
|---|---|---|
| Caminatas | 80% realizadas | **135/135 · 100%** |
| Caminatas | 100% realizadas | **90/135 · 66,7%** |
| Carpetas TOP | Entregadas | **26/135 · 19,3%** (0 aprobadas · 8 rechazadas · 17 en revisión · 1 observada) |
| Detalles P1 | Cierre | **1.147/1.988 · 57,7%** (830 abiertos · 401 vencidos) |

**Semáforo por área:** 2000 *Al día* · 3000 *Crítico* · 4000 *Atención* · 8000 *Atención*.
El área **3000 (Molienda–Flotación)** es la ruta crítica: concentra 41 de las 45 caminatas 100%
pendientes y 707 de los 830 detalles P1 abiertos (85%).

## Definiciones y criterios

- **Abierto** = *Trabajo requerido* + *Iniciado* + *Trabajo no aceptado*.
- **Vencido** = detalle abierto cuya *Fecha de vencimiento* es anterior al corte.
- **% cierre P1** = Cerrados / Total P1 levantados. El estatus *Listo para revisión* se reporta aparte.
- **Carpeta TOP entregada** = con estatus *En Revisión*, *Observada*, *Rechazada* o *Aprobada*.
- **Criterio del semáforo:** *Crítico* si caminata 100% < 60% o cierre P1 < 50% · *Atención* si algún
  frente está entre 60–90% o quedan carpetas TOP sin entregar · *Al día* si las caminatas están
  completas y P1 ≥ 90%.

## Nota de reconciliación

El resumen interno del Excel muestra en un bloque «Total CTOP = 134» (excluye un alcance adicional
sin código de subsistema), mientras que su indicador final de cumplimiento usa el universo completo
de **135** subsistemas (26/135 = 19,3%). Este panel adopta el universo de 135 para mantener una base
única y consistente entre los tres frentes; por eso el porcentaje de entrega puede diferir en una
décima respecto de algún bloque intermedio de la planilla.

## Cómo regenerar los entregables

```bash
# Panel HTML: se edita directamente (index.html es autocontenido).
# PPTX: requiere Node + pptxgenjs
node gen_ppt.js        # ver script en el historial de trabajo (scratchpad)
```

Los datos se recalcularon íntegramente desde `Estatus.xlsx`; ninguna cifra está escrita a mano.
