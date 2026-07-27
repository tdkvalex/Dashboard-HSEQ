# Control de No Conformidades — módulo de la Suite QAQC

Dashboard de los hallazgos de calidad de los tres proyectos **y de Oficina Central**.

**Corte 27-07-2026 · 436 registros** desde 10-02-2023.

| Frente | Levantados | Cerrados | Abiertos | % cierre | Mediana de cierre |
|---|---|---|---|---|---|
| Desaladora | 95 | 92 | 3 | 96,8% | 19 días |
| Talabre | 239 | 213 | 26 | 89,1% | 22 días |
| Arqueros | 94 | 93 | 1 | 98,9% | **139 días** |
| Oficina Central | 8 | 8 | 0 | 100% | 77 días |
| **Total** | **436** | **406** | **30** | **93,1%** | **28 días** |

**Semáforo:** Talabre *Crítico* · el resto *Al día*.

---

## Actualización

```bash
python3 no_conformidades.py --data /ruta/Data_NCR.xlsx
```

Acepta `--hoy AAAA-MM-DD` para fijar la fecha de referencia. Escribe `datos_nc.json`, inyecta
los datos en `index.html` y avisa de todo lo que no pueda clasificar.

Después, para que la suite lo tome:

```bash
cd ../suite_qaqc && node armar_suite.js
```

No hay que copiar nada: el generador lee `modulo_nc/index.html` directamente.

---

## Lo que hay que saber del dato

### El atraso NO es calculable con este archivo

La regla del proyecto es: **NC abierta cuya fecha de cierre comprometida ya venció**. Está
implementada en el script, pero **ninguna de las 30 NC abiertas tiene fecha**: la columna
«Fecha De Cierre» solo se llena **cuando la NC se cierra**, así que es la fecha de cierre real,
no un compromiso.

El panel **no reporta 0 atrasadas** —eso se leería como un buen resultado—, sino que lo declara
abiertamente en «Control de calidad del dato» y usa en su lugar la **antigüedad**: días desde
que se levantó cada hallazgo abierto. Al corte: mediana 108 días, y **2 llevan más de un año**.

> Para activar el indicador de atraso basta agregar al Excel una columna de **fecha
> comprometida de cierre**. El script ya la calcularía sin cambios.

### Otras definiciones

- **Interna** = la levanta Besalco (`Interna BSMT`) · **Externa** = la levanta el cliente o un
  subcontrato (`Externa Cliente`, `Externa Subcontrato`).
- **Abierta** = todo estatus distinto de «Cerrado»: *Iniciado*, *Listo para revisión* y
  *No aceptado*. Se conserva el estatus original, porque «No aceptado» no se gestiona igual
  que «Iniciado».
- **Tiempo de cierre** = días entre la fecha de creación y la de cierre, solo sobre las
  cerradas. Es el indicador de qué tan rápido reacciona cada proyecto.
- **Oficina Central** genera NC pero no es una obra: aparece en el resumen corporativo y en el
  consolidado «Los 3 proyectos» se excluye, para que sea comparable con los otros módulos.
- **180 de 436 hallazgos no declaran costo**, así que las 10.783 UF son un piso, no el costo real.

---

## Lo que muestran los datos

**Talabre concentra el problema:** 26 de las 30 NC abiertas, y es el único con más externas
(186) que internas (53) — el cliente le levanta más hallazgos de los que detecta solo.

**Arqueros cierra 6 veces más lento:** 139 días de mediana contra 22 de Talabre y 19 de
Desaladora, pese a tener solo 1 NC abierta. Cierra todo, pero tarde.

**Obras Civiles domina** con 167 hallazgos de 436 (38%) y 13 abiertos.

**El volumen se disparó en 2025** (293 hallazgos contra 18 en 2024), y 2026 ya lleva 122.

---

## Archivos

| Archivo | Qué es |
|---|---|
| `index.html` | El dashboard, un solo archivo portable con 4 pestañas |
| `no_conformidades.py` | **Punto de entrada.** Procesa el Excel y actualiza el dashboard |
| `datos_nc.json` | Datos consolidados; la suite lee de aquí su KPI |
| `besalco_logo*.png` | Logo corporativo (versión blanca para fondo oscuro) |
