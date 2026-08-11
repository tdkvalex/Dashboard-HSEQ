# Dashboard HSEQ — Besalco Montajes

Dashboards y paneles de control **HSEQ**: Seguridad y Salud Ocupacional, Medio
Ambiente y Calidad/QAQC. Este repositorio es de uso exclusivo para eso.

## Qué hay hoy

**Suite QAQC** — el entregable semanal de calidad. Un solo archivo HTML de ~8,7 MB,
sin dependencias externas, que se envía por correo tal cual y reúne tres módulos
sobre los mismos tres proyectos:

| Módulo | Qué mide |
|---|---|
| **Protocolos** | Registros constructivos pendientes |
| **Cierre QAQC** | Carpetas, caminatas y detalles de terminación |
| **No Conformidades** | Hallazgos y su corrección |

| Proyecto | Cliente |
|---|---|
| P2416 · Desaladora | ANTOFAGASTA MINERALS |
| P2407 · Talabre | CODELCO |
| P2342 · Arqueros | MASA |

## Cómo se actualiza

Doble clic en **«Abrir Centro de Carga»**: se abre una página local, se arrastran
los archivos del corte y se descarga la Suite armada. Sin terminal.

Por consola es lo mismo:

```bash
python3 actualizar_semana.py --entrada <carpeta del corte>
python3 verificar_suite.py
python3 auditoria_datos.py <la misma carpeta>
```

El procedimiento completo —qué pedir, a quién, las trampas conocidas y qué mirar
antes de enviar— está en **[`CADA_LUNES.md`](CADA_LUNES.md)**. Las reglas de
negocio y las decisiones de criterio, en **[`CLAUDE.md`](CLAUDE.md)**.

## Cómo se revisa

Nada se publica sin pasar por dos verificadores, y los dos salen con error si algo
falla:

- **`verificar_suite.py`** — la forma: que los tres módulos monten en 1366, 1920 y
  2560 px, que no haya desbordes ni errores de consola, que las PPT tengan su
  numeración y logo, que sus etiquetas sean legibles (≥4,5:1), que los desgloses
  cuadren con su fracción y que el botón «Descargar Informe» entregue la PPT
  **vigente**, no una anterior.
- **`auditoria_datos.py`** — el fondo: que la misma cifra diga lo mismo en el Excel
  de origen, en el JSON de cada módulo, en la portada de la suite y en el texto de
  las PPT. 71 comprobaciones.

## Versión online (disponible, sin activar)

Existe un workflow que publica la suite en **GitHub Pages** con cada corte:
arma el archivo en el servidor con el mismo `armar_suite.js`, tras pasar la
auditoría de datos y un control de peso. Está probado de punta a punta, pero
**hoy no se usa**: el entregable es el archivo que se manda por correo.

Para activarlo: Settings → Pages → Source «GitHub Actions», y descomentar el
bloque `push` de `.github/workflows/publicar.yml`.

## Aviso

Repositorio **público**. El contenido incluye nombres de cliente, códigos de
contrato, detalle de no conformidades y costos en UF.
