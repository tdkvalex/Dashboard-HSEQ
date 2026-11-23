#!/usr/bin/env python3
"""
Actualiza la Suite QAQC completa desde una carpeta con los archivos del corte.

    python3 actualizar_semana.py --entrada ~/Descargas/corte_10-08

Eso es todo: reconoce cada archivo, corre los tres módulos en el orden correcto
y deja `suite_qaqc/Suite_QAQC.html` listo para enviar.

------------------------------------------------------------------------------
POR QUÉ EXISTE
------------------------------------------------------------------------------
La actualización semanal son 6 comandos con 8 rutas de archivo, y el orden
importa: los `gen_ppt.js` van después de su script de Python y `armar_suite.js`
al final. Escribir eso a mano cada lunes es lento y se presta a dos errores que
ya pasaron:

  · usar un archivo viejo sin darse cuenta (el `Programa_de_Caminatas.xlsm` de
    Talabre trae una hoja «DT» que parece el registro pero es de junio);
  · olvidar `--externas` en No Conformidades, con lo que Arqueros pierde las
    36 NC que le levanta el cliente y aparece con 98,9% de cierre.

Este script reconoce cada archivo **por su contenido** —no por el nombre, que
cambia— y avisa antes de procesar si alguno viene más viejo que lo ya publicado.

------------------------------------------------------------------------------
CÓMO RECONOCE CADA ARCHIVO
------------------------------------------------------------------------------
Abre cada .xlsx/.xlsm de la carpeta y mira sus hojas y encabezados:

  Desaladora · reporte     hoja «REPORTE GERENCIAL»
  Desaladora · punch       hoja «LISTADO PUNCH ITEMS»
  Talabre · status         hoja «STATUS» con ÁREA + SUBSISTEMAS
  Talabre · registro DT    hoja «DT» con «N° SUBSISTEMA» y «FECHA DE CIERRE»
                           (el extracto de pendientes que viene dentro del
                           archivo de status tiene 10 columnas y no califica)
  Arqueros                 hojas «BD Caminatas-CTOP» + «BD Detalles Terminación»
  No Conformidades         hoja «Observaciones»
  NC externas del cliente  hoja «Disposición NC-Externas»
  Protocolos               .html con el bloque PROJECTS del dashboard

Si un rol no aparece en la carpeta, ese módulo **no se toca**: conserva el corte
anterior y el script lo dice. Así se puede actualizar solo lo que llegó.

------------------------------------------------------------------------------
OPCIONES
------------------------------------------------------------------------------
  --entrada RUTA     carpeta con los archivos del corte (obligatorio)
  --corte AAAA-MM-DD fecha de referencia para atrasos y para la semana de NC.
                     Por defecto, hoy. **Usarla si se reprocesa un corte pasado**
  --solo a,b         correr solo algunos módulos: cierre, nc, protocolos, suite
  --dry-run          mostrar qué se detectó y qué se correría, sin ejecutar nada
"""

import argparse
import hashlib
import json
import shutil
import subprocess
import sys
import unicodedata
import warnings
from datetime import datetime
from pathlib import Path

warnings.filterwarnings("ignore")

sys.path.insert(0, str(Path(__file__).resolve().parent))
try:
    from qaqc_excel import hojas, libro
except ImportError:
    sys.exit("Falta qaqc_excel.py junto a este script.")

RAIZ = Path(__file__).resolve().parent
QAQC = RAIZ / "panel_control_TOP_P1"
NC = RAIZ / "modulo_nc"
SUITE = RAIZ / "suite_qaqc"

VERDE, ROJO, AMBAR, GRIS, FIN = "\033[32m", "\033[31m", "\033[33m", "\033[90m", "\033[0m"
OK, MAL, AV = f"{VERDE}✓{FIN}", f"{ROJO}✕{FIN}", f"{AMBAR}▲{FIN}"


def norm(v):
    """Minúsculas, sin tildes y sin espacios sobrantes: las hojas vienen con y
    sin acento según quién guardó el archivo («Disposición» / «Disposicion»)."""
    if v is None:
        return ""
    s = " ".join(str(v).strip().lower().split())
    return "".join(c for c in unicodedata.normalize("NFD", s)
                   if unicodedata.category(c) != "Mn")


# =============================================================================
# 1) Reconocimiento de archivos
# =============================================================================
# Reconocer la carpeta es lo primero que pasa cada lunes y era lo más lento de
# todo: abrir cada libro entero solo para ver sus hojas costaba hasta 8 s por
# archivo. Los nombres de hoja salen del ZIP en milisegundos, y cuando de verdad
# hay que leer celdas se abre UNA vez y se reutiliza (`_ABIERTOS`), en vez de
# tres veces el mismo archivo.
_ABIERTOS = {}


def _abrir(ruta):
    """El libro de `ruta`, abierto una sola vez por corrida."""
    if ruta not in _ABIERTOS:
        try:
            _ABIERTOS[ruta] = libro(ruta)
        except SystemExit:
            _ABIERTOS[ruta] = None
    return _ABIERTOS[ruta]


def encabezados(ruta, hoja, max_filas=12):
    """Todos los textos de las primeras filas de una hoja, normalizados."""
    wb = _abrir(ruta)
    if wb is None or hoja not in wb.sheetnames:
        return set(), 0
    try:
        vals, anchos = set(), 0
        for f in wb[hoja].iter_rows(min_row=1, max_row=max_filas, values_only=True):
            anchos = max(anchos, sum(1 for v in f if v not in (None, "")))
            for v in f:
                if isinstance(v, str) and v.strip():
                    vals.add(norm(v))
        return vals, anchos
    except Exception:
        return set(), 0


def fecha_max(ruta, hoja, col, desde=2):
    """Fecha más reciente de una columna. Sirve para detectar archivos viejos."""
    wb = _abrir(ruta)
    if wb is None or hoja not in wb.sheetnames:
        return None
    try:
        mx = None
        for f in wb[hoja].iter_rows(min_row=desde, values_only=True):
            if col < len(f) and isinstance(f[col], datetime):
                if mx is None or f[col] > mx:
                    mx = f[col]
        return mx
    except Exception:
        return None


def clasificar(ruta):
    """Devuelve (rol, fecha_del_dato) o (None, None) si no se reconoce."""
    if ruta.suffix.lower() in (".html", ".htm"):
        # El dashboard de Protocolos pesa ~3 MB y su bloque PROJECTS aparece
        # pasado el byte 890.000: hay que leerlo entero, no solo la cabecera.
        try:
            txt = ruta.read_text(encoding="utf-8", errors="ignore")
        except Exception:
            return None, None
        if "PROJECTS" in txt and "P2416" in txt and "NODE_HISTORY" in txt:
            return "protocolos", None
        return None, None

    hs = hojas(ruta)
    if not hs:
        return None, None
    hn = {norm(h): h for h in hs}

    if "reporte gerencial" in hn:
        return "des_reporte", None
    if "listado punch items" in hn:
        return "des_punch", None
    if any(h.startswith("bd caminatas") for h in hn) and \
       any(h.startswith("bd detalles") for h in hn):
        return "arqueros", None
    # Export de P1 de Arqueros: hoja «Punch list» con «Atrasado» y la fecha de
    # «Listo Para Revisión», dos columnas que el libro principal no trae. No
    # alimenta ningún módulo —el libro principal sigue siendo la fuente—: lo usa
    # auditoria_datos.py para recontar los P1 por una vía independiente.
    if "punch list" in hn:
        cab, _ = encabezados(ruta, hn["punch list"], max_filas=2)
        if "atrasado" in cab and "estatus" in cab and "especialidad" in cab:
            return "arq_p1", None
    if "observaciones" in hn:
        return "nc_data", fecha_max(ruta, hn["observaciones"], 11)
    if "disposicion nc-externas" in hn:
        return "nc_externas", fecha_max(ruta, hn["disposicion nc-externas"], 6, desde=5)

    # Talabre: el archivo de status trae también una hoja «DT», pero es el
    # extracto de pendientes (10 columnas). El registro completo tiene la
    # fecha de cierre y más de 20 columnas.
    if "dt" in hn:
        cab, ancho = encabezados(ruta, hn["dt"])
        if "n° subsistema" in cab and "fecha de cierre" in cab and ancho >= 20:
            fila = 10 if "status" not in hn else 2
            return "tal_dt", fecha_max(ruta, hn["dt"], 18, desde=fila)
    if "status" in hn:
        cab, _ = encabezados(ruta, hn["status"])
        if "area" in cab or "área" in cab or "subsistemas" in cab:
            return "tal_status", None
    return None, None


ROLES = {
    "des_reporte": "Desaladora · reporte gerencial",
    "des_punch":   "Desaladora · listado punch",
    "tal_status":  "Talabre · status de subsistemas",
    "tal_dt":      "Talabre · registro de detalles de terminación",
    "arqueros":    "Arqueros · estatus resumen general",
    "nc_data":     "No Conformidades · registro principal",
    "nc_externas": "No Conformidades · planilla del cliente MASA",
    "protocolos":  "Protocolos · dashboard actualizado",
}

# Archivos que NO alimentan ningún módulo: sirven para contrastar lo calculado
# por una vía independiente. Que no lleguen no cambia ninguna cifra, así que su
# ausencia no se anuncia como «se conserva el corte anterior».
ROLES_CONTRASTE = {
    "arq_p1": "Arqueros · export de P1 (contraste de la auditoría)",
}


def explorar(carpeta):
    """Clasifica todo lo que hay en la carpeta. Si dos archivos compiten por el
    mismo rol, gana el que trae el dato más nuevo y se avisa del descartado."""
    hallados, descartes, sin_rol = {}, [], []
    for r in sorted(carpeta.rglob("*")):
        if not r.is_file() or r.name.startswith("~$"):
            continue
        if r.suffix.lower() not in (".xlsx", ".xlsm", ".html", ".htm"):
            continue
        rol, fecha = clasificar(r)
        if rol is None:
            sin_rol.append(r)
            continue
        previo = hallados.get(rol)
        if previo is None:
            hallados[rol] = (r, fecha)
        else:
            viejo, f_viejo = previo
            gana_nuevo = (fecha or datetime.min) > (f_viejo or datetime.min)
            hallados[rol] = (r, fecha) if gana_nuevo else previo
            descartes.append((rol, r if not gana_nuevo else viejo,
                              fecha if not gana_nuevo else f_viejo))
    return hallados, descartes, sin_rol


# =============================================================================
# 2) Contraste contra lo ya publicado — el archivo nuevo no puede ser más viejo
# =============================================================================
def corte_publicado():
    """Fecha del dato de cada módulo tal como está hoy en el repositorio."""
    out = {}
    for clave, ruta, camino in (
        ("nc_data", NC / "datos_nc.json", ("control", "rangoFechas", 1)),
        # Se compara contra la EMISIÓN más nueva del registro anterior, no
        # contra el corte: las emisiones siempre son anteriores al corte y
        # compararlas contra él da falsas alarmas todas las semanas.
        ("tal_dt", QAQC / "datos_talabre.json", ("control", "ultimaEmisionDT")),
    ):
        if not ruta.exists():
            continue
        try:
            d = json.loads(ruta.read_text(encoding="utf-8"))
            for paso in camino:
                d = d[paso] if not isinstance(paso, int) else d[paso]
            out[clave] = d
        except Exception:
            pass
    return out


# =============================================================================
# 2.b) Huellas — el archivo repetido
# =============================================================================
# Un archivo que llega byte a byte igual al del corte anterior no rompe nada
# visible: el módulo recalcula lo mismo y el panel sale entero, con el dato de
# la semana pasada y la fecha de hoy. Ha pasado de verdad —el registro de DT de
# Talabre, el Estatus de MASA, y una vez un archivo rotulado «Talabre» que era
# el libro de MASA—. Se guarda la huella de lo procesado para avisarlo solo.
# No DETIENE la corrida: repetido no siempre es un error (desde el 23-11-2026
# el registro de DT de Talabre está legítimamente congelado y el proyecto lo
# confirma en su propia hoja de pendientes). El archivo más VIEJO que el
# cargado sí detiene, y eso lo cubre el contraste de fechas de arriba.
REGISTRO = RAIZ / "cortes_procesados.json"


def huella(ruta):
    """sha256 del archivo tal como llegó."""
    h = hashlib.sha256()
    try:
        with open(ruta, "rb") as f:
            for trozo in iter(lambda: f.read(1 << 20), b""):
                h.update(trozo)
    except OSError:
        return None
    return h.hexdigest()


def registro_leer():
    if not REGISTRO.exists():
        return {}
    try:
        return json.loads(REGISTRO.read_text(encoding="utf-8"))
    except Exception:
        return {}


def registro_guardar(hallados, hoy):
    """Deja la huella de cada archivo procesado, para contrastar el próximo corte."""
    reg = registro_leer()
    for rol, (ruta, _) in hallados.items():
        hh = huella(ruta)
        if hh:
            reg[rol] = {"sha256": hh, "archivo": ruta.name, "corte": hoy}
    try:
        REGISTRO.write_text(json.dumps(reg, ensure_ascii=False, indent=1) + "\n",
                            encoding="utf-8")
    except OSError as e:
        print(f"  {AV} No se pudo guardar {REGISTRO.name}: {e}")


def repetidos(hallados):
    """Qué archivos llegan idénticos al que ya se procesó. Avisa, no detiene."""
    reg, notas = registro_leer(), []
    for rol, (ruta, _) in sorted(hallados.items()):
        ant = reg.get(rol)
        if not ant or ant.get("sha256") != huella(ruta):
            continue
        etq = ROLES.get(rol) or ROLES_CONTRASTE.get(rol, rol)
        notas.append(f"{etq}: «{ruta.name}» llega IDÉNTICO byte a byte al que se "
                     f"procesó el {ant.get('corte', '—')}. Si el proyecto no emitió "
                     f"nada nuevo, está bien; si debía haber emitido, es el archivo "
                     f"equivocado.")
    return notas


def dmy(f):
    return f.strftime("%d-%m-%Y") if isinstance(f, datetime) else "—"


def revisar(hallados, publicado):
    """Lo que hay que mirar ANTES de procesar. Devuelve la lista de avisos.

    Vive aquí, en una sola función, porque la usan dos frentes: la corrida por
    terminal y el Centro de Carga. Si cada uno llevara su copia, un aviso
    corregido en un lado seguiría faltando en el otro — y estos avisos son justo
    los que evitan publicar un corte con cifras equivocadas.
    """
    alertas = []
    for clave, etq, extra in (
        ("nc_data", "El registro de NC", ""),
        ("tal_dt", "El registro de DT de Talabre", " (ojo con Programa_de_Caminatas.xlsm)"),
    ):
        if clave not in hallados or clave not in publicado:
            continue
        f = hallados[clave][1]
        try:
            prev = datetime.strptime(publicado[clave], "%d-%m-%Y")
        except (ValueError, TypeError):
            continue
        if f and f < prev:
            alertas.append(f"{etq} llega hasta {dmy(f)} y lo que ya está cargado cubre "
                           f"hasta {publicado[clave]}: parece un archivo viejo{extra}.")
    if "nc_data" in hallados and "nc_externas" not in hallados:
        alertas.append("Falta la planilla de NC externas del cliente. Sin ella Arqueros "
                       "pierde las 36 NC que le levanta MASA y sale con 98,9% de cierre "
                       "en vez de 78%.")
    # Talabre NO entra aquí: procesa con uno solo y declara lo que conserva.
    for par, proyecto, que in (
        ({"des_reporte", "des_punch"}, "Desaladora", "reporte + punch"),
    ):
        if par & hallados.keys() and not par <= hallados.keys():
            alertas.append(f"{proyecto} necesita SUS DOS archivos ({que}). "
                           f"Con uno solo no se procesa.")
    return alertas


# =============================================================================
# 3) Ejecución
# =============================================================================
def correr(cmd, cwd, titulo, dry):
    print(f"\n{GRIS}$ {' '.join(str(c) for c in cmd)}{FIN}")
    if dry:
        return True, ""
    p = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
    salida = (p.stdout or "") + (p.stderr or "")
    if p.returncode != 0:
        print(f"{MAL} {titulo} falló:\n{salida.strip()[-1500:]}")
        return False, salida
    # De la salida larga de cada script solo interesan los avisos y el cierre.
    for linea in salida.splitlines():
        if linea.strip().startswith(("⚠", "·", "OK", "≠", "i ")) or "Escrit" in linea \
                or "generada" in linea or "embebid" in linea:
            print("   " + linea.strip())
    print(f"{OK} {titulo}")
    return True, salida


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--entrada", required=True, help="Carpeta con los archivos del corte")
    ap.add_argument("--corte", help="Fecha de referencia AAAA-MM-DD (por defecto, hoy)")
    ap.add_argument("--solo", help="cierre,nc,protocolos,suite — por defecto, todo")
    ap.add_argument("--dry-run", action="store_true", help="Solo mostrar el plan")
    ap.add_argument("--igual", action="store_true",
                    help="Procesar aunque haya avisos (archivo viejo, falta la planilla "
                         "del cliente…). Sin esto, los avisos detienen la corrida.")
    args = ap.parse_args()

    entrada = Path(args.entrada).expanduser()
    if not entrada.is_dir():
        sys.exit(f"No es una carpeta: {entrada}")
    hoy = args.corte or datetime.now().strftime("%Y-%m-%d")
    try:
        datetime.strptime(hoy, "%Y-%m-%d")
    except ValueError:
        sys.exit("--corte debe ser AAAA-MM-DD")
    modulos = {m.strip() for m in (args.solo or "cierre,nc,protocolos,suite").split(",")}

    print("=" * 74)
    print(f"  SUITE QAQC — actualización al corte {hoy}")
    print(f"  Entrada: {entrada}")
    print("=" * 74)

    hallados, descartes, sin_rol = explorar(entrada)
    publicado = corte_publicado()

    print("\nARCHIVOS RECONOCIDOS")
    for rol, etq in ROLES_CONTRASTE.items():
        if rol in hallados:
            print(f"  {OK} {etq:46s} {hallados[rol][0].name}")
    for rol, etq in ROLES.items():
        if rol in hallados:
            r, f = hallados[rol]
            extra = f"  {GRIS}dato hasta {dmy(f)}{FIN}" if f else ""
            print(f"  {OK} {etq:46s} {r.name}{extra}")
        else:
            print(f"  {AV} {etq:46s} {AMBAR}no vino — se conserva el corte anterior{FIN}")

    for rol, r, f in descartes:
        print(f"  {AV} {ROLES[rol]}: se descartó «{r.name}» (dato hasta {dmy(f)}), "
              f"hay uno más nuevo")
    if sin_rol:
        print(f"\n{GRIS}Sin clasificar (se ignoran): "
              f"{', '.join(r.name for r in sin_rol[:6])}{FIN}")

    notas = repetidos(hallados)
    if notas:
        print(f"\n{AMBAR}ARCHIVOS QUE LLEGAN IGUALES AL CORTE ANTERIOR{FIN}")
        for n in notas:
            print(f"  {AV} {n}")

    alertas = revisar(hallados, publicado)

    # Los avisos DETIENEN la corrida. Un archivo viejo o una planilla que falta
    # no rompen nada visible: el panel sale entero y con cifras equivocadas, que
    # es peor. Para seguir hay que decirlo a propósito con --igual.
    if alertas:
        print(f"\n{AMBAR}AVISOS ANTES DE PROCESAR{FIN}")
        for a in alertas:
            print(f"  {AV} {a}")
        if not args.igual and not args.dry_run:
            print(f"\n{MAL} No se procesó nada. Revisa los avisos y, si igual quieres "
                  f"seguir, repite el comando con {AMBAR}--igual{FIN}.")
            sys.exit(2)

    g = lambda k: str(hallados[k][0]) if k in hallados else None
    pasos = []
    if "cierre" in modulos:
        if g("des_reporte") and g("des_punch"):
            pasos.append((["python3", "desaladora.py", "--reporte", g("des_reporte"),
                           "--punch", g("des_punch"), "--hoy", hoy], QAQC, "Desaladora"))
        # Talabre se procesa con uno solo de sus dos archivos: el que no llega
        # conserva su bloque del corte anterior y talabre.py lo declara. Desde
        # el corte 23-11-2026 el proyecto manda un libro único de
        # PROTOCOLIZACIÓN y su registro de DT puede no venir.
        if g("tal_status") or g("tal_dt"):
            cmd = ["python3", "talabre.py", "--hoy", hoy]
            if g("tal_status"):
                cmd += ["--status", g("tal_status")]
            if g("tal_dt"):
                cmd += ["--dt", g("tal_dt")]
            pasos.append((cmd, QAQC, "Talabre"))
        if g("arqueros"):
            pasos.append((["python3", "actualizar.py", g("arqueros")], QAQC, "Arqueros"))
        # La PPT va DESPUÉS de los tres: los lee a los tres y se embebe en el panel.
        if any(p[2] in ("Desaladora", "Talabre", "Arqueros") for p in pasos):
            pasos.append((["node", "gen_ppt.js"], QAQC, "PPT de Cierre QAQC"))
    if "nc" in modulos and g("nc_data"):
        cmd = ["python3", "no_conformidades.py", "--data", g("nc_data"), "--hoy", hoy]
        if g("nc_externas"):
            cmd += ["--externas", g("nc_externas")]
        pasos.append((cmd, NC, "No Conformidades"))
        pasos.append((["node", "gen_ppt.js"], NC, "PPT de No Conformidades"))

    if not pasos and "suite" not in modulos:
        sys.exit(f"\n{MAL} No hay nada que procesar con los archivos de esa carpeta.")

    print(f"\n{'PLAN (dry-run, no se ejecuta nada)' if args.dry_run else 'PROCESANDO'}")
    for cmd, cwd, titulo in pasos:
        ok, _ = correr(cmd, cwd, titulo, args.dry_run)
        if not ok:
            sys.exit(f"\n{MAL} Se detuvo en «{titulo}». Nada más se ejecutó.")

    # El dashboard de Protocolos lo arma otro equipo: se copia a modulos/ y
    # sobre esa copia se genera aquí su informe, igual que en los otros dos
    # módulos. El generador va DESPUÉS de la copia —lee el dashboard del corte—
    # y ANTES de `armar_suite.js`, que empaqueta el módulo ya con el informe
    # embebido.
    if "protocolos" in modulos and g("protocolos"):
        destino = SUITE / "modulos" / "protocolos.html"
        print(f"\n{GRIS}$ cp {g('protocolos')} {destino}{FIN}")
        if not args.dry_run:
            destino.parent.mkdir(exist_ok=True)
            shutil.copyfile(g("protocolos"), destino)
        print(f"{OK} Protocolos copiado a modulos/")
        ok, _ = correr(["node", "gen_ppt_protocolos.js"], SUITE,
                       "Informe de Protocolos", args.dry_run)
        if not ok:
            sys.exit(1)

    if "suite" in modulos:
        ok, _ = correr(["node", "armar_suite.js"], SUITE, "Suite QAQC", args.dry_run)
        if not ok:
            sys.exit(1)

    if args.dry_run:
        print(f"\n{GRIS}Dry-run: no se modificó ningún archivo.{FIN}")
        return

    registro_guardar(hallados, hoy)

    print("\n" + "=" * 74)
    print(f"  {OK} LISTO — suite_qaqc/Suite_QAQC.html")
    print("=" * 74)
    print("""
ANTES DE ENVIAR, REVISAR:
  1. Los AVISOS de arriba: cada script informa lo que no pudo clasificar y los
     cruces que dejaron de cuadrar. Si aparece uno nuevo, investigar.
  2. Abrir la suite y recorrer las 4 pestañas de cada módulo.
  3. Que el botón «Descargar Informe» entregue la PPT en los dos módulos.
  4. Que el corte de la franja superior sea el que corresponde.

  Verificación automática:  python3 verificar_suite.py
""")


if __name__ == "__main__":
    main()
