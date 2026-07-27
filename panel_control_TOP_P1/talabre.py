#!/usr/bin/env python3
"""
Procesa el proyecto TALABRE (Codelco) y escribe datos_talabre.json, que alimenta
la pestaña Talabre del panel y la PPT.

Uso (cada semana, con los dos archivos del corte):

    python3 talabre.py \
        --status /ruta/TalabreSTATUS_PEC.xlsx \
        --dt     /ruta/TalabreCuadro_DT.xlsx

Opcional:
    --hoy 2026-07-27     fecha de referencia para verificar los atrasos
                         (por defecto, el día en que se corre el script)

------------------------------------------------------------------------------
DE DÓNDE SALE CADA DATO
------------------------------------------------------------------------------
STATUS_PEC · hoja «STATUS» (encabezado fila 4, datos fila 5+)
    col B  ÁREA         col C  SUBSISTEMAS     col E  CAMINATA 1
    col F  CAMINATA 2   col H  % PEC           col I  DT P1     col J  DT P2
STATUS_PEC · hoja «RESUMEN» — totales que declara el propio archivo, contra los
    que se contrasta el cálculo.

CUADRO_DT · hoja «DT» (encabezado fila 1, datos fila 2+)
    col I  N° SUBSISTEMA   col H  NOMBRE SISTEMA   col U  PRIORIDADES
    col V  FECHA DE COMPROMISO DE CIERRE           col X  FECHA DE CIERRE
    col Z  CAMINATA        col AB DISC.            col AI STATUS   col AK ÁREA

------------------------------------------------------------------------------
REGLAS DE CLASIFICACIÓN
------------------------------------------------------------------------------
ESTADO del DT — Talabre ya trae la columna STATUS calculada (Cerrado / Abierto /
    Atrasado). Se usa tal cual, pero el script VERIFICA que respete la misma
    regla que el resto de los proyectos:
        Atrasado = sin fecha de cierre y con fecha de compromiso ya vencida
    y avisa si algún registro no cuadra. Los abiertos sin fecha de compromiso no
    se dan por atrasados: se informan aparte.

CAMINATAS — se identifican por NÚMERO. En la hoja STATUS:
    «SI» = realizada · una fecha = programada para esa fecha · «NO» = pendiente
    OJO: la hoja RESUMEN publica «C1/C2 Prog.», que suma las realizadas MÁS las
    programadas de la semana. Este panel las informa por separado, porque una
    caminata agendada no es una caminata hecha. El contraste queda documentado
    en el control de calidad del dato.

CARPETAS — Talabre no maneja estados de aprobación como los otros proyectos:
    lleva un % de avance por subsistema (% PEC). Se reporta como avance, con la
    distribución por tramos, y NO se traduce a «entregadas» para no inventar un
    concepto que el proyecto no usa. El propio archivo declara además cuántas
    están en revisión del cliente.

DISCIPLINAS — se unifican los nombres del proyecto con los del resto:
    CANERIAS -> Piping · CIVIL y MOVIMIENTO DE TIERRAS -> Obras Civiles
    ELECTRICOS -> Eléctrica · INSTRUMENTACION -> Instrumentación y Control
"""

import argparse
import json
import re
import sys
import unicodedata
import warnings
from collections import Counter, defaultdict
from datetime import datetime
from math import floor
from pathlib import Path

warnings.filterwarnings("ignore")

try:
    from openpyxl import load_workbook
except ImportError:
    sys.exit("Falta openpyxl.  Instálalo con:  pip install openpyxl")

AQUI = Path(__file__).resolve().parent

PROYECTO = {
    "id": "TALABRE",
    "nombre": "Talabre",
    "descripcion": "Captación de agua, pozos y estaciones de bombeo",
    "cliente": "Codelco",
    "contrato": "4600029355",
}

DISCIPLINAS = ["Obras Civiles", "Piping", "Eléctrica",
               "Instrumentación y Control", "Estructura"]
DISC_CORTA = {"Instrumentación y Control": "Instr. y Control"}

DISC_MAP = {
    "civil": "Obras Civiles",
    "movimiento de tierras": "Obras Civiles",
    "oo.cc": "Obras Civiles",
    "canerias": "Piping",
    "piping": "Piping",
    "electricos": "Eléctrica",
    "electrica": "Eléctrica",
    "instrumentacion": "Instrumentación y Control",
    "instrumentacion y control": "Instrumentación y Control",
    "estructura": "Estructura",
}

# Estados que trae la propia columna STATUS del cuadro de DT.
CERRADO, ABIERTO, ATRASADO = "Cerrado", "Abierto", "Atrasado"

# Tramos del % PEC de la carpeta.
TRAMOS_PEC = [("95% o más", 0.95), ("80% a 95%", 0.80), ("Bajo 80%", 0.0)]

avisos = []


def norm(v):
    """Minúsculas, sin tildes, sin espacios sobrantes."""
    if v is None:
        return ""
    s = str(v).strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFD", s)
                if unicodedata.category(c) != "Mn")
    return " ".join(s.split())


def disciplina(v):
    d = DISC_MAP.get(norm(v))
    if d:
        return d
    if v not in (None, ""):
        avisos.append(f"Disciplina no reconocida: «{v}» — se agrupa en «Otras»")
    return "Otras"


def prioridad(v):
    s = norm(v).upper().replace(" ", "")
    m = re.match(r"^P(\d)", s)
    if m:
        return "P" + m.group(1)
    if s in ("", "SINASIGNAR"):
        return "Sin asignar"
    avisos.append(f"Prioridad no reconocida: «{v}» — se agrupa en «Sin asignar»")
    return "Sin asignar"


def pct1(a, b):
    """Porcentaje con 1 decimal, redondeando .5 hacia arriba (igual que el panel)."""
    if not b:
        return 0.0
    return floor(1000.0 * a / b + 0.5) / 10.0


def as_fecha(v):
    if isinstance(v, datetime):
        return v
    if isinstance(v, str) and v.strip():
        for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d"):
            try:
                return datetime.strptime(v.strip(), fmt)
            except ValueError:
                pass
    return None


# =============================================================================
# 1) STATUS_PEC — subsistemas, caminatas y avance de carpeta
# =============================================================================
def leer_status(ruta):
    wb = load_workbook(ruta, data_only=True)
    ws = wb["STATUS"]

    def estado_cam(v):
        """SI = realizada · una fecha = programada · NO o vacío = pendiente."""
        f = as_fecha(v) if not isinstance(v, str) or not norm(v) in ("si", "no") else None
        if isinstance(v, datetime):
            return "Programada", v
        s = norm(v)
        if s in ("si", "sí"):
            return "Realizada", None
        if s in ("no", ""):
            return "Pendiente", None
        if f:
            return "Programada", f
        avisos.append(f"Estado de caminata no reconocido: «{v}» — se cuenta como pendiente")
        return "Pendiente", None

    subs, corte = [], None
    for f in ws.iter_rows(min_row=5, values_only=True):
        if f[2] in (None, ""):
            continue
        e1, d1 = estado_cam(f[4])
        e2, d2 = estado_cam(f[5])
        for d in (d1, d2):
            if d and (corte is None or d > corte):
                corte = d
        subs.append({
            "id": str(f[2]).strip(),
            "area": str(f[1]).strip() if f[1] else "Sin área",
            "resp": str(f[3]).strip() if f[3] else "",
            "alcance": str(f[6]).strip() if f[6] else "",
            "cam": {1: e1, 2: e2},
            "fechaCam": {1: d1.strftime("%d-%m-%Y") if d1 else None,
                         2: d2.strftime("%d-%m-%Y") if d2 else None},
            "pec": f[7] if isinstance(f[7], (int, float)) else None,
            "dtP1": int(f[8]) if isinstance(f[8], (int, float)) else 0,
            "dtP2": int(f[9]) if isinstance(f[9], (int, float)) else 0,
            "obs": str(f[10]).strip() if f[10] else "",
        })

    # Totales que declara el propio archivo, para contrastarlos.
    declarado = {"areas": {}, "total": None}
    if "RESUMEN" in wb.sheetnames:
        for r in wb["RESUMEN"].iter_rows(min_row=3, max_row=12, values_only=True):
            vals = [r[i] for i in range(2, 8)]
            if not any(isinstance(v, (int, float)) for v in vals):
                continue
            fila = {"subs": r[2], "c1": r[3], "c2": r[4],
                    "sobre80": r[5], "sobre95": r[6], "enRevision": r[7]}
            if r[1]:
                declarado["areas"][str(r[1]).strip()] = fila
            else:
                declarado["total"] = fila
    return corte, subs, declarado


# =============================================================================
# 2) CUADRO_DT — detalles de terminación
# =============================================================================
def leer_dt(ruta, hoy):
    wb = load_workbook(ruta, data_only=True, read_only=True)
    it = wb["DT"].iter_rows(min_row=1, values_only=True)
    next(it)                                   # encabezado

    items, incoherentes = [], []
    for f in it:
        if f[0] is None:
            continue
        st = str(f[35]).strip() if f[35] else ""
        if st not in (CERRADO, ABIERTO, ATRASADO):
            avisos.append(f"STATUS de DT no reconocido: «{f[35]}» — se cuenta como abierto")
            st = ABIERTO

        comp, cierre = as_fecha(f[21]), as_fecha(f[23])
        # Verificación: ¿el STATUS del archivo respeta la misma regla que usamos
        # en los otros proyectos (atrasado = abierto con compromiso vencido)?
        if st == ATRASADO and not (cierre is None and comp and comp < hoy):
            incoherentes.append(("marcado Atrasado sin compromiso vencido", str(f[24])))
        if st == ABIERTO and cierre is None and comp and comp < hoy:
            incoherentes.append(("abierto con compromiso vencido pero no marcado", str(f[24])))
        if st == CERRADO and cierre is None:
            incoherentes.append(("marcado Cerrado sin fecha de cierre", str(f[24])))

        # El registro de DT usa áreas más finas que la hoja STATUS: todos los
        # pozos (PBO/PBBR/PBN/TB) pertenecen al sistema «Pozos Barrera
        # Hidráulica» y se agrupan en POZOS, igual que en STATUS.
        area = str(f[36]).strip() if f[36] else "Sin área"
        sistema = str(f[7]).strip() if f[7] else ""
        if norm(sistema) == "pozos barrera hidraulica":
            area = "POZOS"
        elif norm(area) == "aduccion":
            area = "IMPULSIÓN ADUCCIÓN"

        cam = re.sub(r"\D", "", str(f[25] or ""))
        items.append({
            "id": str(f[24]).strip() if f[24] else str(f[0]),
            "sub": str(f[8]).strip() if f[8] else "",
            "area": area,
            "disc": disciplina(f[27]),
            "prio": prioridad(f[20]),
            "estado": st,
            "cam": int(cam) if cam and int(cam) in (1, 2, 3) else 0,
            "diasAtraso": int(f[34]) if isinstance(f[34], (int, float)) and st == ATRASADO else None,
            "sinCompromiso": st in (ABIERTO, ATRASADO) and comp is None,
        })

    if incoherentes:
        c = Counter(m for m, _ in incoherentes)
        for m, n in c.items():
            avisos.append(f"{n} DT con estado inconsistente: {m}")
    return items, incoherentes


# =============================================================================
# 3) Agregaciones
# =============================================================================
def resumir(items):
    ab = [i for i in items if i["estado"] in (ABIERTO, ATRASADO)]
    atr = [i for i in ab if i["estado"] == ATRASADO]
    return {
        "total": len(items),
        "cerrados": sum(1 for i in items if i["estado"] == CERRADO),
        "abiertos": len(ab),
        "atrasados": len(atr),
        "enPlazo": len(ab) - len(atr),
        "sinCompromiso": sum(1 for i in ab if i["sinCompromiso"]),
    }


def construir(corte, subs, declarado, items, incoherentes, hoy):
    areas = sorted({s["area"] for s in subs})

    # ---------- caminatas ----------
    caminatas = {}
    for n in (1, 2):
        caminatas[str(n)] = {
            "total": len(subs),
            "realizada": sum(1 for s in subs if s["cam"][n] == "Realizada"),
            "programada": sum(1 for s in subs if s["cam"][n] == "Programada"),
            "pendiente": sum(1 for s in subs if s["cam"][n] == "Pendiente"),
        }
        c = caminatas[str(n)]
        c["pct"] = pct1(c["realizada"], c["total"])

    camArea = {}
    for a in areas:
        sa = [s for s in subs if s["area"] == a]
        camArea[a] = {"subs": len(sa)}
        for n in (1, 2):
            camArea[a][str(n)] = {
                "total": len(sa),
                "realizada": sum(1 for s in sa if s["cam"][n] == "Realizada"),
                "programada": sum(1 for s in sa if s["cam"][n] == "Programada"),
                "pendiente": sum(1 for s in sa if s["cam"][n] == "Pendiente"),
            }

    pendientes = sorted(
        [{"id": s["id"], "area": s["area"], "alcance": s["alcance"],
          "cam": n, "estado": s["cam"][n], "fecha": s["fechaCam"][n], "pec": s["pec"]}
         for s in subs for n in (2, 1) if s["cam"][n] != "Realizada"],
        key=lambda x: (x["cam"], x["estado"], x["area"], x["id"]))

    # ---------- carpetas (% PEC) ----------
    pecs = [s["pec"] for s in subs if isinstance(s["pec"], (int, float))]
    tramos = {}
    for i, (nom, piso) in enumerate(TRAMOS_PEC):
        techo = TRAMOS_PEC[i - 1][1] if i else 1.01
        tramos[nom] = sum(1 for p in pecs if piso <= p < techo)
    carpetas = {
        "tipo": "avance",
        "total": len(subs),
        "conDato": len(pecs),
        "promedio": pct1(sum(pecs), len(pecs)) if pecs else 0,
        "sobre95": sum(1 for p in pecs if p >= 0.95),
        "sobre80": sum(1 for p in pecs if p >= 0.80),
        "bajo80": sum(1 for p in pecs if p < 0.80),
        "tramos": tramos,
        "enRevisionCliente": (declarado.get("total") or {}).get("enRevision"),
    }
    carpetas["porArea"] = {}
    for a in areas:
        pa = [s["pec"] for s in subs if s["area"] == a and isinstance(s["pec"], (int, float))]
        carpetas["porArea"][a] = {
            "subs": sum(1 for s in subs if s["area"] == a),
            "promedio": pct1(sum(pa), len(pa)) if pa else 0,
            "sobre95": sum(1 for p in pa if p >= 0.95),
            "sobre80": sum(1 for p in pa if p >= 0.80),
            "bajo80": sum(1 for p in pa if p < 0.80),
        }
    carpetas["rezagados"] = sorted(
        [{"id": s["id"], "area": s["area"], "pec": round(100 * s["pec"], 1), "obs": s["obs"]}
         for s in subs if isinstance(s["pec"], (int, float)) and s["pec"] < 0.80],
        key=lambda x: x["pec"])

    # ---------- DT ----------
    prios = ["P1", "P2", "P3", "P4", "Sin asignar"]
    p1 = [i for i in items if i["prio"] == "P1"]
    areasDT = sorted({i["area"] for i in items})

    dt = {
        "global": resumir(items),
        "porPrioridad": {p: resumir([i for i in items if i["prio"] == p]) for p in prios},
        "porDisciplina": {d: resumir([i for i in items if i["disc"] == d])
                          for d in DISCIPLINAS if any(i["disc"] == d for i in items)},
        "p1PorDisciplina": {d: resumir([i for i in p1 if i["disc"] == d])
                            for d in DISCIPLINAS if any(i["disc"] == d for i in p1)},
        "p1PorArea": {a: resumir([i for i in p1 if i["area"] == a]) for a in areasDT},
        "porArea": {a: resumir([i for i in items if i["area"] == a]) for a in areasDT},
        "porCaminata": {},
        "prioridades": [p for p in prios if any(i["prio"] == p for i in items)],
    }
    for n in (1, 2, 3):
        sub = [i for i in items if i["cam"] == n]
        if sub:
            dt["porCaminata"][str(n)] = resumir(sub)

    # heatmap: abiertos (P1) por disciplina × caminata
    dt["heatmap"] = {
        "disc": [d for d in DISCIPLINAS if any(i["disc"] == d for i in p1)],
        "cam": [str(n) for n in (1, 2, 3) if any(i["cam"] == n for i in p1)],
    }
    dt["heatmap"]["vals"] = [
        [sum(1 for i in p1 if i["disc"] == d and i["cam"] == int(c)
             and i["estado"] in (ABIERTO, ATRASADO))
         for c in dt["heatmap"]["cam"]]
        for d in dt["heatmap"]["disc"]]

    # matriz disciplina × área de DT abiertos (alimenta el heatmap del panel)
    dt["matriz"] = {
        d: {a: sum(1 for i in items if i["disc"] == d and i["area"] == a
                   and i["estado"] in (ABIERTO, ATRASADO))
            for a in areasDT}
        for d in dt["porDisciplina"]}

    dias = sorted(i["diasAtraso"] for i in items if i["diasAtraso"])
    tr = Counter(">180 días" if d > 180 else "91-180 días" if d > 90
                 else "31-90 días" if d > 30 else "1-30 días" for d in dias)
    dt["antiguedad"] = {
        "tramos": {k: tr.get(k, 0) for k in
                   ["1-30 días", "31-90 días", "91-180 días", ">180 días"]},
        "max": dias[-1] if dias else 0,
        "mediana": dias[len(dias) // 2] if dias else 0,
    }

    porsub = defaultdict(list)
    for i in items:
        if i["estado"] in (ABIERTO, ATRASADO):
            porsub[i["sub"]].append(i)
    top = sorted(((s, resumir(v)) for s, v in porsub.items()),
                 key=lambda x: (-x[1]["atrasados"], -x[1]["abiertos"]))[:12]
    dt["topSubsistemas"] = [{"sub": s, **r} for s, r in top]

    # ---------- control de calidad del dato ----------
    control = {"contraste": [], "incoherentes": len(incoherentes)}
    tot = declarado.get("total") or {}
    if tot.get("subs") is not None:
        control["contraste"].append({
            "que": "Subsistemas", "declarado": str(tot["subs"]),
            "calculado": str(len(subs)), "ok": tot["subs"] == len(subs)})
    for n, key in ((1, "c1"), (2, "c2")):
        if tot.get(key) is not None:
            real = caminatas[str(n)]["realizada"]
            prog = caminatas[str(n)]["programada"]
            control["contraste"].append({
                "que": f"Caminata {n} — el RESUMEN publica «programadas»",
                "declarado": f"{tot[key]} programadas",
                "calculado": f"{real} realizadas + {prog} programadas",
                "ok": tot[key] <= real + prog,
                "nota": "el RESUMEN suma las agendadas de la semana; el panel las separa"})
    if tot.get("sobre95") is not None:
        control["contraste"].append({
            "que": "Carpetas sobre 95%", "declarado": str(tot["sobre95"]),
            "calculado": str(carpetas["sobre95"]), "ok": tot["sobre95"] == carpetas["sobre95"]})
    if tot.get("sobre80") is not None:
        ok = tot["sobre80"] == carpetas["sobre80"]
        control["contraste"].append({
            "que": "Carpetas sobre 80%", "declarado": str(tot["sobre80"]),
            "calculado": str(carpetas["sobre80"]), "ok": ok})
        if not ok:
            avisos.append(f"Carpetas sobre 80%: el RESUMEN declara {tot['sobre80']} "
                          f"y el detalle da {carpetas['sobre80']} — revisar")

    # DT abiertos: la hoja STATUS lleva su propio conteo por subsistema
    sP1 = sum(s["dtP1"] for s in subs)
    sP2 = sum(s["dtP2"] for s in subs)
    rP1 = dt["porPrioridad"]["P1"]["abiertos"]
    rP2 = dt["porPrioridad"]["P2"]["abiertos"]
    control["difDT"] = []
    for nom, sv, rv in (("P1", sP1, rP1), ("P2", sP2, rP2)):
        if sv != rv:
            control["difDT"].append({"prio": nom, "status": sv, "registro": rv})
    if control["difDT"]:
        avisos.append("La hoja STATUS y el registro de DT no cuadran en "
                      + ", ".join(d["prio"] for d in control["difDT"])
                      + " — el panel usa el registro de DT, que es la fuente ítem a ítem")

    ids_status = {s["id"] for s in subs}
    control["subsEnStatus"] = len(ids_status)
    control["areasStatus"] = len(areas)
    control["areasDT"] = len(areasDT)
    control["dtSinCompromiso"] = dt["global"]["sinCompromiso"]

    # ---------- semáforo por área ----------
    semaforo = {}
    for a in areas:
        c2 = camArea[a]["2"]
        r = dt["p1PorArea"].get(a) or resumir([])
        cp = carpetas["porArea"][a]
        avanceCam = pct1(c2["realizada"], c2["total"])
        cierre = pct1(r["cerrados"], r["total"]) if r["total"] else 100.0
        if avanceCam < 60 or cierre < 50 or r["atrasados"] > 20:
            est = ["crit", "Crítico"]
        elif avanceCam < 90 or cierre < 90 or r["atrasados"] > 5 or cp["bajo80"]:
            est = ["warn", "Atención"]
        else:
            est = ["good", "Al día"]
        semaforo[a] = est + [r]

    return {
        "meta": {
            **PROYECTO,
            "corte": corte.strftime("%Y-%m-%d") if corte else None,
            "corteTexto": corte.strftime("%d-%m-%Y") if corte else "sin fecha",
            "hoy": hoy.strftime("%d-%m-%Y"),
            "generado": datetime.now().strftime("%Y-%m-%d %H:%M"),
        },
        "subsistemas": {"total": len(subs), "porArea": {a: sum(1 for s in subs if s["area"] == a)
                                                        for a in areas}},
        "areas": areas,
        "caminatas": caminatas,
        "caminatasPorArea": camArea,
        "caminatasPendientes": pendientes,
        "carpetas": carpetas,
        "dt": dt,
        "semaforo": semaforo,
        "control": control,
        "disciplinas": DISCIPLINAS,
        "discCorta": DISC_CORTA,
    }


MARCA_INI = "<!-- === TALABRE:INICIO ==="
MARCA_FIN = "<!-- === TALABRE:FIN === -->"


def inyectar_en_html(ruta_html, datos):
    """Reemplaza el bloque de datos de Talabre dentro de index.html."""
    if not ruta_html.exists():
        avisos.append(f"No se encontró {ruta_html.name}; solo se escribió el JSON.")
        return False
    html = ruta_html.read_text(encoding="utf-8")
    i, f = html.find(MARCA_INI), html.find(MARCA_FIN)
    if i == -1 or f == -1 or f < i:
        avisos.append(f"No se encontraron las marcas TALABRE en {ruta_html.name}; "
                      f"el panel quedó con los datos anteriores.")
        return False
    bloque = (
        MARCA_INI + " (bloque generado por talabre.py — no editar a mano) -->\n"
        "<script>\nconst TAL = "
        + json.dumps(datos, ensure_ascii=False, indent=1)
        + ";\n</script>\n"
    )
    ruta_html.write_text(html[:i] + bloque + html[f:], encoding="utf-8")
    return True


# =============================================================================
def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--status", required=True, help="TalabreSTATUS_PEC.xlsx")
    ap.add_argument("--dt", required=True, help="TalabreCuadro_DT.xlsx")
    ap.add_argument("--hoy", help="Fecha de referencia AAAA-MM-DD (por defecto, hoy)")
    args = ap.parse_args()

    hoy = datetime.strptime(args.hoy, "%Y-%m-%d") if args.hoy else \
        datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

    corte, subs, declarado = leer_status(args.status)
    items, incoherentes = leer_dt(args.dt, hoy)
    datos = construir(corte, subs, declarado, items, incoherentes, hoy)

    destino = AQUI / "datos_talabre.json"
    destino.write_text(json.dumps(datos, ensure_ascii=False, indent=1), encoding="utf-8")
    inyectado = inyectar_en_html(AQUI / "index.html", datos)

    # ---------------- resumen en pantalla ----------------
    pc = lambda a, b: f"{pct1(a, b):.1f}%".replace(".", ",") if b else "—"
    m, S, C, CP, T = (datos["meta"], datos["subsistemas"], datos["caminatas"],
                      datos["carpetas"], datos["dt"])

    print("=" * 70)
    print(f"  {m['nombre'].upper()} — {m['descripcion']}")
    print(f"  {m['cliente']}   ·   contrato {m['contrato']}")
    print(f"  Última caminata agendada: {m['corteTexto']}   ·   atrasos al {m['hoy']}")
    print("=" * 70)

    print(f"\nSUBSISTEMAS: {S['total']} en {len(datos['areas'])} áreas")
    for a, n in sorted(S["porArea"].items(), key=lambda x: -x[1]):
        print(f"    {a:26s} {n:4d}")

    print("\nCAMINATAS")
    for n in ("1", "2"):
        c = C[n]
        print(f"  Caminata {n}: {c['realizada']}/{c['total']} realizadas ({pc(c['realizada'], c['total'])})"
              f" · {c['programada']} programadas · {c['pendiente']} pendientes")

    print(f"\nCARPETAS — avance PEC (Talabre no usa estados de aprobación)")
    print(f"  Promedio: {CP['promedio']}%".replace(".", ","))
    for k, v in CP["tramos"].items():
        print(f"    {k:12s} {v:4d}")
    if CP["enRevisionCliente"] is not None:
        print(f"  En revisión del cliente (según el archivo): {CP['enRevisionCliente']}")

    g = T["global"]
    print(f"\nDETALLES DE TERMINACIÓN — {g['total']} registros")
    print(f"  Cerrados  {g['cerrados']:5d}  ({pc(g['cerrados'], g['total'])})")
    print(f"  Abiertos  {g['abiertos']:5d}  →  atrasados {g['atrasados']} · "
          f"en plazo {g['enPlazo']} ({g['sinCompromiso']} sin fecha de compromiso)")
    print(f"\n  {'':12s} {'total':>6s} {'cerr.':>6s} {'abie.':>6s} {'atras.':>7s} {'% cierre':>9s}")
    for p in T["prioridades"]:
        r = T["porPrioridad"][p]
        print(f"  {p:12s} {r['total']:6d} {r['cerrados']:6d} {r['abiertos']:6d} "
              f"{r['atrasados']:7d} {pc(r['cerrados'], r['total']):>9s}")
    print("\n  P1 por disciplina:")
    for d, r in sorted(T["p1PorDisciplina"].items(), key=lambda x: -x[1]["atrasados"]):
        print(f"    {d:26s} {r['total']:4d} tot · {r['abiertos']:3d} abiertos · "
              f"{r['atrasados']:3d} atrasados")
    a = T["antiguedad"]
    if a["max"]:
        print(f"\n  Antigüedad de los atrasados: mediana {a['mediana']} días · máximo {a['max']} días")
        print("   ", " · ".join(f"{k}: {v}" for k, v in a["tramos"].items() if v))

    K = datos["control"]
    print("\nCONTROL DE CRUCE")
    for c in K["contraste"]:
        print(f"  {'OK ' if c['ok'] else '≠  '} {c['que']}: declara {c['declarado']} · "
              f"calculado {c['calculado']}")
    for d in K["difDT"]:
        print(f"  ⚠  DT {d['prio']} abiertos: hoja STATUS {d['status']} · "
              f"registro DT {d['registro']} (el panel usa el registro)")
    if K["incoherentes"]:
        print(f"  ⚠  {K['incoherentes']} registros con estado inconsistente respecto de sus fechas")
    else:
        print("  OK  Todos los DT respetan la regla: atrasado = abierto con compromiso vencido")

    if avisos:
        print("\n⚠  AVISOS — revisar antes de publicar:")
        for x in sorted(set(avisos))[:20]:
            print(f"   · {x}")

    print(f"\nEscritos: {destino.name}" + (", index.html (pestaña Talabre)" if inyectado else ""))


if __name__ == "__main__":
    main()
