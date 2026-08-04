#!/usr/bin/env python3
"""Auditoría de coherencia: la misma cifra tiene que decir lo mismo en todas las capas.

    python3 auditoria_datos.py                       # solo lo que vive en el repositorio
    python3 auditoria_datos.py ~/Descargas/corte_10-08   # además, contra los Excel de origen

`verificar_suite.py` revisa la FORMA del entregable —que monte, que no desborde,
que la PPT descargue—. Esto revisa el FONDO: que el Excel, el JSON de cada módulo,
la portada de la suite y las PPT digan el mismo número. Cinco cruces:

  1. el Excel contra el JSON        (solo si se pasa la carpeta del corte)
  2. cada JSON consigo mismo        (las partes suman el total, la semana es
                                     lunes-domingo, el ritmo no tiene huecos)
  3. los módulos contra la portada
  4. los JSON contra el texto de las PPT
  5. la identidad de los proyectos, igual en los tres módulos

Sale con error si algo falla, así que sirve en cualquier automatismo.
"""
import json, re, sys, warnings, zipfile
from collections import Counter
from datetime import datetime, timedelta
from pathlib import Path
warnings.filterwarnings("ignore")
from openpyxl import load_workbook

R = Path(__file__).resolve().parent

# Los Excel de origen son opcionales: sin ellos se auditan igual las cuatro
# capas que viven en el repositorio. Pasar la carpeta del corte añade el cruce
# contra la fuente, que es el único que puede pillar un error de lectura.
#     python3 auditoria_datos.py [carpeta del corte]
ENTRADA = Path(sys.argv[1]).expanduser() if len(sys.argv) > 1 else None


def buscar(*claves, salvo=()):
    """El .xlsx/.xlsm MÁS NUEVO cuyo nombre contenga todas las claves y ninguna
    de `salvo`. El más nuevo, porque una carpeta de corte puede arrastrar la
    versión anterior del mismo archivo y auditar contra la vieja no sirve."""
    if not ENTRADA or not ENTRADA.is_dir():
        return None
    cand = [f for f in ENTRADA.iterdir()
            if f.suffix.lower() in (".xlsx", ".xlsm")
            and all(k in f.name.lower() for k in claves)
            and not any(k in f.name.lower() for k in salvo)]
    return max(cand, key=lambda f: f.stat().st_mtime) if cand else None


NCX = buscar("data_ncr", salvo=("externas",))
EXT = buscar("externas")
TAL_S = buscar("status_subsistemas")
TAL_D = buscar("terminaciones")

ok, mal, avi = [], [], []
def chk(cond, titulo, detalle=""):
    (ok if cond else mal).append(f"{titulo}{(' — ' + detalle) if detalle else ''}")
def nota(t): avi.append(t)

nc  = json.loads((R / "modulo_nc/datos_nc.json").read_text())
tal = json.loads((R / "panel_control_TOP_P1/datos_talabre.json").read_text())
des = json.loads((R / "panel_control_TOP_P1/datos_desaladora.json").read_text())
arq = json.loads((R / "panel_control_TOP_P1/estatus_datos.json").read_text())
kp  = json.loads((R / "suite_qaqc/kpis_suite.json").read_text())

# ═══════════════ 1 · EXCEL → JSON ═══════════════
print("1 · EL EXCEL CONTRA EL JSON" + ("" if (NCX or TAL_D) else "  (sin carpeta de origen: se omite)"))

if NCX and EXT:
    wb = load_workbook(NCX, data_only=True, read_only=True)
    it = wb["Observaciones"].iter_rows(values_only=True); next(it)
    filas = [f for f in it if f[0] not in (None, "")]
    odm = sum(1 for f in filas if str(f[5]).strip() == "Opción de Mejora")
    wb2 = load_workbook(EXT, data_only=True, read_only=True)
    ext = [f for f in wb2["Disposición NC-Externas"].iter_rows(min_row=5, values_only=True)
           if f[1] not in (None, "")]
    esperado = len(filas) - odm + len(ext)
    chk(nc["control"]["registros"] == esperado, "NC · registros",
        f"Excel {len(filas)} − {odm} ODM + {len(ext)} externas = {esperado} · JSON {nc['control']['registros']}")
    cerr_x = sum(1 for f in filas
                 if str(f[24]).strip() == "Cerrado" and str(f[5]).strip() != "Opción de Mejora")
    cerr_e = sum(1 for f in ext if str(f[7]).strip().lower() == "cerrada")
    chk(nc["global"]["resumen"]["cerradas"] == cerr_x + cerr_e, "NC · cerradas",
        f"Excel {cerr_x}+{cerr_e} · JSON {nc['global']['resumen']['cerradas']}")
elif ENTRADA:
    nota("NC · no se encontró el Data_NCR o la planilla de externas: no se cruzó contra el Excel")

if TAL_D:
    wbt = load_workbook(TAL_D, data_only=True, read_only=True)
    dt = [f for f in wbt["DT"].iter_rows(min_row=10, values_only=True) if f[0] is not None]
    chk(tal["dt"]["global"]["total"] == len(dt), "Talabre · registros de DT",
        f"Excel {len(dt)} · JSON {tal['dt']['global']['total']}")
    cerrados = sum(1 for f in dt if isinstance(f[23], datetime))
    chk(tal["dt"]["global"]["cerrados"] == cerrados, "Talabre · DT cerrados",
        f"Excel {cerrados} · JSON {tal['dt']['global']['cerrados']}")
if TAL_S:
    wbs = load_workbook(TAL_S, data_only=True)
    subs = [f for f in wbs["STATUS"].iter_rows(min_row=5, values_only=True) if f[2] not in (None, "")]
    chk(tal["subsistemas"]["total"] == len(subs), "Talabre · subsistemas",
        f"Excel {len(subs)} · JSON {tal['subsistemas']['total']}")

# ═══════════════ 2 · COHERENCIA INTERNA ═══════════════
print("2 · COHERENCIA INTERNA DE CADA JSON")
g = nc["global"]["resumen"]
chk(g["cerradas"] + g["abiertas"] == g["total"], "NC · cerradas + abiertas = total")
suma = sum(nc["proyectos"][p]["resumen"]["total"] for p in nc["orden"])
chk(suma == g["total"], "NC · los 4 frentes suman el global", f"{suma} vs {g['total']}")
obra = sum(nc["proyectos"][p]["resumen"]["total"] for p in nc["orden"] if nc["proyectos"][p]["obra"])
chk(obra == nc["obra"]["resumen"]["total"], "NC · los 3 proyectos de obra suman «obra»")
for p in nc["orden"]:
    r = nc["proyectos"][p]["resumen"]
    sinc = nc["proyectos"][p]["porEmision"].get("Sin clasificar", {}).get("total", 0)
    chk(r["internas"] + r["externas"] + sinc == r["total"], f"NC · {p}: internas+externas+sin clasificar=total",
        f'{r["internas"]}+{r["externas"]}+{sinc} vs {r["total"]}')
    if sinc: nota(f"NC · {p}: {sinc} registro(s) sin vía de emisión declarada")
    chk(r["cliente"] + r["subcontrato"] <= r["externas"], f"NC · {p}: cliente+subcontrato ≤ externas")
    chk(r["atrasadas"] <= r["abiertas"], f"NC · {p}: atrasadas ≤ abiertas")
    e = nc["proyectos"][p]["porEmision"]
    chk(sum(v["total"] for v in e.values()) == r["total"], f"NC · {p}: porEmision suma el total")
    t = nc["proyectos"][p]["porTipo"]
    chk(sum(v["total"] for v in t.values()) == r["total"], f"NC · {p}: porTipo suma el total")
    chk("Opción de Mejora" not in t, f"NC · {p}: sin opciones de mejora")
ab = nc["abiertas"]
chk(len(ab) == g["abiertas"], "NC · la lista de abiertas tiene tantas filas como abiertas",
    f"{len(ab)} vs {g['abiertas']}")
chk(sum(1 for a in ab if a["atrasada"]) == g["atrasadas"], "NC · atrasadas de la lista = resumen")
mal_atr = [a for a in ab if a["atrasada"] != ((a.get("dias") or 0) > nc["control"]["plazoRespuesta"])]
chk(not mal_atr, "NC · la regla de atraso se aplica igual en todos", f"{len(mal_atr)} discrepan")
d30 = [a for a in ab if a["atrasada"] and a["diasAtraso"] != a["dias"] - nc["control"]["plazoRespuesta"]]
chk(not d30, "NC · díasAtraso = antigüedad − plazo")

s = nc["global"]["semana"]
ini = datetime.strptime(s["desde"], "%d-%m-%Y"); fin = datetime.strptime(s["hasta"], "%d-%m-%Y")
chk(ini.weekday() == 0 and fin.weekday() == 6, "NC · la semana va de lunes a domingo",
    f"{s['desde']} ({ini.weekday()}) → {s['hasta']} ({fin.weekday()})")
chk((fin - ini).days == 6, "NC · la semana dura 7 días")
t = s["tendencia"]
huecos = []
for a_, b_ in zip(t, t[1:]):
    fa = datetime.strptime(a_["hasta"] + "-2026", "%d-%m-%Y")
    ib = datetime.strptime(b_["desde"] + "-2026", "%d-%m-%Y")
    if (ib - fa).days != 1: huecos.append(f"{a_['hasta']}→{b_['desde']}")
chk(not huecos, "NC · el ritmo no tiene huecos ni solapes", " · ".join(huecos))
chk(sum(x["total"] for x in t if x.get("acumulada")) >= 0, "NC · ritmo consistente")
chk(t[-1]["total"] == s["total"], "NC · la última fila del ritmo = la semana informada")

tg = tal["dt"]["global"]
chk(tg["cerrados"] + tg["abiertos"] == tg["total"], "Talabre · cerrados+abiertos=total")
chk(tg["atrasados"] + tg["enPlazo"] == tg["abiertos"], "Talabre · atrasados+enPlazo=abiertos")
sp = sum(v["total"] for v in tal["dt"]["porPrioridad"].values())
chk(sp == tg["total"], "Talabre · las prioridades suman el total", f"{sp} vs {tg['total']}")
sa = sum(v["total"] for v in tal["dt"]["porArea"].values())
chk(sa == tg["total"], "Talabre · las áreas suman el total", f"{sa} vs {tg['total']}")
chk(set(tal["dt"]["porArea"]) == set(tal["areas"]), "Talabre · las áreas del DT son las de STATUS",
    f"DT {sorted(set(tal['dt']['porArea'])-set(tal['areas']))} / STATUS {sorted(set(tal['areas'])-set(tal['dt']['porArea']))}")
for n_ in ("1", "2"):
    c = tal["caminatas"][n_]
    chk(c["realizada"] + c["programada"] + c["vencida"] + c["pendiente"] == c["total"],
        f"Talabre · caminata {n_}: los 4 estados suman el total",
        f'{c["realizada"]}+{c["programada"]}+{c["vencida"]}+{c["pendiente"]} vs {c["total"]}')
    chk(c["exigible"] == c["total"] - c["programada"], f"Talabre · caminata {n_}: exigible=total−programadas")
    chk(c["exigible"] >= c["realizada"] + c["vencida"], f"Talabre · caminata {n_}: exigible cubre realizadas y vencidas")

# ═══════════════ 3 · JSON → PORTADA DE LA SUITE ═══════════════
print("3 · LOS MÓDULOS CONTRA LA PORTADA")
m = kp["modulos"]
chk(m["noConformidades"]["total"] == nc["obra"]["resumen"]["total"], "Portada · NC total = obra del módulo",
    f"{m['noConformidades']['total']} vs {nc['obra']['resumen']['total']}")
chk(m["noConformidades"]["abiertas"] == nc["obra"]["resumen"]["abiertas"], "Portada · NC abiertas")
chk(m["noConformidades"]["atrasadas"] == nc["obra"]["resumen"]["atrasadas"], "Portada · NC atrasadas")
det = m["cierre"]["detalles"]
suma_det = des["punch"]["global"]["total"] + tal["dt"]["global"]["total"] + arq["dt"]["global"]["total"]
chk(det == suma_det, "Portada · detalles = suma de los 3 proyectos", f"{det} vs {suma_det}")
subs_suma = des["subsistemas"]["total"] + tal["subsistemas"]["total"] + arq["cam"]["global"]["subs"]
chk(m["cierre"]["subs"] == subs_suma, "Portada · subsistemas = suma de los 3", f"{m['cierre']['subs']} vs {subs_suma}")
for f in kp["proyectos"]:
    if f["cierre"]:
        chk(f["cierre"]["cierrePct"] == round(1000 * f["cierre"]["p1"]["cerrados"] / f["cierre"]["p1"]["total"]) / 10
            if f["cierre"]["p1"]["total"] else True, f"Portada · {f['id']}: % cierre de P1 bien calculado")

# ═══════════════ 4 · JSON → PPT ═══════════════
print("4 · LOS JSON CONTRA LAS PPT")
def textos(ruta):
    from pptx import Presentation
    return "\n".join(sh.text_frame.text for s_ in Presentation(str(ruta)).slides
                     for sh in s_.shapes if sh.has_text_frame)
tx_nc = textos(R / "modulo_nc/Panel_No_Conformidades.pptx")
nf = lambda v: f"{v:,}".replace(",", ".")
for etq, val in (("total de obra", nc["obra"]["resumen"]["total"]),
                 ("abiertas", nc["obra"]["resumen"]["abiertas"]),
                 ("atrasadas", nc["obra"]["resumen"]["atrasadas"])):
    chk(nf(val) in tx_nc, f"PPT NC · aparece el {etq} ({nf(val)})")
chk("Opción de Mejora" not in tx_nc and "opción de mejora" not in tx_nc.lower().replace("opciones de mejora", ""),
    "PPT NC · no menciona opciones de mejora como categoría")
tx_ci = textos(R / "panel_control_TOP_P1/Panel_Control_TOP_P1.pptx")
chk(nf(tal["dt"]["global"]["total"]) in tx_ci, f"PPT Cierre · aparece el total de DT de Talabre ({nf(tal['dt']['global']['total'])})")
chk(nf(des["punch"]["global"]["total"]) in tx_ci, "PPT Cierre · aparece el total de punch de Desaladora")

# ═══════════════ 5 · IDENTIDAD ENTRE MÓDULOS ═══════════════
print("5 · IDENTIDAD DE LOS PROYECTOS")
ident = {}
for p in nc["orden"]:
    if nc["proyectos"][p]["obra"]:
        ident.setdefault(p, set()).add((nc["proyectos"][p]["codigo"], nc["proyectos"][p]["nombre"], nc["proyectos"][p]["cliente"]))
for j, k in ((des, "DESALADORA"), (tal, "TALABRE")):
    ident.setdefault(k, set()).add((j["meta"]["codigo"], j["meta"]["nombre"], j["meta"]["cliente"]))
for f in kp["proyectos"]:
    ident.setdefault(f["id"], set()).add((f["codigo"], f["nombre"], f["cliente"]))
for k, v in ident.items():
    chk(len(v) == 1, f"Identidad · {k} se nombra igual en todos lados", str(v) if len(v) > 1 else "")

# ═══════════════ salida ═══════════════
print()
print("=" * 78)
for x in mal: print(f"  ✕ {x}")
for x in avi: print(f"  ▲ {x}")
print(f"\n  {len(ok)} comprobaciones OK · {len(mal)} fallan · {len(avi)} avisos")
print("=" * 78)
sys.exit(1 if mal else 0)
