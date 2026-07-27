#!/usr/bin/env python3
"""
Procesa el registro de No Conformidades y escribe datos_nc.json, que alimenta el
dashboard del módulo y la portada de la Suite QAQC.

Uso (cada corte, con el archivo del período):

    python3 no_conformidades.py --data /ruta/Data_NCR.xlsx

Opcional:
    --hoy 2026-07-27     fecha de referencia para antigüedad y atrasos
                         (por defecto, el día en que se corre el script)

------------------------------------------------------------------------------
DE DÓNDE SALE CADA DATO
------------------------------------------------------------------------------
Hoja «Observaciones», encabezado fila 1, datos fila 2+:

    col A  Nombre Del Proyecto      col D  Tipo De Emisión (interna/externa)
    col F  Tipo (NC / PNC / …)      col I  Especialidad
    col J  Persona Asignada         col L  Fecha De Creación
    col M  Título                   col T  Costo en UF
    col V  Fecha De Cierre          col Y  Estatus

------------------------------------------------------------------------------
REGLAS DE CLASIFICACIÓN
------------------------------------------------------------------------------
ORIGEN     «Interna BSMT» -> Interna · «Externa Cliente» y «Externa Subcontrato»
           -> Externa (se guarda además el detalle de cuál de las dos).

ESTADO     «Cerrado» -> Cerrada · el resto (Iniciado, Listo para revisión,
           No aceptado) -> Abierta. Se conserva el estatus original, porque
           «No aceptado» no es lo mismo que «Iniciado» para gestionar.

ATRASO     La regla del proyecto es: NC **abierta** cuya fecha de cierre
           comprometida ya venció. Está implementada, PERO en el archivo actual
           la columna «Fecha De Cierre» **solo se llena al cerrar la NC**: las
           30 abiertas no tienen fecha, así que el atraso no es calculable.
           El script lo detecta y lo informa, en vez de dar 0 atrasadas como si
           fuera un buen resultado. Mientras tanto se reporta la ANTIGÜEDAD de
           cada NC abierta (días desde que se creó), que sí es medible y sirve
           para priorizar.

TIEMPO DE  Solo para las cerradas: días entre la fecha de creación y la de
CIERRE     cierre. Es el indicador de qué tan rápido reacciona cada proyecto.
"""

import argparse
import json
import sys
import unicodedata
import warnings
from collections import Counter, defaultdict
from datetime import datetime
from math import floor
from pathlib import Path
from statistics import median

warnings.filterwarnings("ignore")

try:
    from openpyxl import load_workbook
except ImportError:
    sys.exit("Falta openpyxl.  Instálalo con:  pip install openpyxl")

AQUI = Path(__file__).resolve().parent

MODULO = {
    "id": "NC",
    "nombre": "No Conformidades",
    "descripcion": "Hallazgos de calidad y su corrección",
    "cliente": "Besalco Montajes",
}

# El registro nombra los proyectos con su código; la suite los conoce por su
# nombre corto. Oficina Central no es un proyecto de obra pero sí genera NC,
# así que se reporta aparte y no se mezcla con los tres proyectos.
PROYECTOS = {
    "P2416 - MLP - DESALADORA": {"id": "DESALADORA", "nombre": "Desaladora",
                                 "cliente": "Minera Los Pelambres", "obra": True},
    "P2407 - CODELCO - TALABRE": {"id": "TALABRE", "nombre": "Talabre",
                                  "cliente": "Codelco", "obra": True},
    "P2342 - ARQUEROS MASA": {"id": "ARQUEROS", "nombre": "Arqueros",
                              "cliente": "MASA", "obra": True},
    "OFICINA CENTRAL": {"id": "OFICINA", "nombre": "Oficina Central",
                        "cliente": "Besalco Montajes", "obra": False},
}
ORDEN = ["DESALADORA", "TALABRE", "ARQUEROS", "OFICINA"]

CERRADA, ABIERTA = "Cerrada", "Abierta"
TIPOS = ["No Conformidad", "Producto No Conforme", "Observación",
         "Opción de Mejora", "Reporte HSE Nivel Alto"]
TRAMOS = [("1-30 días", 0), ("31-90 días", 30), ("91-180 días", 90),
          ("181-365 días", 180), ("Más de un año", 365)]

avisos = []


def norm(v):
    if v is None:
        return ""
    s = str(v).strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFD", s)
                if unicodedata.category(c) != "Mn")
    return " ".join(s.split())


def pct1(a, b):
    """Porcentaje con 1 decimal, redondeando .5 hacia arriba (igual que el panel)."""
    return floor(1000.0 * a / b + 0.5) / 10.0 if b else 0.0


def texto(v, alt=""):
    return str(v).strip() if v not in (None, "") else alt


# =============================================================================
def leer(ruta, hoy):
    wb = load_workbook(ruta, data_only=True)
    hoja = "Observaciones" if "Observaciones" in wb.sheetnames else wb.sheetnames[0]
    it = wb[hoja].iter_rows(values_only=True)
    next(it)                                     # encabezado

    items, sin_proyecto = [], Counter()
    for f in it:
        if all(x is None for x in f):
            continue
        clave = texto(f[0]).upper()
        p = PROYECTOS.get(clave)
        if p is None:
            sin_proyecto[clave or "(vacío)"] += 1
            continue

        emision = texto(f[3], "Sin clasificar")
        if norm(emision).startswith("interna"):
            origen = "Interna"
        elif norm(emision).startswith("externa"):
            origen = "Externa"
        else:
            origen = "Sin clasificar"
            # Vacío es un dato faltante, no un valor raro: se cuenta aparte y
            # solo se avisa si aparece un texto que no se sabe clasificar.
            if f[3] not in (None, ""):
                avisos.append(f"Tipo de emisión no reconocido: «{emision}»")

        estatus = texto(f[24], "Sin estatus")
        estado = CERRADA if norm(estatus) == "cerrado" else ABIERTA

        creada = f[11] if isinstance(f[11], datetime) else None
        cerrada = f[21] if isinstance(f[21], datetime) else None

        tipo = texto(f[5], "Sin tipo")
        if tipo not in TIPOS:
            avisos.append(f"Tipo no reconocido: «{tipo}»")

        items.append({
            "proy": p["id"],
            "n": texto(f[1]),
            "origen": origen,
            "emision": emision,
            "codigoExterno": texto(f[4]),
            "tipo": tipo,
            "especialidad": texto(f[8], "Sin especialidad"),
            "responsable": texto(f[9], "Sin asignar"),
            "creada": creada,
            "titulo": texto(f[12])[:150],
            "costo": f[19] if isinstance(f[19], (int, float)) else None,
            "cerradaEl": cerrada,
            "estatus": estatus,
            "estado": estado,
            # Regla del proyecto: abierta y con la fecha comprometida ya vencida.
            "atrasada": estado == ABIERTA and bool(cerrada) and cerrada < hoy,
            "diasAbierta": (hoy - creada).days if estado == ABIERTA and creada else None,
            "diasCierre": (cerrada - creada).days if estado == CERRADA and cerrada and creada else None,
        })

    for k, n in sin_proyecto.items():
        avisos.append(f"{n} registros con proyecto no reconocido «{k}» — quedaron fuera")
    return items


# =============================================================================
def resumir(items):
    ab = [i for i in items if i["estado"] == ABIERTA]
    cer = [i for i in items if i["estado"] == CERRADA]
    dc = [i["diasCierre"] for i in cer if i["diasCierre"] is not None]
    costos = [i["costo"] for i in items if isinstance(i["costo"], (int, float))]
    return {
        "total": len(items),
        "cerradas": len(cer),
        "abiertas": len(ab),
        "atrasadas": sum(1 for i in ab if i["atrasada"]),
        "pctCierre": pct1(len(cer), len(items)),
        "internas": sum(1 for i in items if i["origen"] == "Interna"),
        "externas": sum(1 for i in items if i["origen"] == "Externa"),
        "costo": round(sum(costos), 1),
        "conCosto": len(costos),
        "medianaCierre": int(median(dc)) if dc else None,
        "maxCierre": max(dc) if dc else None,
    }


def tramo(d):
    t = TRAMOS[0][0]
    for nom, piso in TRAMOS:
        if d > piso:
            t = nom
    return t


def construir(items, hoy, fuente):
    proys = [p for p in ORDEN if any(i["proy"] == p for i in items)]
    meta_p = {v["id"]: v for v in PROYECTOS.values()}

    tipos = [t for t in TIPOS if any(i["tipo"] == t for i in items)]
    esps = [e for e, _ in Counter(i["especialidad"] for i in items).most_common()]

    def bloque(sub):
        ab = [i for i in sub if i["estado"] == ABIERTA]
        return {
            "resumen": resumir(sub),
            "porTipo": {t: resumir([i for i in sub if i["tipo"] == t]) for t in tipos
                        if any(i["tipo"] == t for i in sub)},
            "porOrigen": {o: resumir([i for i in sub if i["origen"] == o])
                          for o in ("Interna", "Externa", "Sin clasificar")
                          if any(i["origen"] == o for i in sub)},
            "porEmision": dict(Counter(i["emision"] for i in sub)),
            "porEspecialidad": {e: resumir([i for i in sub if i["especialidad"] == e])
                                for e in esps if any(i["especialidad"] == e for i in sub)},
            "porEstatus": dict(Counter(i["estatus"] for i in ab)),
            "antiguedad": {t: sum(1 for i in ab if i["diasAbierta"] is not None
                                  and tramo(i["diasAbierta"]) == t)
                           for t, _ in TRAMOS},
            "porAnio": dict(sorted(Counter(i["creada"].year for i in sub if i["creada"]).items())),
            "porMes": dict(sorted(Counter(i["creada"].strftime("%Y-%m")
                                          for i in sub if i["creada"]).items())),
        }

    datos = {
        "meta": {
            **MODULO,
            "hoy": hoy.strftime("%d-%m-%Y"),
            "corte": hoy.strftime("%Y-%m-%d"),
            "corteTexto": hoy.strftime("%d-%m-%Y"),
            "generado": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "fuente": fuente,
        },
        "orden": proys,
        "proyectos": {p: {**meta_p[p], **bloque([i for i in items if i["proy"] == p])}
                      for p in proys},
        "global": bloque(items),
        # Los tres proyectos de obra, sin Oficina Central: es lo comparable con
        # los otros módulos de la suite.
        "obra": bloque([i for i in items if meta_p[i["proy"]]["obra"]]),
        "tipos": tipos,
        "especialidades": esps,
    }

    # ---- detalle de las abiertas: es la lista accionable ----
    ab = sorted([i for i in items if i["estado"] == ABIERTA],
                key=lambda x: -(x["diasAbierta"] or 0))
    datos["abiertas"] = [{
        "proy": i["proy"], "proyNombre": meta_p[i["proy"]]["nombre"], "n": i["n"],
        "tipo": i["tipo"], "origen": i["origen"], "emision": i["emision"],
        "especialidad": i["especialidad"], "responsable": i["responsable"],
        "titulo": i["titulo"], "estatus": i["estatus"],
        "creada": i["creada"].strftime("%d-%m-%Y") if i["creada"] else None,
        "dias": i["diasAbierta"], "atrasada": i["atrasada"], "costo": i["costo"],
    } for i in ab]

    # ---- responsables con NC abiertas ----
    resp = Counter(i["responsable"] for i in ab)
    datos["responsables"] = [{"nombre": n, "abiertas": c,
                              "dias": max((x["diasAbierta"] or 0) for x in ab
                                          if x["responsable"] == n)}
                             for n, c in resp.most_common(12)]

    # ---- control de calidad del dato ----
    conFecha = sum(1 for i in items if i["estado"] == ABIERTA and i["cerradaEl"])
    nAb = datos["global"]["resumen"]["abiertas"]
    datos["control"] = {
        "registros": len(items),
        "abiertas": nAb,
        "abiertasConFechaComprometida": conFecha,
        # Si ninguna abierta trae fecha, el atraso no se puede calcular: hay que
        # decirlo, no reportar 0 atrasadas como si fuera un buen resultado.
        "atrasoCalculable": conFecha > 0,
        "rangoFechas": [min(i["creada"] for i in items if i["creada"]).strftime("%d-%m-%Y"),
                        max(i["creada"] for i in items if i["creada"]).strftime("%d-%m-%Y")],
        "sinCosto": sum(1 for i in items if i["costo"] is None),
        "sinEspecialidad": sum(1 for i in items if i["especialidad"] == "Sin especialidad"),
    }
    if not datos["control"]["atrasoCalculable"] and nAb:
        avisos.append(
            f"Las {nAb} NC abiertas no tienen fecha de cierre comprometida: la columna "
            f"«Fecha De Cierre» solo se llena al cerrar. El atraso NO es calculable con "
            f"este archivo; el panel reporta la antigüedad de cada NC abierta en su lugar.")

    # ---- semáforo por proyecto ----
    datos["semaforo"] = {}
    for p in proys:
        r = datos["proyectos"][p]["resumen"]
        if r["abiertas"] > 15 or r["pctCierre"] < 85:
            est = ["crit", "Crítico"]
        elif r["abiertas"] > 3 or r["pctCierre"] < 95:
            est = ["warn", "Atención"]
        else:
            est = ["good", "Al día"]
        datos["semaforo"][p] = est
    return datos


MARCA_INI = "<!-- === NC:INICIO ==="
MARCA_FIN = "<!-- === NC:FIN === -->"


def inyectar_en_html(ruta_html, datos):
    """Reemplaza el bloque de datos dentro de index.html."""
    if not ruta_html.exists():
        avisos.append(f"No se encontró {ruta_html.name}; solo se escribió el JSON.")
        return False
    html = ruta_html.read_text(encoding="utf-8")
    i, f = html.find(MARCA_INI), html.find(MARCA_FIN)
    if i == -1 or f == -1 or f < i:
        avisos.append(f"No se encontraron las marcas NC en {ruta_html.name}.")
        return False
    bloque = (MARCA_INI + " (bloque generado por no_conformidades.py — no editar a mano) -->\n"
              "<script>\nconst NC = " + json.dumps(datos, ensure_ascii=False, indent=1)
              + ";\n</script>\n")
    ruta_html.write_text(html[:i] + bloque + html[f:], encoding="utf-8")
    return True


# =============================================================================
def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--data", required=True, help="Data_NCR*.xlsx")
    ap.add_argument("--hoy", help="Fecha de referencia AAAA-MM-DD (por defecto, hoy)")
    args = ap.parse_args()

    hoy = datetime.strptime(args.hoy, "%Y-%m-%d") if args.hoy else \
        datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

    items = leer(args.data, hoy)
    if not items:
        sys.exit("No se leyó ningún registro. Revisa el archivo.")
    datos = construir(items, hoy, Path(args.data).name)

    destino = AQUI / "datos_nc.json"
    destino.write_text(json.dumps(datos, ensure_ascii=False, indent=1), encoding="utf-8")
    inyectado = inyectar_en_html(AQUI / "index.html", datos)

    # ---------------- resumen en pantalla ----------------
    pc = lambda a, b: f"{pct1(a, b):.1f}%".replace(".", ",") if b else "—"
    g = datos["global"]["resumen"]
    print("=" * 70)
    print(f"  NO CONFORMIDADES — {datos['meta']['fuente']}")
    print(f"  {g['total']} registros · referencia {datos['meta']['hoy']}")
    print("=" * 70)

    print(f"\n{'PROYECTO':17s} {'total':>6s} {'cerr':>5s} {'abier':>6s} {'%cierre':>8s} "
          f"{'inter':>6s} {'exter':>6s} {'UF':>8s} {'med.cierre':>11s}")
    for p in datos["orden"]:
        r = datos["proyectos"][p]["resumen"]
        n = datos["proyectos"][p]["nombre"]
        mc = f"{r['medianaCierre']} d" if r["medianaCierre"] is not None else "—"
        print(f"{n:17s} {r['total']:6d} {r['cerradas']:5d} {r['abiertas']:6d} "
              f"{pc(r['cerradas'], r['total']):>8s} {r['internas']:6d} {r['externas']:6d} "
              f"{r['costo']:8,.0f}".replace(",", ".") + f" {mc:>11s}")
    print(f"{'TOTAL':17s} {g['total']:6d} {g['cerradas']:5d} {g['abiertas']:6d} "
          f"{pc(g['cerradas'], g['total']):>8s} {g['internas']:6d} {g['externas']:6d} "
          f"{g['costo']:8,.0f}".replace(",", ".") +
          f" {(str(g['medianaCierre']) + ' d') if g['medianaCierre'] is not None else '—':>11s}")

    print("\nPOR TIPO")
    for t, r in datos["global"]["porTipo"].items():
        print(f"  {t:24s} {r['total']:4d} · abiertas {r['abiertas']:2d}")

    print("\nESPECIALIDADES con NC abiertas")
    for e, r in sorted(datos["global"]["porEspecialidad"].items(),
                       key=lambda x: -x[1]["abiertas"]):
        if r["abiertas"]:
            print(f"  {e:18s} {r['total']:4d} levantadas · {r['abiertas']:2d} abiertas")

    print("\nANTIGÜEDAD de las NC abiertas")
    for t, n in datos["global"]["antiguedad"].items():
        if n:
            print(f"  {t:15s} {n:3d}")

    K = datos["control"]
    print("\nCONTROL DE CALIDAD DEL DATO")
    print(f"  Registros procesados: {K['registros']} · rango {K['rangoFechas'][0]} → {K['rangoFechas'][1]}")
    print(f"  NC abiertas: {K['abiertas']} · con fecha de cierre comprometida: "
          f"{K['abiertasConFechaComprometida']}")
    if K["atrasoCalculable"]:
        print(f"  OK  El atraso es calculable: {g['atrasadas']} NC atrasadas")
    else:
        print("  ⚠  El ATRASO NO ES CALCULABLE con este archivo (ver avisos)")
    print(f"  Sin costo declarado: {K['sinCosto']} · sin especialidad: {K['sinEspecialidad']}")

    if avisos:
        print("\n⚠  AVISOS — revisar antes de publicar:")
        for a in sorted(set(avisos))[:15]:
            print(f"   · {a}")

    print(f"\nEscritos: {destino.name}" + (", index.html" if inyectado else ""))


if __name__ == "__main__":
    main()
