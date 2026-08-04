#!/usr/bin/env python3
"""
Centro de Carga — actualizar la Suite QAQC arrastrando los archivos, sin terminal.

    python3 centro_carga.py          (o doble clic en el lanzador)

Abre una página local donde se sueltan los archivos del corte. La página muestra
qué llegó y qué falta, corre la actualización completa y deja la Suite lista para
descargar.

------------------------------------------------------------------------------
POR QUÉ ESTÁ SEPARADO DEL ENTREGABLE
------------------------------------------------------------------------------
`suite_qaqc/Suite_QAQC.html` se envía al cliente: no lleva —ni debe llevar— una
consola de actualización. Esta página es una herramienta local; nunca se empaqueta
dentro de la suite.

------------------------------------------------------------------------------
POR QUÉ HAY UN SERVIDOR Y NO SE HACE TODO EN EL NAVEGADOR
------------------------------------------------------------------------------
Las reglas de negocio —los 10 días de plazo de una NC, la homologación de
caminatas, qué prioridades entran al cierre, los semáforos— viven en los cuatro
scripts de Python y están validadas con los jefes de calidad. Reescribirlas en
JavaScript para que el navegador calcule solo dejaría **dos implementaciones de
las mismas reglas**, que es exactamente la forma en que una cifra empieza a
divergir sin que nadie se entere.

Así que el navegador no calcula nada: recibe los archivos, se los pasa a
`actualizar_semana.py` —el mismo comando de siempre, sin atajos— y muestra su
salida. Si el pipeline cambia, esta página hereda el cambio sin tocarse.

------------------------------------------------------------------------------
ALCANCE
------------------------------------------------------------------------------
El servidor escucha **solo en 127.0.0.1**: no queda accesible desde la red. Los
archivos subidos van a una carpeta temporal que se borra al cerrar, y los nombres
se reducen a su parte final, así que nada puede escribirse fuera de ahí.
"""

import http.server
import json
import shutil
import socket
import socketserver
import subprocess
import sys
import tempfile
import threading
import urllib.parse
import webbrowser
from datetime import datetime
from pathlib import Path

RAIZ = Path(__file__).resolve().parent
PAGINA = RAIZ / "centro_carga.html"
SUITE = RAIZ / "suite_qaqc" / "Suite_QAQC.html"
PUERTO_INICIAL = 8731

sys.path.insert(0, str(RAIZ))
from actualizar_semana import ROLES, corte_publicado, explorar, revisar  # noqa: E402


class Corrida:
    """Una ejecución del pipeline: su log, en vivo, y cómo terminó."""

    def __init__(self):
        self.lineas = []
        self.viva = False
        self.ok = None
        self._cerrojo = threading.Lock()

    def escribir(self, texto):
        with self._cerrojo:
            self.lineas.append(texto)

    def desde(self, n):
        with self._cerrojo:
            return list(self.lineas[n:]), len(self.lineas)

    def lanzar(self, carpeta, corte, igual):
        if self.viva:
            return
        self.lineas, self.viva, self.ok = [], True, None
        cmd = [sys.executable, "-u", "actualizar_semana.py", "--entrada", str(carpeta)]
        if corte:
            cmd += ["--corte", corte]
        if igual:
            cmd.append("--igual")

        def correr():
            try:
                p = subprocess.Popen(cmd, cwd=RAIZ, stdout=subprocess.PIPE,
                                     stderr=subprocess.STDOUT, text=True, bufsize=1)
                for linea in p.stdout:
                    self.escribir(linea.rstrip("\n"))
                self.ok = p.wait() == 0
            except Exception as e:
                self.escribir(f"Falló el arranque: {e}")
                self.ok = False
            finally:
                self.viva = False

        threading.Thread(target=correr, daemon=True).start()


ESTADO = {"carpeta": None, "corrida": Corrida()}


def carpeta_temporal():
    if ESTADO["carpeta"] is None:
        ESTADO["carpeta"] = Path(tempfile.mkdtemp(prefix="corte_qaqc_"))
    return ESTADO["carpeta"]


def reconocer():
    """La MISMA clasificación que usa la terminal: se importa, no se reescribe."""
    carpeta = carpeta_temporal()
    hallados, descartes, sin_rol = explorar(carpeta)
    publicado = corte_publicado()

    archivos = [{
        "rol": rol,
        "etiqueta": ROLES[rol],
        "nombre": ruta.name,
        "mb": round(ruta.stat().st_size / 1e6, 1),
        "fecha": fecha.strftime("%d-%m-%Y") if isinstance(fecha, datetime) else None,
    } for rol, (ruta, fecha) in sorted(hallados.items())]

    # Los avisos NO se recalculan aquí: son los mismos que detienen la corrida
    # por terminal, importados de `actualizar_semana.revisar`. Lo único que
    # cambia es el momento — se ven antes de procesar, que es cuando todavía se
    # puede pedir el archivo que falta.
    alertas = revisar(hallados, publicado)

    faltan = [ROLES[r] for r in ROLES if r not in hallados]
    return {
        "archivos": archivos,
        "alertas": alertas,
        "faltan": faltan,
        "sinRol": [r.name for r in sin_rol],
        "descartes": [{"etiqueta": ROLES[rol], "nombre": ruta.name} for rol, ruta, _ in descartes],
        "hoy": datetime.now().strftime("%Y-%m-%d"),
    }


class Manejador(http.server.BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def log_message(self, *a):          # la consola es para el usuario, no para el log
        pass

    # ---------------------------------------------------------------- utilidades
    def _responder(self, codigo, cuerpo, tipo="application/json; charset=utf-8",
                   extra=None):
        if isinstance(cuerpo, (dict, list)):
            cuerpo = json.dumps(cuerpo, ensure_ascii=False).encode("utf-8")
        elif isinstance(cuerpo, str):
            cuerpo = cuerpo.encode("utf-8")
        self.send_response(codigo)
        self.send_header("Content-Type", tipo)
        self.send_header("Content-Length", str(len(cuerpo)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(cuerpo)

    def _cuerpo(self):
        n = int(self.headers.get("Content-Length") or 0)
        return self.rfile.read(n) if n else b""

    # ------------------------------------------------------------------- rutas
    def do_GET(self):
        ruta = urllib.parse.urlparse(self.path)
        q = urllib.parse.parse_qs(ruta.query)

        if ruta.path in ("/", "/index.html"):
            if not PAGINA.exists():
                return self._responder(500, "Falta centro_carga.html", "text/plain; charset=utf-8")
            return self._responder(200, PAGINA.read_text(encoding="utf-8"),
                                   "text/html; charset=utf-8")

        if ruta.path == "/log":
            desde = int((q.get("desde") or ["0"])[0])
            lineas, total = ESTADO["corrida"].desde(desde)
            return self._responder(200, {"lineas": lineas, "total": total,
                                         "viva": ESTADO["corrida"].viva,
                                         "ok": ESTADO["corrida"].ok})

        if ruta.path == "/suite":
            if not SUITE.exists():
                return self._responder(404, {"error": "todavía no hay suite"})
            datos = SUITE.read_bytes()
            marca = datetime.now().strftime("%Y-%m-%d")
            return self._responder(
                200, datos, "text/html; charset=utf-8",
                {"Content-Disposition": f'attachment; filename="Suite_QAQC_{marca}.html"'})

        return self._responder(404, {"error": "no existe"})

    def do_POST(self):
        ruta = urllib.parse.urlparse(self.path)
        q = urllib.parse.parse_qs(ruta.query)

        if ruta.path == "/subir":
            nombre = urllib.parse.unquote((q.get("nombre") or [""])[0])
            # Solo la parte final del nombre: nada puede escribirse fuera de la
            # carpeta temporal por más que venga con «../» o una ruta entera.
            nombre = Path(nombre.replace("\\", "/")).name
            if not nombre:
                return self._responder(400, {"error": "sin nombre de archivo"})
            destino = carpeta_temporal() / nombre
            destino.write_bytes(self._cuerpo())
            return self._responder(200, {"ok": True, "nombre": nombre})

        if ruta.path == "/reconocer":
            try:
                return self._responder(200, reconocer())
            except Exception as e:
                return self._responder(500, {"error": f"{type(e).__name__}: {e}"})

        if ruta.path == "/correr":
            if ESTADO["corrida"].viva:
                return self._responder(409, {"error": "ya hay una actualización en curso"})
            cfg = json.loads(self._cuerpo() or b"{}")
            ESTADO["corrida"].lanzar(carpeta_temporal(), cfg.get("corte"),
                                     bool(cfg.get("igual")))
            return self._responder(200, {"ok": True})

        if ruta.path == "/limpiar":
            if ESTADO["carpeta"] and ESTADO["carpeta"].exists():
                shutil.rmtree(ESTADO["carpeta"], ignore_errors=True)
            ESTADO["carpeta"] = None
            return self._responder(200, {"ok": True})

        return self._responder(404, {"error": "no existe"})


class Servidor(socketserver.ThreadingTCPServer):
    allow_reuse_address = True
    daemon_threads = True


def puerto_libre(inicio):
    for p in range(inicio, inicio + 20):
        with socket.socket() as s:
            if s.connect_ex(("127.0.0.1", p)) != 0:
                return p
    sys.exit("No hay puertos libres entre "
             f"{inicio} y {inicio + 19}. Cierra otro Centro de Carga y reintenta.")


def main():
    if not PAGINA.exists():
        sys.exit(f"Falta {PAGINA.name} junto a este script.")
    puerto = puerto_libre(PUERTO_INICIAL)
    url = f"http://127.0.0.1:{puerto}/"

    print("=" * 70)
    print("  CENTRO DE CARGA — Suite QAQC")
    print("=" * 70)
    print(f"\n  Abierto en:  {url}")
    print("  Arrastra ahí los archivos del corte.")
    print("\n  Para cerrar: Ctrl+C en esta ventana.\n")

    with Servidor(("127.0.0.1", puerto), Manejador) as httpd:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\n  Cerrado.")
        finally:
            if ESTADO["carpeta"] and ESTADO["carpeta"].exists():
                shutil.rmtree(ESTADO["carpeta"], ignore_errors=True)


if __name__ == "__main__":
    main()
