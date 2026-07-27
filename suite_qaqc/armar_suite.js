/*
 * Arma Suite_QAQC.html: un solo archivo con la portada ejecutiva y los módulos
 * dentro. Lee los módulos de modulos/, extrae el KPI de cada proyecto y rellena
 * plantilla_suite.html.
 *
 * Uso:  node armar_suite.js
 *
 * Por qué en Node y no con expresiones regulares: el dashboard de Protocolos
 * guarda sus datos en un literal JavaScript (claves sin comillas, comentarios,
 * comas finales). Se extrae el bloque por conteo de llaves y se evalúa, que es
 * la única forma robusta de leerlo sin reescribir el dashboard.
 */
const fs = require("fs");
const path = require("path");

const AQUI = __dirname;
const MODULOS = path.join(AQUI, "modulos");

// Los tres módulos cubren los MISMOS tres proyectos con distinto nombre.
// Esta tabla es la que permite cruzarlos en la portada ejecutiva.
const PROYECTOS = [
  { id: "DESALADORA", nombre: "Desaladora", cliente: "Antofagasta Minerals",
    protocolos: "P2416", qaqc: "DES" },
  { id: "TALABRE", nombre: "Talabre", cliente: "Codelco",
    protocolos: "P2407", qaqc: "TAL" },
  { id: "ARQUEROS", nombre: "Arqueros", cliente: "MASA",
    protocolos: "P2342", qaqc: "MASA" },
];

// ---------------------------------------------------------------- utilidades
const pct1 = (a, b) => (b ? Math.round((1000 * a) / b) / 10 : 0);

/** Extrae un literal `const NOMBRE = {...}` contando llaves (tolera strings). */
function literal(txt, nombre) {
  const ini = txt.indexOf(`const ${nombre} =`);
  if (ini === -1) return null;
  let i = txt.indexOf("{", ini), n = 0, com = null;
  for (let j = i; j < txt.length; j++) {
    const c = txt[j], prev = txt[j - 1];
    if (com) { if (c === com && prev !== "\\") com = null; continue; }
    if (c === "'" || c === '"' || c === "`") { com = c; continue; }
    if (c === "{") n++;
    else if (c === "}") { n--; if (n === 0) return txt.slice(i, j + 1); }
  }
  return null;
}

// ------------------------------------------------------- módulo · Protocolos
function leerProtocolos() {
  const f = path.join(MODULOS, "protocolos.html");
  if (!fs.existsSync(f)) return null;
  const lit = literal(fs.readFileSync(f, "utf8"), "PROJECTS");
  if (!lit) { console.log("Aviso: no se encontró PROJECTS en protocolos.html"); return null; }

  let P;
  try { P = eval("(" + lit + ")"); }
  catch (e) { console.log("Aviso: no se pudo leer PROJECTS —", e.message); return null; }

  // Suma recursiva de los estados de cada fila (algunas tienen hijos).
  const acum = (filas, t) => {
    for (const r of filas || []) {
      if (r.children) acum(r.children, t);
      else for (const k of ["S", "C", "P", "AP", "AE"]) t[k] += r[k] || 0;
    }
    return t;
  };

  const out = {};
  for (const [id, p] of Object.entries(P)) {
    if (p.status && p.status !== "active") continue;
    const t = acum(p.rows, { S: 0, C: 0, P: 0, AP: 0, AE: 0 });
    // KPI del dashboard de Protocolos: (AP + P) / (AP + P + AE + C) = % pendiente.
    // Es un indicador donde MENOS es mejor: 0% = cumple.
    const num = t.AP + t.P, den = t.AP + t.P + t.AE + t.C;
    out[id] = {
      nombre: p.name, cliente: p.client, actualizado: p.updated,
      estados: t, protocolos: den, pendientes: num, kpi: pct1(num, den),
    };
  }
  return out;
}

// ------------------------------------------------------ módulo · Cierre QAQC
function leerCierre() {
  const lee = (f) => {
    const r = path.join(AQUI, "..", "panel_control_TOP_P1", f);
    return fs.existsSync(r) ? JSON.parse(fs.readFileSync(r, "utf8")) : null;
  };
  const D = lee("estatus_datos.json"), X = lee("datos_desaladora.json"), Y = lee("datos_talabre.json");
  const out = {};
  if (D) {
    const p1 = D.dt.tipos.P1, c = D.cam.global, t = D.ctop.global;
    out.MASA = { corte: D.meta.corte_texto, subs: c.subs,
      cam: { realizada: c.c100, aplica: c.subs },
      p1: { total: p1.total, cerrados: p1.cerrados, abiertos: p1.abiertos, atrasados: p1.vencidos },
      det: { total: D.dt.global.total, cerrados: D.dt.global.cerrados,
             abiertos: D.dt.global.abiertos, atrasados: D.dt.global.vencidos },
      carpeta: { tipo: "estado", entregadas: t.entregadas, total: t.total } };
  }
  if (X) {
    let r = 0, a = 0;
    for (const v of Object.values(X.caminatasPorTipo)) {
      const d = v[v.vigente]; if (d) { r += d.realizada; a += d.aplica; }
    }
    const p1 = X.punch.porCategoria.P1;
    out.DES = { corte: X.meta.hoy, subs: X.subsistemas.total,
      cam: { realizada: r, aplica: a },
      p1: { total: p1.total, cerrados: p1.cerrados, abiertos: p1.abiertos, atrasados: p1.atrasados },
      det: X.punch.global,
      carpeta: { tipo: "estado", entregadas: X.carpetas.entregadas, total: X.carpetas.total } };
  }
  if (Y) {
    const p1 = Y.dt.porPrioridad.P1;
    out.TAL = { corte: Y.meta.hoy, subs: Y.subsistemas.total,
      cam: { realizada: Y.caminatas["2"].realizada, aplica: Y.caminatas["2"].total },
      p1: { total: p1.total, cerrados: p1.cerrados, abiertos: p1.abiertos, atrasados: p1.atrasados },
      det: Y.dt.global,
      carpeta: { tipo: "avance", promedio: Y.carpetas.promedio, total: Y.carpetas.total } };
  }
  return Object.keys(out).length ? out : null;
}

// ---------------------------------------------------------------- semáforos
// Cada módulo tiene su propia escala; no se mezclan.
const semProtocolos = (kpi) =>            // % pendiente: menos es mejor
  kpi === 0 ? ["good", "Cumple"] : kpi <= 3 ? ["good", "Aceptable"]
  : kpi <= 15 ? ["warn", "Medianamente"] : kpi <= 50 ? ["crit", "Deficiente"]
  : ["crit", "No cumple"];

const semCierre = (v) => {                // % cierre de P1: más es mejor
  const c = pct1(v.p1.cerrados, v.p1.total);
  if (c < 60 || v.p1.atrasados > 300) return ["crit", "Crítico"];
  if (c < 90 || v.p1.atrasados > 50) return ["warn", "Atención"];
  return ["good", "Al día"];
};

// ---------------------------------------------------------------------- main
const prot = leerProtocolos();
const cierre = leerCierre();

const filas = PROYECTOS.map((p) => {
  const a = prot && prot[p.protocolos];
  const b = cierre && cierre[p.qaqc];
  return {
    ...p,
    protocolos: a ? { kpi: a.kpi, protocolos: a.protocolos, pendientes: a.pendientes,
                      actualizado: a.actualizado, sem: semProtocolos(a.kpi) } : null,
    cierre: b ? { ...b, cierrePct: pct1(b.p1.cerrados, b.p1.total), sem: semCierre(b) } : null,
    noConformidades: null,     // módulo aún por construir
  };
});

const sum = (f) => filas.reduce((a, x) => a + (f(x) || 0), 0);
const datos = {
  generado: new Date().toISOString().slice(0, 16).replace("T", " "),
  proyectos: filas,
  modulos: {
    protocolos: prot ? {
      activo: true,
      proyectos: filas.filter((f) => f.protocolos).length,
      protocolos: sum((x) => x.protocolos && x.protocolos.protocolos),
      pendientes: sum((x) => x.protocolos && x.protocolos.pendientes),
      actualizado: (filas.find((f) => f.protocolos) || {}).protocolos?.actualizado || null,
    } : { activo: false },
    cierre: cierre ? {
      activo: true,
      proyectos: filas.filter((f) => f.cierre).length,
      subs: sum((x) => x.cierre && x.cierre.subs),
      detalles: sum((x) => x.cierre && x.cierre.det.total),
      cerrados: sum((x) => x.cierre && x.cierre.det.cerrados),
      abiertos: sum((x) => x.cierre && x.cierre.det.abiertos),
      atrasados: sum((x) => x.cierre && x.cierre.det.atrasados),
      p1Total: sum((x) => x.cierre && x.cierre.p1.total),
      p1Cerrados: sum((x) => x.cierre && x.cierre.p1.cerrados),
      p1Atrasados: sum((x) => x.cierre && x.cierre.p1.atrasados),
    } : { activo: false },
    noConformidades: { activo: false },
  },
};
datos.modulos.protocolos.kpi = pct1(datos.modulos.protocolos.pendientes,
                                    datos.modulos.protocolos.protocolos);

fs.writeFileSync(path.join(AQUI, "kpis_suite.json"), JSON.stringify(datos, null, 1), "utf8");

// ============================================================================
// Ensamblado del archivo único
// ============================================================================
function bloque(txt, ini, fin, contenido) {
  const i = txt.indexOf(ini), f = txt.indexOf(fin);
  if (i === -1 || f === -1 || f < i) {
    console.log(`Aviso: faltan las marcas ${ini} en la plantilla.`);
    return txt;
  }
  return txt.slice(0, i) + contenido + txt.slice(f);
}

// Dentro de un <script type="text/plain"> lo único que corta el bloque es la
// secuencia "</script". Se sustituye por un centinela y el shell la restaura al
// montar el módulo.
//
// OJO: el escape NO puede ser "<\/script". Los módulos ya traen esa secuencia
// dentro de sus propias cadenas JS, así que al restaurar se convertirían en
// "</script" reales y cerrarían el bloque antes de tiempo, rompiendo el módulo.
// Por eso se usa un centinela que se verifica ausente antes de empaquetar.
const CENTINELA = "@@CIERRE_SCRIPT_a7f3@@";
function empaquetar(html, archivo) {
  if (html.includes(CENTINELA))
    throw new Error(`${archivo} contiene el centinela ${CENTINELA}; cambiarlo en armar_suite.js`);
  return html.replace(/<\/script/gi, CENTINELA);
}

const plantilla = path.join(AQUI, "plantilla_suite.html");
if (!fs.existsSync(plantilla)) {
  console.log("\nNo se encontró plantilla_suite.html: solo se escribió kpis_suite.json.");
} else {
  let html = fs.readFileSync(plantilla, "utf8");

  html = bloque(html, "<!-- === KPIS:INICIO === -->", "<!-- === KPIS:FIN === -->",
    "<!-- === KPIS:INICIO === -->\n<script>const KPIS = " +
    JSON.stringify(datos) + ";<\/script>\n");

  const logo = path.join(AQUI, "besalco_logo_blanco.png");
  html = bloque(html, "<!-- === LOGO:INICIO === -->", "<!-- === LOGO:FIN === -->",
    "<!-- === LOGO:INICIO === -->\n<script>const LOGO = " +
    (fs.existsSync(logo)
      ? JSON.stringify("data:image/png;base64," + fs.readFileSync(logo).toString("base64"))
      : "null") + ";<\/script>\n");

  // Cada módulo puede venir de modulos/ o, si no está, de donde se genera.
  // Así el de Cierre QAQC no hay que copiarlo a mano en cada corte.
  const mods = [
    { clave: "protocolos", rutas: [path.join(MODULOS, "protocolos.html")] },
    { clave: "cierre", rutas: [path.join(MODULOS, "cierre_qaqc.html"),
                               path.join(AQUI, "..", "panel_control_TOP_P1", "index.html")] },
  ];
  let bloques = "<!-- === MODULOS:INICIO === -->\n", pesos = [];
  for (const { clave, rutas } of mods) {
    const r = rutas.find((x) => fs.existsSync(x));
    const hay = Boolean(r);
    const archivo = path.basename(r || rutas[0]);
    const cuerpo = hay ? empaquetar(fs.readFileSync(r, "utf8"), archivo) : "";
    if (hay) pesos.push(`${clave} ${(fs.statSync(r).size / 1024 / 1024).toFixed(1)} MB`);
    else console.log(`Aviso: falta ${archivo} — la suite sale sin el módulo ${clave}.`);
    bloques += `<script type="text/plain" id="mod-${clave}">${cuerpo}<\/script>\n`;
  }
  html = bloque(html, "<!-- === MODULOS:INICIO === -->", "<!-- === MODULOS:FIN === -->", bloques);

  const salida = path.join(AQUI, "Suite_QAQC.html");
  fs.writeFileSync(salida, html, "utf8");
  console.log(`\nEscrito: Suite_QAQC.html (${(fs.statSync(salida).size / 1024 / 1024).toFixed(1)} MB` +
              (pesos.length ? ` · ${pesos.join(" · ")}` : "") + ")");
}

// ------------------------------------------------------------------- consola
const nf = (n) => (n || 0).toLocaleString("es-CL");
console.log("=".repeat(72));
console.log("  SUITE QAQC — resumen ejecutivo");
console.log("=".repeat(72));
console.log(`\n${"PROYECTO".padEnd(14)}${"CLIENTE".padEnd(24)}${"PROTOCOLOS".padStart(16)}${"CIERRE QAQC".padStart(18)}`);
for (const f of filas) {
  const a = f.protocolos ? `${f.protocolos.kpi}% pend. ${f.protocolos.sem[1]}` : "—";
  const b = f.cierre ? `${f.cierre.cierrePct}% P1 ${f.cierre.sem[1]}` : "—";
  console.log(`${f.nombre.padEnd(14)}${f.cliente.padEnd(24)}${a.padStart(16)}${b.padStart(18)}`);
}
const M = datos.modulos;
console.log(`\nProtocolos   ${nf(M.protocolos.protocolos)} protocolos · ${nf(M.protocolos.pendientes)} pendientes · KPI ${M.protocolos.kpi}%`);
console.log(`Cierre QAQC  ${nf(M.cierre.subs)} subsistemas · ${nf(M.cierre.detalles)} detalles · ` +
            `${nf(M.cierre.cerrados)} cerrados · ${nf(M.cierre.atrasados)} atrasados`);
console.log(`No Conform.  módulo por construir`);
console.log(`\nEscrito: kpis_suite.json`);
