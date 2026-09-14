/*
 * Genera la PPT ejecutiva de Protocolos desde `modulos/protocolos.html` y la
 * embebe en el propio módulo, igual que hacen los otros dos.
 *
 * Por qué existe
 * --------------
 * Protocolos llega armado de otro equipo y traía su propio camino de descarga:
 * un modal donde se elegía alcance y formato, y la PPT se generaba EN EL
 * NAVEGADOR al apretar el botón. Eso hacía dos cosas incómodas:
 *
 *   · el mazo salía distinto del de Cierre QAQC y No Conformidades —otra
 *     portada, otra tipografía, sin numeración ni etiqueta de proyecto—, y los
 *     tres viajan juntos en la misma suite;
 *   · lo que descargaba el usuario NO era lo que `verificar_suite.py` revisa,
 *     porque no existía hasta que alguien apretaba el botón. Los otros dos se
 *     generan aquí, se verifican y recién ahí se embeben.
 *
 * Este script iguala las dos cosas: arma el mazo con el mismo formato y lo deja
 * embebido, de modo que el botón descargue de un clic el archivo ya revisado.
 * Se corre DESPUÉS de copiar el dashboard del corte a `modulos/protocolos.html`
 * y ANTES de `armar_suite.js`.
 *
 * Uso:  node gen_ppt_protocolos.js
 */
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");

const AQUI = __dirname;
const MOD = path.join(AQUI, "modulos", "protocolos.html");
const SALIDA = path.join(AQUI, "Informe_Protocolos.pptx");

// El bloque que este script inyecta en el módulo. Se reconoce por el centinela
// para poder reemplazarlo en cada corte sin acumular copias.
const MARCA_INI = "<!-- PPT-PROTOCOLOS:INICIO (lo escribe gen_ppt_protocolos.js — no editar) -->";
const MARCA_FIN = "<!-- PPT-PROTOCOLOS:FIN -->";

if (!fs.existsSync(MOD)) {
  console.log("No está modulos/protocolos.html — nada que hacer.");
  process.exit(0);
}
const HTML0 = fs.readFileSync(MOD, "utf8");

// ---------- formato (mismo que los otros dos mazos) ----------
const nf = (n) => (n || 0).toLocaleString("es-CL");
const pct1 = (a, b) => (b ? Math.round((1000 * a) / b) / 10 : 0);
const pc = (v) => Number(v).toLocaleString("es-CL") + "%";
const corta = (t, n) => {
  const v = String(t || "").trim();
  return v.length > n ? v.slice(0, n - 1) + "…" : v;
};

const C = {
  navy: "12233B", navy2: "1B3252", steel: "6E8CA8", ice: "DCE6F0",
  paper: "F4F6F9", white: "FFFFFF", ink: "1A2433", ink2: "56606E",
  copper: "C87A32", copperL: "E3A55B",
  good: "2E9E5B", goodD: "1F7A44", warn: "E0A32E", crit: "C0392B",
  blue: "2E6FB5", blueL: "9DBDE0", track: "D9DFE6",
  neutral: "5B6B7C", blueM: "3A6FA8", warnD: "9A6B0F", copperD: "8A5520",
};
const FT = "Calibri", FH = "Cambria";
const z = (v) => (v > 0 ? v : null);
const ETIQ = { showValue: true, dataLabelPosition: "ctr", dataLabelFontFace: FT,
  dataLabelFontSize: 11, dataLabelColor: "FFFFFF", dataLabelFontBold: true };
const BARRAS = {
  barDir: "bar", barGrouping: "stacked", ...ETIQ,
  titleFontFace: FT, titleFontSize: 13, titleColor: C.ink, showTitle: true,
  showLegend: true, legendPos: "b", legendFontFace: FT, legendFontSize: 11, legendColor: C.ink2,
  catAxisLabelFontFace: FT, catAxisLabelColor: C.ink, catAxisLabelFontBold: true,
  valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
  chartArea: { fill: { color: C.white } },
};
// Escala del KPI — las mismas cinco bandas y los mismos rótulos que muestra el
// dashboard en su leyenda. El KPI es el % PENDIENTE sobre lo exigible: menos es
// mejor, al revés que un avance.
//
// Los colores no son los del panel sino los del mazo: el verde y el naranja de
// pantalla dan menos de 4,5:1 con texto blanco encima de un chip, y el número
// queda ilegible impreso. Las bandas y los nombres sí son idénticos.
const ESCALA = [
  { hasta: 0,   txt: "Cumple",       rango: "0%",       c: C.goodD },
  { hasta: 3,   txt: "Aceptable",    rango: "1%–3%",    c: C.blue },
  { hasta: 15,  txt: "Medianamente", rango: "4%–15%",   c: C.warnD },
  { hasta: 50,  txt: "Deficiente",   rango: "16%–50%",  c: C.copperD },
  { hasta: 100, txt: "No Cumple",    rango: "51%–100%", c: C.crit },
];
const SEM = (kpi) => {
  const b = ESCALA.find((x) => kpi <= x.hasta) || ESCALA[ESCALA.length - 1];
  return [b.c, b.txt];
};

// Glosario de estados, con los códigos que usa el propio dashboard. Va impreso
// en el mazo: quien lo recibe no siempre tiene el panel al lado.
const ESTADOS = [
  ["(P)",  "Pendiente",         "Actividad ejecutada o en ejecución sin protocolo."],
  ["(AP)", "Abierto Pendiente", "Protocolo abierto pendiente de cierre con actividad terminada."],
  ["(AE)", "Abierto Ejecución", "Actividad en ejecución con protocolo en proceso."],
  ["(C)",  "Cerrado",           "Actividad ejecutada con protocolización cerrada."],
  ["(S)",  "Sí Aplica",         "Actividad detectada y proyectada por protocolizar."],
  ["(RV)", "En Revisión",       "Protocolo en revisión del cliente (cuando aplique)."],
];

// Código corto de cada disciplina, el mismo que usaba el tablero anterior.
const COD_DISC = {
  "Topografía": "TOPO", "Obras Civiles": "OOCC", "Arquitectura": "ARQT",
  "Estructuras": "ESTR", "Mecánica": "MECA", "Piping": "PIPN",
  "Eléctrico": "ELEC", "Instrumentación": "INST", "Precomisionamiento": "PCOM",
};
const cod = (n) => COD_DISC[n] || String(n).slice(0, 4).toUpperCase();

// Las tres partes en que se reparte el universo, con los mismos nombres y
// colores del tablero: lo avanzado, lo que falta y lo que todavía no se exige.
const PARTES = (t) => [
  { n: "Avance (C+AE)",   v: t.C + t.AE,  c: C.blue },
  { n: "En Falta (AP+P)", v: t.AP + t.P,  c: C.crit },
  { n: "Remanente (S)",   v: t.S,         c: C.track },
];

/** Dona de avance con el universo escrito en el centro. */
function dona(s, t, x, y, d, leyenda) {
  const pr = PARTES(t).filter((q) => q.v > 0);
  s.addChart(p.ChartType.doughnut,
    [{ name: "Universo", labels: pr.map((q) => q.n), values: pr.map((q) => q.v) }],
    { x, y, w: d, h: d, chartColors: pr.map((q) => q.c), holeSize: 62,
      showLegend: false, showValue: false, dataBorder: { pt: 1, color: "FFFFFF" },
      chartArea: { fill: { color: C.white } } });
  s.addText(nf(t.universo), { x: x + d * 0.18, y: y + d * 0.42, w: d * 0.64, h: d * 0.2,
    fontFace: FH, fontSize: d > 2 ? 17 : 10, bold: true, color: C.ink,
    align: "center", margin: 0, valign: "middle" });
  // La leyenda se sitúa aparte: colgarla de la dona la hacía chocar con la
  // fila de disciplinas de más abajo.
  if (leyenda) {
    let ly = leyenda[1];
    PARTES(t).forEach((q) => {
      s.addShape(p.ShapeType.rect, { x: leyenda[0], y: ly + 0.05, w: 0.12, h: 0.12,
        fill: { color: q.c }, line: { color: C.steel, width: 0.5 } });
      s.addText(`${q.n}  ${pct1(q.v, t.universo)}%`, { x: leyenda[0] + 0.2, y: ly - 0.02,
        w: 1.9, h: 0.24, fontFace: FT, fontSize: 8.5, color: C.ink2, margin: 0, valign: "middle" });
      ly += 0.23;
    });
  }
}

/** Dibuja la escala del KPI como cinco chips seguidos. */
function escala(s, x0, y, ancho) {
  const w = ancho || 2.28;
  s.addText("ESCALA DEL KPI · % PENDIENTE, MENOS ES MEJOR", { x: x0, y: y - 0.28, w: 8, h: 0.24,
    fontFace: FT, fontSize: 8.5, bold: true, color: C.ink2, charSpacing: 1.2, margin: 0 });
  let x = x0;
  ESCALA.forEach((b) => {
    s.addShape(p.ShapeType.roundRect, { x, y, w: w - 0.12, h: 0.34, rectRadius: 0.16,
      fill: { color: b.c }, line: { type: "none" } });
    s.addText(`${b.txt}  ${b.rango}`, { x, y: y + 0.005, w: w - 0.12, h: 0.33, fontFace: FT,
      fontSize: 9, bold: true, color: C.white, align: "center", margin: 0, valign: "middle" });
    x += w;
  });
}

// ---------- lectura del módulo ----------
/** Extrae un literal `const NOMBRE = {...}` / `let NOMBRE = [...]` contando
 *  llaves. Es el mismo lector que usa `armar_suite.js`: si cambia el formato
 *  del dashboard, los dos fallan igual y se arregla en un solo sitio. */
function literal(txt, nombre, abre) {
  const cierra = abre === "[" ? "]" : "}";
  const re = new RegExp("(?:const|let|var)\\s+" + nombre + "\\s*=\\s*");
  const m = re.exec(txt);
  if (!m) return null;
  let i = txt.indexOf(abre, m.index), n = 0, com = null;
  for (let j = i; j < txt.length; j++) {
    const c = txt[j], prev = txt[j - 1];
    if (com) { if (c === com && prev !== "\\") com = null; continue; }
    if (c === "'" || c === '"' || c === "`") { com = c; continue; }
    if (c === abre) n++;
    else if (c === cierra) { n--; if (n === 0) return txt.slice(i, j + 1); }
  }
  return null;
}
const litP = literal(HTML0, "PROJECTS", "{");
if (!litP) { console.log("No se encontró PROJECTS en protocolos.html — no se genera PPT."); process.exit(0); }
const PROJECTS = eval("(" + litP + ")");
const litH = literal(HTML0, "KPI_HISTORY", "[");
const KPI_HISTORY = litH ? eval("(" + litH + ")") : [];

// Suma recursiva igual que el `sumCh()` del dashboard: un nodo con hijos no
// aporta sus propios estados, solo el `rv`.
const acum = (filas, t) => {
  for (const r of filas || []) {
    if (r.children) { acum(r.children, t); t.rv += r.rv || 0; }
    else for (const k of ["S", "C", "P", "AP", "AE", "rv"]) t[k] += r[k] || 0;
  }
  return t;
};
const cero = () => ({ S: 0, C: 0, P: 0, AP: 0, AE: 0, rv: 0 });
/** Universo = S+C+P+AP+AE+rv · En falta = AP+P · Base KPI = universo − S ·
 *  KPI = en falta / base. Mismas definiciones que el panel y la portada. */
function metricas(t) {
  const universo = t.S + t.C + t.P + t.AP + t.AE + t.rv;
  const enFalta = t.AP + t.P;
  const base = t.AP + t.P + t.AE + t.C;
  return { ...t, universo, enFalta, base, avance: t.C + t.AE, remanente: t.S,
           kpi: pct1(enFalta, base) };
}

const PIDS = Object.keys(PROJECTS).filter((k) => (PROJECTS[k].status || "active") === "active");
const PROY = PIDS.map((id) => {
  const p = PROJECTS[id];
  const m = metricas(acum(p.rows, cero()));
  const disc = (p.rows || []).map((r) => {
    const d = metricas(acum([r], cero()));
    return { nombre: r.name || r.id, ...d };
  }).filter((d) => d.universo > 0);
  return { id, nombre: (p.name || id).replace(/^P\d+\s*—\s*/, "").replace(/^Proyecto\s+/i, ""),
           etiqueta: p.name || id, cliente: p.client || "", corte: p.updated || "", ...m, disc };
});
const TOT = metricas(PROY.reduce((a, x) => {
  for (const k of ["S", "C", "P", "AP", "AE", "rv"]) a[k] += x[k];
  return a;
}, cero()));

// El corte del mazo es el MÁS RECIENTE de los tres, no el del primer proyecto
// de la lista: los proyectos no siempre llegan el mismo día y la portada de la
// suite ya declara el más nuevo más el aviso de cortes mixtos. Las fechas del
// dashboard vienen como «13-sep-2026», así que hay que ordenarlas de verdad.
const MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
const orden = (t) => {
  const m = /^(\d{1,2})-([a-zA-Z]{3})\w*-(\d{4})$/.exec(String(t || "").trim());
  if (!m) return "";
  const mes = MESES.indexOf(m[2].toLowerCase());
  return mes < 0 ? "" : `${m[3]}-${String(mes + 1).padStart(2, "0")}-${m[1].padStart(2, "0")}`;
};
const CORTES = [...new Set(PROY.map((x) => x.corte).filter(Boolean))]
  .sort((a, b) => orden(a).localeCompare(orden(b)));
const CORTE = CORTES.length ? CORTES[CORTES.length - 1] : "";
const CAP = `Fuente: dashboard de Protocolos · Corte ${CORTE}` +
  (CORTES.length > 1 ? ` · cortes mixtos: ${CORTES.join(" · ")}` : "") +
  ` · ${nf(TOT.universo)} protocolos`;

// ---------- lienzo ----------
const p = new pptxgen();
p.defineLayout({ name: "W", width: 13.333, height: 7.5 });
p.layout = "W";
p.author = "Gestión de Calidad";
p.title = "Control de Protocolos — Gestión de Calidad";
const bg = (s, color) => { s.background = { color }; };

const LOGO = [path.join(AQUI, "besalco_logo.png"),
              path.join(AQUI, "..", "panel_control_TOP_P1", "besalco_logo.png")]
  .find((r) => fs.existsSync(r));
if (!LOGO) console.log("Aviso: no se encontró besalco_logo.png — la PPT sale sin logo.");
function logo(s, portadaP) {
  if (!LOGO) return;
  const g = portadaP ? { x: 9.14, y: 0.46, w: 3.88, h: 0.84 }
                     : { x: 9.97, y: 0.24, w: 3.03, h: 0.65 };
  s.addImage({ path: LOGO, ...g });
}
let nSlide = 1;
function footer(s, dark, cap, proy) {
  s.addText(cap || CAP, { x: 0.5, y: 7.08, w: 10.6, h: 0.3, fontFace: FT, fontSize: 8.5,
    color: dark ? C.steel : C.ink2, align: "left", margin: 0 });
  if (proy) s.addText(`PROYECTO ${proy.toUpperCase()}`,
    { x: 2.0, y: 7.09, w: 10.67, h: 0.3, fontFace: FT, fontSize: 8.5,
      color: dark ? C.steel : C.ink2, align: "right", margin: 0 });
  s.addText(String(nSlide++), { x: 12.5, y: 7.08, w: 0.4, h: 0.3, fontFace: FT, fontSize: 9,
    color: dark ? C.steel : C.ink2, align: "right", margin: 0 });
}
const kicker = (s, txt, x, y, color) =>
  s.addText(txt.toUpperCase(), { x, y, w: 8, h: 0.28, fontFace: FT, fontSize: 11,
    bold: true, color: color || C.copper, charSpacing: 2, margin: 0, align: "left" });
const dot = (s, x, y, color, d) =>
  s.addShape(p.ShapeType.ellipse, { x, y, w: d || 0.14, h: d || 0.14,
    fill: { color }, line: { type: "none" } });
const titulo = (s, txt, color) =>
  s.addText(txt, { x: 0.5, y: 0.72, w: 12.3, h: 0.6, fontFace: FH, fontSize: 30,
    bold: true, color: color || C.ink, margin: 0 });
const bajada = (s, txt) =>
  s.addText(txt, { x: 0.5, y: 1.35, w: 12.3, h: 0.55, fontFace: FT, fontSize: 13.5,
    color: C.ink2, margin: 0, valign: "top" });

function portada(s, num, nombre, descripcion, detalle, kpis, pie, cap) {
  bg(s, C.navy);
  s.addShape(p.ShapeType.ellipse, { x: 11.0, y: -1.6, w: 4.2, h: 4.2, fill: { color: C.navy2 }, line: { type: "none" } });
  s.addShape(p.ShapeType.ellipse, { x: 12.1, y: 5.0, w: 3.0, h: 3.0, fill: { color: C.navy2 }, line: { type: "none" } });
  logo(s, true);
  dot(s, 0.55, 1.35, C.copper, 0.22);
  s.addText(num, { x: 0.85, y: 1.2, w: 10, h: 0.4, fontFace: FT, fontSize: 13,
    bold: true, color: C.copperL, charSpacing: 3, margin: 0 });
  s.addText(nombre, { x: 0.82, y: 1.72, w: 11.6, h: 1.1, fontFace: FH, fontSize: 46,
    bold: true, color: C.white, margin: 0 });
  s.addText(descripcion, { x: 0.85, y: 2.9, w: 10.5, h: 0.5, fontFace: FT, fontSize: 16,
    color: C.ice, margin: 0 });
  s.addText(detalle, { x: 0.85, y: 3.45, w: 9.6, h: 1.5, fontFace: FT, fontSize: 13.5,
    color: C.steel, lineSpacing: 20, margin: 0, valign: "top" });
  let kx = 0.85;
  kpis.forEach(([v, l]) => {
    s.addText(v, { x: kx, y: 5.3, w: 2.9, h: 0.7, fontFace: FH, fontSize: 34, bold: true, color: C.copperL, margin: 0 });
    s.addText(l, { x: kx, y: 6.0, w: 2.9, h: 0.55, fontFace: FT, fontSize: 11.5, color: C.ice, margin: 0 });
    kx += 3.0;
  });
  s.addText(pie, { x: 0.85, y: 6.75, w: 8, h: 0.3, fontFace: FT, fontSize: 11,
    italic: true, color: C.steel, margin: 0 });
  footer(s, true, cap);
}
function tarjetas(s, cards) {
  let cx = 0.5; const cw = 3.0, gp = 0.13;
  cards.forEach((cd) => {
    s.addShape(p.ShapeType.roundRect, { x: cx, y: 1.6, w: cw, h: 2.5, rectRadius: 0.08,
      fill: { color: C.white }, line: { color: "E4E9EF", width: 1 },
      shadow: { type: "outer", color: "9AA7B5", blur: 7, offset: 2, angle: 90, opacity: 0.28 } });
    dot(s, cx + 0.28, 1.88, cd.c, 0.16);
    s.addText(cd.p, { x: cx + 0.28, y: 2.12, w: cw - 0.5, h: 0.82, fontFace: FH, fontSize: 40, bold: true, color: cd.c, margin: 0 });
    s.addText(cd.v, { x: cx + 0.3, y: 2.95, w: cw - 0.5, h: 0.32, fontFace: FT, fontSize: 14, bold: true, color: C.ink, margin: 0 });
    s.addText(cd.t, { x: cx + 0.3, y: 3.28, w: cw - 0.55, h: 0.32, fontFace: FT, fontSize: 12, color: C.ink2, margin: 0 });
    s.addText(cd.d, { x: cx + 0.3, y: 3.60, w: cw - 0.55, h: 0.42, fontFace: FT, fontSize: 9.5, color: C.ink2, margin: 0, valign: "top" });
    cx += cw + gp;
  });
}
function tabla(s, cols, filas, y0, rh, opt) {
  const o = opt || {};
  const W = o.ancho || 12.3;
  cols.forEach((c) => s.addText(c.t, { x: c.x, y: y0, w: c.w, h: 0.3, fontFace: FT,
    fontSize: 9.5, bold: true, color: C.ink2, align: c.al || "right", charSpacing: 0.5, margin: 0 }));
  s.addShape(p.ShapeType.line, { x: cols[0].x, y: y0 + 0.34, w: W, h: 0, line: { color: C.steel, width: 1 } });
  let y = y0 + 0.48;
  filas.forEach((f, i) => {
    if (i % 2 === 1) s.addShape(p.ShapeType.rect, { x: cols[0].x - 0.05, y: y - 0.05, w: W + 0.05, h: rh,
      fill: { color: C.white }, line: { type: "none" } });
    f.forEach((celda, j) => {
      if (celda == null) return;
      const cc = typeof celda === "object" ? celda : { txt: celda };
      s.addText(String(cc.txt), { x: cols[j].x, y: y, w: cols[j].w, h: rh - 0.1,
        fontFace: FT, fontSize: cc.size || o.size || 11.5, bold: !!cc.bold,
        color: cc.color || C.ink, align: cols[j].al || "right", margin: 0, valign: "middle" });
    });
    s.addShape(p.ShapeType.line, { x: cols[0].x, y: y + rh - 0.06, w: W, h: 0, line: { color: "D9DFE6", width: 0.75 } });
    y += rh;
  });
  return y;
}

// =====================================================================
// 1 — Portada · consolidado
// =====================================================================
let s = p.addSlide();
const peor = PROY.slice().sort((a, b) => b.kpi - a.kpi)[0];
portada(s, "CONTROL DE PROTOCOLOS", "Gestión de Calidad",
  "Estado de protocolos de los tres proyectos",
  `Universo ${nf(TOT.universo)} = (S) ${nf(TOT.S)} + (AE) ${nf(TOT.AE)} + (AP) ${nf(TOT.AP)} + ` +
  `(P) ${nf(TOT.P)} + (RV) ${nf(TOT.rv)} + (C) ${nf(TOT.C)}.\n` +
  `KPI = (P) + (AP) sobre (P) + (AP) + (AE) + (C) — el % pendiente, donde menos es mejor.`,
  [[pc(TOT.kpi), "KPI\n% pendiente"],
   [nf(TOT.C), "(C)\nCerrados"],
   [nf(TOT.enFalta), "(P) + (AP)\nnumerador del KPI"],
   [String(PROY.length), "PROYECTOS\nactivos"]],
  `Corte ${CORTE}`, CAP);

// =====================================================================
// 2 — Resumen ejecutivo
// =====================================================================
s = p.addSlide(); bg(s, C.paper); logo(s, false);
kicker(s, "Resumen ejecutivo", 0.5, 0.45);
titulo(s, "Dónde está el pendiente de protocolos");
const semT = SEM(TOT.kpi);
bajada(s, `El KPI corporativo es ${pc(TOT.kpi)} de pendiente sobre ${nf(TOT.base)} protocolos exigibles. ` +
  `${peor.nombre} es el frente con más pendiente (${pc(peor.kpi)}).`);
tarjetas(s, [
  { p: pc(TOT.kpi), v: semT[1], t: "KPI — % pendiente",
    d: `(P) + (AP) sobre (P)+(AP)+(AE)+(C) · ${nf(TOT.enFalta)} de ${nf(TOT.base)}`, c: semT[0] },
  { p: nf(TOT.C), v: pct1(TOT.C, TOT.universo) + "% del universo", t: "(C) Cerrados",
    d: `Más ${nf(TOT.AE)} en (AE) Ab. Ejecución`, c: C.blue },
  { p: nf(TOT.enFalta), v: `${nf(TOT.P)} en (P) · ${nf(TOT.AP)} en (AP)`, t: "(P) + (AP)",
    d: "Es el numerador del KPI", c: TOT.enFalta ? C.warnD : C.goodD },
  { p: nf(TOT.remanente), v: `${pct1(TOT.remanente, TOT.universo)}% del universo`, t: "(S) Sí Aplica",
    d: "Aplica pero todavía no se exige: queda fuera del KPI", c: C.neutral },
]);
escala(s, 0.5, 4.5);
s.addShape(p.ShapeType.roundRect, { x: 0.5, y: 5.08, w: 12.33, h: 1.42, rectRadius: 0.08,
  fill: { color: C.navy }, line: { type: "none" } });
kicker(s, "Lectura del corte", 0.8, 5.26, C.copperL);
s.addText(
  `Universo ${nf(TOT.universo)} = (S) ${nf(TOT.S)} + (AE) ${nf(TOT.AE)} + (AP) ${nf(TOT.AP)} + ` +
  `(P) ${nf(TOT.P)} + (RV) ${nf(TOT.rv)} + (C) ${nf(TOT.C)}. El KPI deja fuera el (S): se mide sobre ` +
  `${nf(TOT.base)}.\n` +
  PROY.map((x) => `${x.nombre} ${pc(x.kpi)} (${SEM(x.kpi)[1]}).`).join(" ") +
  (CORTES.length > 1 ? ` Cortes mixtos: ${CORTES.join(" · ")}.` : ""),
  { x: 0.8, y: 5.58, w: 11.7, h: 0.85, fontFace: FT, fontSize: 12, color: C.ice,
    lineSpacing: 17, margin: 0, valign: "top" });
footer(s, false, CAP, "consolidado");

// =====================================================================
// 3 — Semáforo por proyecto
// =====================================================================
s = p.addSlide(); bg(s, C.paper); logo(s, false);
kicker(s, "Estado por frente", 0.5, 0.45);
titulo(s, "Semáforo de protocolos");
bajada(s, `KPI = (P) + (AP) sobre (P) + (AP) + (AE) + (C). El (S) queda fuera del denominador: ` +
  `aplica, pero todavía no se exige.`);
// Mismas columnas y mismo orden que la tabla del dashboard, para que quien mire
// las dos no tenga que traducir nada.
const colsSem = [
  { t: "FRENTE", x: 0.55, w: 2.5, al: "left" },
  { t: "(S)", x: 3.15, w: 1.0 },
  { t: "(AE)", x: 4.25, w: 1.0 },
  { t: "(AP)", x: 5.35, w: 0.9 },
  { t: "(P)", x: 6.35, w: 0.9 },
  { t: "(RV)", x: 7.35, w: 0.9 },
  { t: "(C)", x: 8.35, w: 1.2 },
  { t: "UNIVERSO", x: 9.65, w: 1.2 },
  { t: "KPI %", x: 10.95, w: 0.9 },
  { t: "ESTADO", x: 11.95, w: 1.0, al: "center" },
];
colsSem.forEach((c) => s.addText(c.t, { x: c.x, y: 2.1, w: c.w, h: 0.3, fontFace: FT,
  fontSize: 9.5, bold: true, color: C.ink2, align: c.al || "right", charSpacing: 0.5, margin: 0 }));
s.addShape(p.ShapeType.line, { x: 0.55, y: 2.44, w: 12.3, h: 0, line: { color: C.steel, width: 1 } });
let rowy = 2.58; const rh = 0.82;
PROY.forEach((x, i) => {
  const est = SEM(x.kpi);
  if (i % 2 === 1) s.addShape(p.ShapeType.rect, { x: 0.5, y: rowy - 0.05, w: 12.35, h: rh, fill: { color: C.white }, line: { type: "none" } });
  s.addText(x.nombre, { x: colsSem[0].x, y: rowy, w: 2.45, h: 0.32, fontFace: FT, fontSize: 14, bold: true, color: C.ink, margin: 0 });
  s.addText(corta(`${x.cliente}${x.corte ? " · corte " + x.corte : ""}`, 34),
    { x: colsSem[0].x, y: rowy + 0.33, w: 2.45, h: 0.28, fontFace: FT, fontSize: 9, color: C.ink2, margin: 0 });
  const cy = rowy + 0.08;
  const celda = (i2, txt, color) => s.addText(txt, { x: colsSem[i2].x, y: cy, w: colsSem[i2].w, h: 0.4,
    fontFace: FT, fontSize: 13, color: color || C.ink, align: "right", margin: 0, valign: "middle" });
  celda(1, nf(x.S)); celda(2, nf(x.AE));
  celda(3, nf(x.AP), x.AP ? C.copper : C.ink); celda(4, nf(x.P), x.P ? C.copper : C.ink);
  celda(5, nf(x.rv)); celda(6, nf(x.C)); celda(7, nf(x.universo));
  celda(8, pc(x.kpi), est[0]);
  s.addShape(p.ShapeType.roundRect, { x: colsSem[9].x, y: rowy + 0.1, w: 1.0, h: 0.4,
    rectRadius: 0.2, fill: { color: est[0] }, line: { type: "none" } });
  s.addText(est[1], { x: colsSem[9].x, y: rowy + 0.11, w: 1.0, h: 0.38, fontFace: FT,
    fontSize: 9, bold: true, color: C.white, align: "center", margin: 0 });
  s.addShape(p.ShapeType.line, { x: 0.55, y: rowy + rh - 0.05, w: 12.3, h: 0, line: { color: "D9DFE6", width: 0.75 } });
  rowy += rh;
});
s.addText("Los 3 proyectos", { x: colsSem[0].x, y: rowy + 0.02, w: 2.45, h: 0.32, fontFace: FT, fontSize: 14, bold: true, color: C.ink, margin: 0 });
s.addText("consolidado", { x: colsSem[0].x, y: rowy + 0.35, w: 2.45, h: 0.28, fontFace: FT, fontSize: 9.5, color: C.ink2, margin: 0 });
[[1, nf(TOT.S)], [2, nf(TOT.AE)], [3, nf(TOT.AP)], [4, nf(TOT.P)],
 [5, nf(TOT.rv)], [6, nf(TOT.C)], [7, nf(TOT.universo)], [8, pc(TOT.kpi)]].forEach(([i2, t]) =>
  s.addText(t, { x: colsSem[i2].x, y: rowy + 0.1, w: colsSem[i2].w, h: 0.4, fontFace: FT,
    fontSize: 13, bold: true, color: C.ink, align: "right", margin: 0, valign: "middle" }));
escala(s, 0.55, 5.95);
s.addText(ESTADOS.map(([c, n]) => `${c} ${n}`).join("   ·   "),
  { x: 0.55, y: 6.48, w: 12.3, h: 0.3, fontFace: FT, fontSize: 9.5, color: C.ink2, margin: 0 });
s.addText(`El (S) no entra en el denominador del KPI: son ${nf(TOT.S)} protocolos que aplican pero ` +
  `todavía no se exigen, y contarlos haría ver mejor de lo que se está.`,
  { x: 0.55, y: 6.76, w: 12.3, h: 0.3, fontFace: FT, fontSize: 9.5, italic: true, color: C.ink2, margin: 0 });
footer(s, false, CAP, "consolidado");

// =====================================================================
// 4 — Pendiente por disciplina
// =====================================================================
s = p.addSlide(); bg(s, C.paper); logo(s, false);
kicker(s, "Dónde está lo que falta", 0.5, 0.45);
titulo(s, "Pendiente por disciplina");
// Se agregan las disciplinas de los tres proyectos por nombre: es el corte que
// se mira en la reunión, y una disciplina puede estar en más de un frente.
const porDisc = {};
PROY.forEach((x) => x.disc.forEach((d) => {
  const t = porDisc[d.nombre] || (porDisc[d.nombre] = cero());
  for (const k of ["S", "C", "P", "AP", "AE", "rv"]) t[k] += d[k];
}));
const DISC = Object.entries(porDisc).map(([nombre, t]) => ({ nombre, ...metricas(t) }))
  .sort((a, b) => b.enFalta - a.enFalta || b.base - a.base);
const conFalta = DISC.filter((d) => d.enFalta > 0).slice(0, 8);
if (conFalta.length) {
  s.addChart(p.ChartType.bar, [{
    name: "En falta", labels: conFalta.map((d) => d.nombre), values: conFalta.map((d) => z(d.enFalta)),
  }], { ...BARRAS, x: 0.5, y: 1.95, w: 6.5, h: 4.2, chartColors: [C.warnD],
        title: `(P) + (AP) por disciplina · ${nf(TOT.enFalta)} en total`,
        catAxisLabelFontSize: 11, barGapWidthPct: 55, showLegend: false });
} else {
  s.addText("Ninguna disciplina tiene protocolos en (P) ni en (AP).",
    { x: 0.5, y: 3.4, w: 6.5, h: 0.5, fontFace: FT, fontSize: 14, color: C.goodD, margin: 0 });
}
bajada(s, conFalta.length
  ? `${nf(TOT.enFalta)} protocolos en (P) o (AP) —el numerador del KPI—, repartidos en ` +
    `${DISC.filter((d) => d.enFalta > 0).length} disciplinas. ${conFalta[0].nombre} concentra ${nf(conFalta[0].enFalta)}.`
  : `Ninguna de las ${DISC.length} disciplinas tiene protocolos en (P) ni en (AP).`);
tabla(s, [
  { t: "DISCIPLINA", x: 7.3, w: 2.1, al: "left" },
  { t: "(C)", x: 9.5, w: 1.0 },
  { t: "(P)+(AP)", x: 10.6, w: 1.1 },
  { t: "UNIVERSO", x: 11.8, w: 1.05 },
], DISC.slice(0, 8).map((d) => [
  { txt: d.nombre, bold: true }, nf(d.C),
  { txt: nf(d.enFalta), color: d.enFalta ? C.copper : C.ink2 }, nf(d.universo),
]), 2.1, 0.44, { ancho: 5.55, size: 11 });
s.addText(ESTADOS.map(([c, n]) => `${c} ${n}`).join("   ·   "),
  { x: 0.5, y: 6.22, w: 12.3, h: 0.3, fontFace: FT, fontSize: 9.5, color: C.ink2, margin: 0 });
footer(s, false, CAP, "consolidado");

// =====================================================================
// 5 — Tendencia del KPI
// =====================================================================
if (KPI_HISTORY.length >= 2) {
  s = p.addSlide(); bg(s, C.paper); logo(s, false);
  kicker(s, "Cómo viene la curva", 0.5, 0.45);
  titulo(s, "Tendencia del pendiente");
  const hist = KPI_HISTORY.slice().sort((a, b) => String(a.date).localeCompare(String(b.date))).slice(-10);
  const mesCorto = (d) => {
    const M = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
    return `${String(d).slice(8, 10)}-${M[+String(d).slice(5, 7) - 1]}`;
  };
  const labels = hist.map((h) => mesCorto(h.date));
  const series = PIDS.map((id, i) => ({
    name: PROY[i] ? PROY[i].nombre : id,
    labels, values: hist.map((h) => (h.kpi && h.kpi[id] != null ? h.kpi[id] : null)),
  }));
  series.push({ name: "Corporativo", labels, values: hist.map((h) => (h.kpi ? h.kpi.CORP : null)) });
  s.addChart(p.ChartType.line, series, {
    x: 0.5, y: 1.95, w: 12.33, h: 4.2,
    chartColors: [C.blue, C.copperD, C.goodD, C.navy],
    lineDataSymbol: "circle", lineDataSymbolSize: 6, lineSize: 2.5,
    showTitle: true, title: "% pendiente por corte — menos es mejor",
    titleFontFace: FT, titleFontSize: 13, titleColor: C.ink,
    showLegend: true, legendPos: "b", legendFontFace: FT, legendFontSize: 11, legendColor: C.ink2,
    catAxisLabelFontFace: FT, catAxisLabelColor: C.ink, catAxisLabelFontSize: 10,
    valAxisLabelFontFace: FT, valAxisLabelColor: C.ink2, valAxisLabelFontSize: 10,
    valGridLine: { color: "E4E9EF", style: "solid" }, catGridLine: { style: "none" },
    chartArea: { fill: { color: C.white } },
  });
  const ult = hist[hist.length - 1], ant = hist[hist.length - 2];
  const dCorp = (ult.kpi?.CORP ?? 0) - (ant.kpi?.CORP ?? 0);
  s.addText(
    `Del corte ${ant.date} al ${ult.date} el pendiente corporativo ` +
    (Math.abs(dCorp) < 0.005 ? "no se movió" : dCorp < 0 ? `bajó ${pc(Math.abs(Math.round(dCorp * 100) / 100))}` : `subió ${pc(Math.round(dCorp * 100) / 100)}`) +
    `, hasta ${pc(ult.kpi?.CORP ?? 0)}. Un alza no siempre es retroceso: cuando un proyecto entrega ` +
    `protocolos a revisión del cliente, estos pasan a contar como pendientes.`,
    { x: 0.5, y: 6.35, w: 12.3, h: 0.55, fontFace: FT, fontSize: 9.5, italic: true,
      color: C.ink2, margin: 0, valign: "top" });
  footer(s, false, CAP, "consolidado");
}

// =====================================================================
// Por proyecto — portada + detalle por disciplina
// =====================================================================
PROY.forEach((x) => {
  const est = SEM(x.kpi);
  s = p.addSlide();
  portada(s, x.etiqueta.toUpperCase(), x.nombre, x.cliente || "Control de protocolos",
    `Universo ${nf(x.universo)} = (S) ${nf(x.S)} + (AE) ${nf(x.AE)} + (AP) ${nf(x.AP)} + ` +
    `(P) ${nf(x.P)} + (RV) ${nf(x.rv)} + (C) ${nf(x.C)}.\n` +
    `KPI = (P) + (AP) sobre (P)+(AP)+(AE)+(C): ${nf(x.enFalta)} de ${nf(x.base)}.`,
    [[pc(x.kpi), "KPI\n% pendiente"],
     [nf(x.C), "(C)\nCerrados"],
     [nf(x.enFalta), "(P) + (AP)\nnumerador del KPI"],
     [est[1], "ESTADO"]],
    x.corte ? `Corte ${x.corte}` : "", CAP);

  // ── panel del proyecto ────────────────────────────────────────────────
  // Réplica de la banda que usaba el tablero anterior: la matriz ESTATUS ×
  // DISCIPLINA a la izquierda, la dona de avance al medio y una pastilla con
  // su dona por disciplina abajo. Es el corte que la gente ya sabe leer.
  s = p.addSlide(); bg(s, C.paper); logo(s, false);
  kicker(s, `${x.etiqueta} · ${x.cliente}`, 0.5, 0.45);
  titulo(s, "Indicadores de protocolos");

  // Matriz: una fila por estado, una columna por disciplina.
  const ds = x.disc.slice().sort((a, b) => b.universo - a.universo);
  const hayRV = ds.some((d) => d.rv > 0) || x.rv > 0;
  const FILAS = [["(AE) Abierto en Ejecución", "AE"], ["(AP) Abierto Pendiente", "AP"],
                 ["(C) Cerrados", "C"], ["(P) Pendiente", "P"], ["(S) Sí Aplica", "S"]]
    .concat(hayRV ? [["(RV) Revisión", "rv"]] : []);
  // El ancho de columna se ajusta al número de disciplinas para que la matriz
  // cierre siempre antes de 8,4": desde ahí arranca la dona.
  const wCol = Math.min(0.70, 6.0 / Math.max(ds.length + 1, 1));
  const x0 = 0.45, wEt = 1.95;
  const colX = (i) => x0 + wEt + i * wCol;

  s.addText("ESTATUS", { x: x0, y: 1.72, w: wEt, h: 0.3, fontFace: FT, fontSize: 9.5,
    bold: true, color: C.ink2, charSpacing: 0.5, margin: 0 });
  ds.forEach((d, i) => s.addText(cod(d.nombre), { x: colX(i), y: 1.72, w: wCol, h: 0.3,
    fontFace: FT, fontSize: 9.5, bold: true, color: C.ink2, align: "right", margin: 0 }));
  s.addText("TOTAL", { x: colX(ds.length), y: 1.72, w: wCol, h: 0.3, fontFace: FT,
    fontSize: 9.5, bold: true, color: C.ink2, align: "right", margin: 0 });
  s.addShape(p.ShapeType.line, { x: x0, y: 2.06, w: wEt + (ds.length + 1) * wCol, h: 0,
    line: { color: C.steel, width: 1 } });

  // La altura de fila se ajusta al número de estados: Talabre trae además (RV)
  // y con altura fija la matriz crecía hasta pisar las pastillas de abajo.
  let fy = 2.16; const fh = Math.min(0.36, 2.06 / (FILAS.length + 1));
  FILAS.forEach(([et, k], n) => {
    if (n % 2 === 1) s.addShape(p.ShapeType.rect, { x: x0 - 0.04, y: fy - 0.04,
      w: wEt + (ds.length + 1) * wCol + 0.08, h: fh, fill: { color: C.white }, line: { type: "none" } });
    s.addText(et, { x: x0, y: fy, w: wEt, h: fh - 0.06, fontFace: FT, fontSize: 10,
      color: C.ink, margin: 0, valign: "middle" });
    ds.forEach((d, i) => s.addText(nf(d[k]), { x: colX(i), y: fy, w: wCol, h: fh - 0.06,
      fontFace: FT, fontSize: 10, color: (k === "AP" || k === "P") && d[k] ? C.crit : C.ink,
      align: "right", margin: 0, valign: "middle" }));
    s.addText(nf(x[k]), { x: colX(ds.length), y: fy, w: wCol, h: fh - 0.06, fontFace: FT,
      fontSize: 10, bold: true, color: C.ink, align: "right", margin: 0, valign: "middle" });
    fy += fh;
  });
  s.addShape(p.ShapeType.line, { x: x0, y: fy - 0.04, w: wEt + (ds.length + 1) * wCol, h: 0,
    line: { color: C.steel, width: 1 } });
  s.addText("Total", { x: x0, y: fy + 0.02, w: wEt, h: fh - 0.06, fontFace: FT, fontSize: 10,
    bold: true, color: C.ink, margin: 0, valign: "middle" });
  ds.forEach((d, i) => s.addText(nf(d.universo), { x: colX(i), y: fy + 0.02, w: wCol, h: fh - 0.06,
    fontFace: FT, fontSize: 10, bold: true, color: C.ink, align: "right", margin: 0, valign: "middle" }));
  s.addText(nf(x.universo), { x: colX(ds.length), y: fy + 0.02, w: wCol, h: fh - 0.06,
    fontFace: FT, fontSize: 10, bold: true, color: C.navy, align: "right", margin: 0, valign: "middle" });

  // Dona del proyecto y su KPI, a la derecha de la matriz.
  dona(s, x, 8.6, 1.85, 2.2, [10.95, 3.45]);
  const estP = SEM(x.kpi);
  s.addShape(p.ShapeType.roundRect, { x: 10.95, y: 1.95, w: 1.9, h: 1.25, rectRadius: 0.1,
    fill: { color: estP[0] }, line: { type: "none" } });
  s.addText(pc(x.kpi), { x: 10.95, y: 2.1, w: 1.9, h: 0.55, fontFace: FH, fontSize: 30,
    bold: true, color: C.white, align: "center", margin: 0 });
  s.addText(`KPI ${x.etiqueta.split(" ")[0]}`, { x: 10.95, y: 2.66, w: 1.9, h: 0.26, fontFace: FT,
    fontSize: 10, color: C.white, align: "center", margin: 0 });
  s.addText(estP[1], { x: 10.95, y: 2.9, w: 1.9, h: 0.24, fontFace: FT, fontSize: 10.5,
    bold: true, color: C.white, align: "center", margin: 0 });
  // Una pastilla con su dona por disciplina, en el mismo orden de la matriz.
  const anchoD = Math.min(1.62, 12.43 / Math.max(ds.length, 1));
  let dx = 0.45;
  ds.forEach((d) => {
    const e = SEM(d.kpi);
    s.addShape(p.ShapeType.roundRect, { x: dx, y: 4.42, w: anchoD - 0.1, h: 0.62,
      rectRadius: 0.1, fill: { color: e[0] }, line: { type: "none" } });
    s.addText(pc(d.kpi), { x: dx, y: 4.46, w: anchoD - 0.1, h: 0.3, fontFace: FH,
      fontSize: 15, bold: true, color: C.white, align: "center", margin: 0 });
    s.addText(`${cod(d.nombre)} · ${corta(d.nombre, 14)}`, { x: dx, y: 4.76, w: anchoD - 0.1, h: 0.24,
      fontFace: FT, fontSize: 8, color: C.white, align: "center", margin: 0 });
    dona(s, d, dx + (anchoD - 0.1 - 1.18) / 2, 5.12, 1.18, false);
    dx += anchoD;
  });
  s.addText(`Universo = (S) + (AE) + (AP) + (P)${hayRV ? " + (RV)" : ""} + (C). ` +
    `El KPI es (P)+(AP) sobre (P)+(AP)+(AE)+(C): el (S) no entra en el denominador.`,
    { x: 0.45, y: 6.52, w: 12.4, h: 0.3, fontFace: FT, fontSize: 9.5, italic: true,
      color: C.ink2, margin: 0 });
  footer(s, false, CAP, x.nombre);
});

// =====================================================================
// Guardar y embeber
// =====================================================================
p.writeFile({ fileName: SALIDA }).then(() => {
  const laminas = nSlide - 1;
  const b64 = fs.readFileSync(SALIDA).toString("base64");
  const peso = (fs.statSync(SALIDA).size / 1048576).toFixed(1) + " MB";
  const fecha = (CORTE || "").replace(/[^0-9a-zA-Z]/g, "") || "corte";
  const nombre = `Informe_Protocolos_${fecha}.pptx`;

  // Se reemplaza el bloque anterior si ya estaba: el dashboard llega nuevo cada
  // semana, pero si alguien corre esto dos veces no se acumulan copias.
  let html = HTML0;
  const a = html.indexOf(MARCA_INI);
  if (a >= 0) {
    const b = html.indexOf(MARCA_FIN, a);
    html = html.slice(0, a) + html.slice(b + MARCA_FIN.length);
  }
  const bloque = `
${MARCA_INI}
<script>
/* El informe ya generado y verificado, embebido igual que en Cierre QAQC y No
   Conformidades. El botón lo descarga de un clic: no se arma nada en el
   navegador, así que lo que baja el usuario es exactamente lo que revisó
   verificar_suite.py. */
const PPTX_PROT = {"nombre":${JSON.stringify(nombre)},"corte":${JSON.stringify(CORTE)},"laminas":${laminas},"peso":${JSON.stringify(peso)},"b64":"${b64}"};
(function () {
  const b = document.getElementById('reportBtn');
  if (!b) return;
  // El modal de generación queda fuera: sobra cuando el archivo ya existe.
  b.removeAttribute('onclick');
  const m = document.getElementById('reportModal');
  if (m) m.remove();
  b.title = 'PPT del corte ' + PPTX_PROT.corte + ' · ' + PPTX_PROT.laminas + ' láminas · ' + PPTX_PROT.peso;
  b.addEventListener('click', function () {
    try {
      const bin = atob(PPTX_PROT.b64);
      const buf = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
      const url = URL.createObjectURL(new Blob([buf],
        { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }));
      const a = document.createElement('a');
      a.href = url; a.download = PPTX_PROT.nombre;
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(function () { URL.revokeObjectURL(url); }, 2000);
    } catch (e) {
      b.textContent = '⚠ No se pudo descargar';
      console.error(e);
    }
  });
})();
</script>
${MARCA_FIN}
`;
  const cierre = html.lastIndexOf("</body>");
  html = cierre >= 0 ? html.slice(0, cierre) + bloque + html.slice(cierre) : html + bloque;
  fs.writeFileSync(MOD, html, "utf8");

  console.log(`PPT generada: ${SALIDA}`);
  console.log(`Informe embebido en modulos/protocolos.html (${nombre} · ${laminas} láminas · ${peso}).`);
  if (CORTES.length > 1) console.log(`Aviso: cortes mixtos entre proyectos — ${CORTES.join(" · ")}`);
});
