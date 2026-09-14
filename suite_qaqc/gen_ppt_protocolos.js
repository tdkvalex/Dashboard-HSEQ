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
// Mismo semáforo que usa la portada de la suite: % pendiente, menos es mejor.
const SEM = (kpi) =>
  kpi === 0 ? [C.goodD, "Cumple"] : kpi <= 3 ? [C.goodD, "Aceptable"]
  : kpi <= 15 ? [C.warnD, "Medianamente"] : kpi <= 50 ? [C.crit, "Deficiente"]
  : [C.crit, "No cumple"];

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
  `${nf(TOT.universo)} protocolos en el universo, de los que ${nf(TOT.C)} están cerrados y ` +
  `${nf(TOT.enFalta)} quedan en falta.\nEl KPI mide lo pendiente sobre lo exigible ` +
  `(${nf(TOT.base)} protocolos): el remanente de ${nf(TOT.remanente)} que aún no aplica queda fuera.`,
  [[pc(TOT.kpi), "PENDIENTE\nsobre lo exigible"],
   [nf(TOT.C), "PROTOCOLOS\ncerrados"],
   [nf(TOT.enFalta), "EN FALTA\npendientes y abiertos"],
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
  { p: pc(TOT.kpi), v: `${nf(TOT.enFalta)} de ${nf(TOT.base)}`, t: "Pendiente sobre lo exigible",
    d: "Menos es mejor: mide lo que falta por levantar o cerrar", c: semT[0] },
  { p: nf(TOT.C), v: pct1(TOT.C, TOT.universo) + "% del universo", t: "Protocolos cerrados",
    d: `Más ${nf(TOT.AE)} abiertos en ejecución`, c: C.blue },
  { p: nf(TOT.enFalta), v: `${nf(TOT.P)} pendientes · ${nf(TOT.AP)} abiertos`, t: "En falta",
    d: "Es el numerador del KPI", c: TOT.enFalta ? C.warnD : C.goodD },
  { p: nf(TOT.remanente), v: `${pct1(TOT.remanente, TOT.universo)}% del universo`, t: "Remanente",
    d: "Aplica pero todavía no es exigible: queda fuera del KPI", c: C.neutral },
]);
s.addShape(p.ShapeType.roundRect, { x: 0.5, y: 4.35, w: 12.33, h: 2.1, rectRadius: 0.08,
  fill: { color: C.navy }, line: { type: "none" } });
kicker(s, "Lectura del corte", 0.8, 4.6, C.copperL);
s.addText(
  `El universo son ${nf(TOT.universo)} protocolos. De ellos, ${nf(TOT.remanente)} son remanente —aplican pero ` +
  `aún no se exigen— y quedan ${nf(TOT.base)} sobre los que se mide el KPI.\n` +
  PROY.map((x) => `${x.nombre}: ${pc(x.kpi)} de pendiente sobre ${nf(x.base)} exigibles, ${nf(x.C)} cerrados.`).join(" ") +
  (CORTES.length > 1 ? `\nOjo: los proyectos no vienen del mismo corte (${CORTES.join(" · ")}).` : ""),
  { x: 0.8, y: 4.95, w: 11.7, h: 1.3, fontFace: FT, fontSize: 12.5, color: C.ice,
    lineSpacing: 18, margin: 0, valign: "top" });
footer(s, false, CAP, "consolidado");

// =====================================================================
// 3 — Semáforo por proyecto
// =====================================================================
s = p.addSlide(); bg(s, C.paper); logo(s, false);
kicker(s, "Estado por frente", 0.5, 0.45);
titulo(s, "Semáforo de protocolos");
bajada(s, "Cumple con 0% de pendiente · Aceptable hasta 3% · Medianamente hasta 15% · " +
  "Deficiente hasta 50% · No cumple por encima. Menos es mejor.");
const colsSem = [
  { t: "FRENTE", x: 0.55, w: 2.7, al: "left" },
  { t: "UNIVERSO", x: 3.35, w: 1.2 },
  { t: "CERRADOS", x: 4.65, w: 1.2 },
  { t: "EN EJECUCIÓN", x: 5.95, w: 1.3 },
  { t: "REMANENTE", x: 7.35, w: 1.2 },
  { t: "EN FALTA", x: 8.65, w: 1.1 },
  { t: "BASE KPI", x: 9.85, w: 1.1 },
  { t: "KPI", x: 11.05, w: 0.8 },
  { t: "ESTADO", x: 11.95, w: 1.0, al: "center" },
];
colsSem.forEach((c) => s.addText(c.t, { x: c.x, y: 2.1, w: c.w, h: 0.3, fontFace: FT,
  fontSize: 9.5, bold: true, color: C.ink2, align: c.al || "right", charSpacing: 0.5, margin: 0 }));
s.addShape(p.ShapeType.line, { x: 0.55, y: 2.44, w: 12.3, h: 0, line: { color: C.steel, width: 1 } });
let rowy = 2.58; const rh = 0.82;
PROY.forEach((x, i) => {
  const est = SEM(x.kpi);
  if (i % 2 === 1) s.addShape(p.ShapeType.rect, { x: 0.5, y: rowy - 0.05, w: 12.35, h: rh, fill: { color: C.white }, line: { type: "none" } });
  s.addText(x.nombre, { x: colsSem[0].x, y: rowy, w: 2.65, h: 0.32, fontFace: FT, fontSize: 14, bold: true, color: C.ink, margin: 0 });
  s.addText(`${x.cliente}${x.corte ? " · corte " + x.corte : ""}`,
    { x: colsSem[0].x, y: rowy + 0.33, w: 2.7, h: 0.28, fontFace: FT, fontSize: 9.5, color: C.ink2, margin: 0 });
  const cy = rowy + 0.08;
  const celda = (i2, txt, color) => s.addText(txt, { x: colsSem[i2].x, y: cy, w: colsSem[i2].w, h: 0.4,
    fontFace: FT, fontSize: 13, color: color || C.ink, align: "right", margin: 0, valign: "middle" });
  celda(1, nf(x.universo)); celda(2, nf(x.C)); celda(3, nf(x.AE));
  celda(4, nf(x.remanente)); celda(5, nf(x.enFalta), x.enFalta ? C.copper : C.ink);
  celda(6, nf(x.base)); celda(7, pc(x.kpi), est[0]);
  s.addShape(p.ShapeType.roundRect, { x: colsSem[8].x, y: rowy + 0.1, w: 1.0, h: 0.4,
    rectRadius: 0.2, fill: { color: est[0] }, line: { type: "none" } });
  s.addText(est[1], { x: colsSem[8].x, y: rowy + 0.11, w: 1.0, h: 0.38, fontFace: FT,
    fontSize: 9.5, bold: true, color: C.white, align: "center", margin: 0 });
  s.addShape(p.ShapeType.line, { x: 0.55, y: rowy + rh - 0.05, w: 12.3, h: 0, line: { color: "D9DFE6", width: 0.75 } });
  rowy += rh;
});
s.addText("Los 3 proyectos", { x: colsSem[0].x, y: rowy + 0.02, w: 2.65, h: 0.32, fontFace: FT, fontSize: 14, bold: true, color: C.ink, margin: 0 });
s.addText("consolidado", { x: colsSem[0].x, y: rowy + 0.35, w: 2.7, h: 0.28, fontFace: FT, fontSize: 9.5, color: C.ink2, margin: 0 });
[[1, nf(TOT.universo)], [2, nf(TOT.C)], [3, nf(TOT.AE)], [4, nf(TOT.remanente)],
 [5, nf(TOT.enFalta)], [6, nf(TOT.base)], [7, pc(TOT.kpi)]].forEach(([i2, t]) =>
  s.addText(t, { x: colsSem[i2].x, y: rowy + 0.1, w: colsSem[i2].w, h: 0.4, fontFace: FT,
    fontSize: 13, bold: true, color: C.ink, align: "right", margin: 0, valign: "middle" }));
s.addText(`El KPI se mide sobre la BASE, no sobre el universo: el remanente de ${nf(TOT.remanente)} protocolos ` +
  `aplica pero todavía no es exigible, y meterlo en el denominador haría ver mejor de lo que se está.`,
  { x: 0.5, y: 6.5, w: 12.3, h: 0.4, fontFace: FT, fontSize: 9.5, italic: true, color: C.ink2, margin: 0 });
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
        title: `Protocolos en falta por disciplina · ${nf(TOT.enFalta)} en total`,
        catAxisLabelFontSize: 11, barGapWidthPct: 55, showLegend: false });
} else {
  s.addText("Ninguna disciplina tiene protocolos en falta.",
    { x: 0.5, y: 3.4, w: 6.5, h: 0.5, fontFace: FT, fontSize: 14, color: C.goodD, margin: 0 });
}
bajada(s, conFalta.length
  ? `${nf(TOT.enFalta)} protocolos en falta, repartidos en ${DISC.filter((d) => d.enFalta > 0).length} disciplinas. ` +
    `${conFalta[0].nombre} concentra ${nf(conFalta[0].enFalta)}.`
  : `No hay protocolos en falta: las ${DISC.length} disciplinas están al día.`);
tabla(s, [
  { t: "DISCIPLINA", x: 7.3, w: 2.3, al: "left" },
  { t: "UNIVERSO", x: 9.7, w: 1.0 },
  { t: "CERRADOS", x: 10.8, w: 1.0 },
  { t: "EN FALTA", x: 11.9, w: 0.95 },
], DISC.slice(0, 9).map((d) => [
  { txt: d.nombre, bold: true }, nf(d.universo), nf(d.C),
  { txt: nf(d.enFalta), color: d.enFalta ? C.copper : C.ink2 },
]), 2.1, 0.44, { ancho: 5.55, size: 11 });
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
    `${nf(x.universo)} protocolos en el universo · ${nf(x.C)} cerrados · ${nf(x.AE)} abiertos en ejecución.\n` +
    `El KPI mide ${nf(x.enFalta)} en falta sobre ${nf(x.base)} exigibles; ${nf(x.remanente)} de remanente quedan fuera.`,
    [[pc(x.kpi), "PENDIENTE\nsobre lo exigible"],
     [nf(x.C), "CERRADOS"],
     [nf(x.enFalta), "EN FALTA"],
     [est[1], "ESTADO"]],
    x.corte ? `Corte ${x.corte}` : "", CAP);

  s = p.addSlide(); bg(s, C.paper); logo(s, false);
  kicker(s, x.nombre, 0.5, 0.45);
  titulo(s, "Estado por disciplina");
  const conF = x.disc.filter((d) => d.enFalta > 0).sort((a, b) => b.enFalta - a.enFalta);
  bajada(s, conF.length
    ? `${nf(x.enFalta)} protocolos en falta en ${conF.length} de las ${x.disc.length} disciplinas. ` +
      `${conF[0].nombre} concentra ${nf(conF[0].enFalta)}.`
    : `Ninguna de las ${x.disc.length} disciplinas tiene protocolos en falta al corte.`);
  const filas = x.disc.slice().sort((a, b) => b.universo - a.universo).map((d) => {
    const e = SEM(d.kpi);
    return [
      { txt: d.nombre, bold: true }, nf(d.universo), nf(d.C), nf(d.AE),
      nf(d.remanente), { txt: nf(d.enFalta), color: d.enFalta ? C.copper : C.ink2 },
      { txt: pc(d.kpi), color: e[0], bold: true },
    ];
  });
  tabla(s, [
    { t: "DISCIPLINA", x: 0.55, w: 3.2, al: "left" },
    { t: "UNIVERSO", x: 4.0, w: 1.4 },
    { t: "CERRADOS", x: 5.5, w: 1.4 },
    { t: "EN EJECUCIÓN", x: 7.0, w: 1.5 },
    { t: "REMANENTE", x: 8.6, w: 1.4 },
    { t: "EN FALTA", x: 10.1, w: 1.3 },
    { t: "KPI", x: 11.5, w: 1.3 },
  ], filas, 2.1, Math.min(0.52, 4.3 / Math.max(filas.length, 1)), { size: 11.5 });
  s.addText(`Universo = cerrados + en ejecución + remanente + en falta. El KPI deja el remanente fuera ` +
    `del denominador: aplica, pero todavía no es exigible.`,
    { x: 0.5, y: 6.6, w: 12.3, h: 0.4, fontFace: FT, fontSize: 9.5, italic: true, color: C.ink2, margin: 0 });
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
