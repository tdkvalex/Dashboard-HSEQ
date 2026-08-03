/*
 * Genera la PPT ejecutiva de No Conformidades desde datos_nc.json.
 * No contiene cifras escritas a mano: todo se lee de los datos.
 *
 * Mismo formato y metodología que la PPT de Cierre QAQC
 * (panel_control_TOP_P1/gen_ppt.js): portada + láminas corporativas, después
 * una portada y dos láminas por proyecto, logo en todas, etiqueta del proyecto
 * al pie y numeración correlativa automática.
 *
 * Uso:  node gen_ppt.js     (correr siempre DESPUÉS de no_conformidades.py)
 */
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");

const AQUI = __dirname;
const D = JSON.parse(fs.readFileSync(path.join(AQUI, "datos_nc.json"), "utf8"));

// ---------- formato ----------
const nf = (n) => (n || 0).toLocaleString("es-CL");
const pctN = (a, b) => (b ? Math.round((1000 * a) / b) / 10 : 0);
const pct = (a, b) => pctN(a, b).toLocaleString("es-CL") + "%";
const pc1 = (v) => (v == null ? "—" : Number(v).toLocaleString("es-CL") + "%");
const corta = (t, n) => {
  const s = String(t || "").trim();
  return s.length > n ? s.slice(0, n - 1) + "…" : s;
};

const C = {
  navy: "12233B", navy2: "1B3252", steel: "6E8CA8", ice: "DCE6F0",
  paper: "F4F6F9", white: "FFFFFF", ink: "1A2433", ink2: "56606E",
  copper: "C87A32", copperL: "E3A55B",
  good: "2E9E5B", goodD: "1F7A44", warn: "E0A32E", crit: "C0392B",
  blue: "2E6FB5", blueL: "9DBDE0", track: "D9DFE6",
  neutral: "5B6B7C", blueM: "3A6FA8", warnD: "9A6B0F",
  // Cobre oscurecido para el tercer segmento de «origen»: el copper original
  // daba 3,4:1 con texto blanco y el número quedaba ilegible encima.
  copperD: "8A5520",
};

const FT = "Calibri", FH = "Cambria";
// Los valores 0 no se dibujan: PowerPoint los imprime pegados al eje y pisan
// la etiqueta vecina.
const z = (v) => (v > 0 ? v : null);

// Etiquetas de los gráficos: número blanco, negrita, 11 pt. Los colores de
// segmento que las llevan encima deben dar >=4,5:1 con blanco.
const ETIQ = { showValue: true, dataLabelPosition: "ctr", dataLabelFontFace: FT,
  dataLabelFontSize: 11, dataLabelColor: "FFFFFF", dataLabelFontBold: true };

// Opciones comunes de los gráficos de barras horizontales, para que las 16
// láminas se lean igual.
const BARRAS = {
  barDir: "bar", barGrouping: "stacked", ...ETIQ,
  titleFontFace: FT, titleFontSize: 13, titleColor: C.ink, showTitle: true,
  showLegend: true, legendPos: "b", legendFontFace: FT, legendFontSize: 11, legendColor: C.ink2,
  catAxisLabelFontFace: FT, catAxisLabelColor: C.ink, catAxisLabelFontBold: true,
  valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
  chartArea: { fill: { color: C.white } },
};

// Color por vía de emisión. La externa se abre en cliente y subcontrato porque
// no se gestionan igual: una compromete el contrato, la otra la absorbe Besalco.
const COL_EMI = {
  "Interna BSMT": C.blue,
  "Externa Cliente": C.warnD,
  "Externa Subcontrato": C.copperD,
  "Sin clasificar": C.neutral,
};
const EMI = D.etiquetasEmision || {};
const emi = (k) => EMI[k] || k;
// Desglose de las TRES vías de emisión, nunca «internas y externas»: la del
// cliente compromete el contrato y la del subcontrato la absorbe Besalco, así
// que no pueden compartir una sola cifra. Lo que vale 0 no se escribe.
const desg = (v) => [
  `${nf(v.internas)} ${v.internas === 1 ? "interna" : "internas"}`,
  v.cliente ? `${nf(v.cliente)} del cliente` : null,
  v.subcontrato ? `${nf(v.subcontrato)} de subcontrato` : null,
].filter(Boolean).join(" · ");
const EMISIONES = ["Interna BSMT", "Externa Cliente", "Externa Subcontrato", "Sin clasificar"];

const G = D.global, R = G.resumen, K = D.control;
const OBRA = D.obra.resumen;
const CAP = `Fuente: ${D.meta.fuente} — hoja «Observaciones» · Corte ${D.meta.corteTexto} · ` +
  `${nf(K.registros)} registros desde ${K.rangoFechas[0]}`;

const p = new pptxgen();
p.defineLayout({ name: "W", width: 13.333, height: 7.5 });
p.layout = "W";
p.author = "Gestión de Calidad";
p.title = "Control de No Conformidades — Gestión de Calidad";

const bg = (s, color) => { s.background = { color }; };

// ---------- logo corporativo ----------
// Se busca en esta carpeta y, si no está (el .gitignore local los excluye), en
// la del panel de Cierre QAQC, que sí los tiene versionados. Si no aparece en
// ninguna, las láminas salen igual: nunca rompe la generación.
const LOGO = [path.join(AQUI, "besalco_logo.png"),
              path.join(AQUI, "..", "panel_control_TOP_P1", "besalco_logo.png")]
  .find((r) => fs.existsSync(r));
if (!LOGO) console.log("Aviso: no se encontró besalco_logo.png — la PPT sale sin logo.");
function logo(s, portada) {
  if (!LOGO) return;
  const g = portada ? { x: 9.14, y: 0.46, w: 3.88, h: 0.84 }
                    : { x: 9.97, y: 0.24, w: 3.03, h: 0.65 };
  s.addImage({ path: LOGO, ...g });
}

// Pie: fuente + proyecto al que pertenece la lámina + número de página.
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

// ---------- portada de proyecto ----------
// Misma estructura para el consolidado y para los tres proyectos, para que el
// mazo se lea parejo.
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

// ---------- tarjetas KPI ----------
// Geometría fija: 1.6 -> 4.1. Todo el texto de la tarjeta cierra antes de 4.1.
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

// ---------- tabla ----------
// cols: [{t, x, w, al}] · filas: [[celda, ...]] · una celda puede ser
// {txt, color, bold} o un string.
// `o.ancho` es el largo de las líneas y de la banda cebra: por defecto ocupa la
// lámina entera, pero una tabla que comparte fila con un gráfico tiene que
// cortarse antes o le raya encima.
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

// ---------- derivados ----------
const PROY = D.orden.map((k) => ({ k, ...D.proyectos[k] }));
const OBRAS = PROY.filter((x) => x.obra);
const NOMBRES = Object.fromEntries(PROY.map((x) => [x.k, x.nombre]));

// El frente que concentra las NC abiertas: es la lectura de portada.
const foco = PROY.slice().sort((a, b) => b.resumen.abiertas - a.resumen.abiertas)[0];
// Frentes donde el cliente levanta más de lo que detecta Besalco. Se compara
// contra las del CLIENTE, no contra el total de externas: sumarle las del
// subcontrato —que Besalco emite— falsearía la comparación.
const masExternas = PROY.filter((x) => x.resumen.cliente > x.resumen.internas);
// Cierre más lento, en días de mediana.
const conMediana = PROY.filter((x) => x.resumen.medianaCierre != null);
const lento = conMediana.slice().sort((a, b) => b.resumen.medianaCierre - a.resumen.medianaCierre)[0];
const rapido = conMediana.slice().sort((a, b) => a.resumen.medianaCierre - b.resumen.medianaCierre)[0];

const ANT = G.antiguedad || {};
const viejas = (ANT["181-365 días"] || 0) + (ANT["Más de un año"] || 0);
const anioMas = ANT["Más de un año"] || 0;

const SEM = G.semana;
const semDisc = Object.entries(SEM.porEspecialidad || {});
const semPrev = SEM.tendencia[SEM.tendencia.length - 2] || { total: 0 };

const espGlobal = Object.entries(G.porEspecialidad).sort((a, b) => b[1].total - a[1].total);
const espAbiertas = espGlobal.filter(([, v]) => v.abiertas > 0)
  .sort((a, b) => b[1].abiertas - a[1].abiertas);

// =====================================================================
// 1 — Portada · consolidado
// =====================================================================
let s = p.addSlide();
portada(s, "GESTIÓN DE CALIDAD", "No Conformidades",
  `${D.meta.descripcion} · ${OBRAS.length} proyectos y Oficina Central`,
  `${nf(R.total)} hallazgos levantados entre ${K.rangoFechas[0]} y ${K.rangoFechas[1]} ` +
  `(${PROY.map((x) => `${x.nombre} ${nf(x.resumen.total)}`).join(" · ")}).\n` +
  "Estado de cierre · origen interno y externo · antigüedad de lo que sigue abierto · " +
  "velocidad de corrección por frente.",
  [[pct(R.cerradas, R.total), "Hallazgos cerrados"],
   [nf(R.abiertas), "Siguen abiertos"],
   [`${nf(R.medianaCierre)} d`, "Mediana de cierre"],
   [nf(Math.round(R.costo)), "UF declaradas"]],
  `Corte: ${D.meta.corteTexto}`, CAP);

// =====================================================================
// 2 — Resumen ejecutivo
// =====================================================================
s = p.addSlide(); bg(s, C.paper); logo(s, false);
kicker(s, "Resumen ejecutivo", 0.5, 0.45);
titulo(s, "Estado global de los cuatro frentes");

tarjetas(s, [
  { p: pct(R.cerradas, R.total), v: `${nf(R.cerradas)}/${nf(R.total)}`, t: "Cierre de hallazgos",
    d: `${nf(R.abiertas)} abiertos · ${nf(viejas)} con más de 180 días`, c: C.blue },
  { p: nf(R.abiertas), v: `de ${nf(R.total)} levantados`, t: "Hallazgos abiertos",
    d: Object.entries(G.porEstatus || {}).map(([k, v]) => `${nf(v)} ${k.toLowerCase()}`).join(" · "), c: C.crit },
  { p: pc1(G.autodeteccion), v: `${nf(R.internas)}/${nf(R.internas + R.externas)} clasificados`, t: "Autodetección",
    d: `Lo que levanta Besalco antes que el cliente o un subcontrato`, c: C.copper },
  { p: nf(Math.round(R.costo)), v: `${nf(R.conCosto)} de ${nf(R.total)} con costo`, t: "Costo declarado (UF)",
    d: `${nf(K.sinCosto)} hallazgos no declaran costo: la cifra es un piso`, c: C.goodD },
]);

s.addShape(p.ShapeType.roundRect, { x: 0.5, y: 4.3, w: 12.33, h: 2.42, rectRadius: 0.06, fill: { color: C.navy }, line: { type: "none" } });
s.addText("LECTURA DEL PERÍODO", { x: 0.85, y: 4.48, w: 6, h: 0.3, fontFace: FT, fontSize: 11, bold: true, color: C.copperL, charSpacing: 2, margin: 0 });

const reads = [
  [`${foco.nombre} concentra el pendiente.`,
   `${nf(foco.resumen.abiertas)} de los ${nf(R.abiertas)} hallazgos abiertos (${pct(foco.resumen.abiertas, R.abiertas)})` +
   (masExternas.length
     ? `, y es ${masExternas.length === 1 ? "el único frente" : "uno de los frentes"} donde el cliente levanta más de lo que detecta Besalco ` +
       `(${nf(foco.resumen.cliente)} del cliente contra ${nf(foco.resumen.internas)} internas).`
     : `.`)],
  [SEM.total
    ? `La semana trajo ${nf(SEM.total)} ${SEM.total === 1 ? "hallazgo nuevo" : "hallazgos nuevos"}${SEM.internas === 0 ? ", ninguno de autodetección." : "."}`
    : "La semana cerró sin hallazgos nuevos.",
   SEM.total
     ? `${SEM.desde} al ${SEM.hasta}: ${Object.entries(SEM.porEmision).map(([k, v]) => `${nf(v)} ${(D.frasesEmision || {})[k] || k}`).join(" · ")}` +
       ` en ${semDisc.map(([e, v]) => `${e} (${nf(v.total)})`).join(" y ")}. ` +
       `${SEM.abiertas === SEM.total ? "Todos siguen abiertos" : `${nf(SEM.abiertas)} siguen abiertos`}; la semana anterior entraron ${nf(semPrev.total)}.`
     : `Entre el ${SEM.desde} y el ${SEM.hasta} no se levantó ningún hallazgo en ningún frente.`],
  [K.atrasoCalculable
    ? `${nf(R.atrasadas)} de los ${nf(R.abiertas)} abiertos están fuera del plazo de respuesta.`
    : `El atraso no es medible con este archivo; se mide la antigüedad.`,
   K.atrasoCalculable
     ? `El plazo es de ${nf(K.plazoRespuesta)} días desde que se emite la NC. De los atrasados, ` +
       `${nf(viejas)} llevan más de 180 días abiertos y ${nf(anioMas)} más de un año: no es atraso de días, es deuda vieja.`
     : `Ninguno de los ${nf(K.abiertas)} abiertos trae fecha de emisión, así que no se les puede calcular el atraso.`],
];

let ry = 4.86;
reads.forEach(([h, d]) => {
  dot(s, 0.9, ry + 0.06, C.copper, 0.12);
  s.addText([{ text: h + "  ", options: { bold: true, color: C.white } },
             { text: d, options: { color: C.ice } }],
    { x: 1.15, y: ry - 0.05, w: 11.4, h: 0.6, fontFace: FT, fontSize: 12.5, lineSpacing: 16, margin: 0, valign: "top" });
  ry += 0.6;
});
footer(s, false, CAP, "consolidado");

// =====================================================================
// 3 — Semáforo por frente
// =====================================================================
s = p.addSlide(); bg(s, C.paper); logo(s, false);
kicker(s, "Estado por frente", 0.5, 0.45);
titulo(s, "Semáforo de no conformidades");
bajada(s, `Crítico si hay más de 15 abiertas o el cierre baja de 85% · Atención si hay más de 3 abiertas ` +
  `o el cierre queda bajo 95% · Al día en caso contrario.`);

const COLOR_EST = { good: C.good, warn: C.warn, crit: C.crit };
// «EXTERNAS» se parte en DEL CLIENTE y DE SUBCONTRATO: no son lo mismo y
// sumarlas escondía justo la que compromete el contrato.
const colsSem = [
  { t: "FRENTE", x: 0.55, w: 2.6, al: "left" },
  { t: "LEVANTADOS", x: 3.25, w: 1.0 },
  { t: "CERRADOS", x: 4.35, w: 0.95 },
  { t: "ABIERTOS", x: 5.40, w: 0.9 },
  { t: "% CIERRE", x: 6.40, w: 0.9 },
  { t: "INTERNAS", x: 7.40, w: 0.9 },
  { t: "DEL CLIENTE", x: 8.40, w: 1.1 },
  { t: "DE SUBCONTR.", x: 9.60, w: 1.1 },
  { t: "COSTO UF", x: 10.80, w: 1.0 },
  { t: "ESTADO", x: 11.95, w: 1.0, al: "center" },
];
colsSem.forEach((c) => s.addText(c.t, { x: c.x, y: 2.1, w: c.w, h: 0.3, fontFace: FT,
  fontSize: 9.5, bold: true, color: C.ink2, align: c.al || "right", charSpacing: 0.5, margin: 0 }));
s.addShape(p.ShapeType.line, { x: 0.55, y: 2.44, w: 12.3, h: 0, line: { color: C.steel, width: 1 } });

let rowy = 2.58; const rh = 0.82;
PROY.forEach((x, i) => {
  const v = x.resumen, est = D.semaforo[x.k] || ["warn", "Atención"];
  if (i % 2 === 1) s.addShape(p.ShapeType.rect, { x: 0.5, y: rowy - 0.05, w: 12.35, h: rh, fill: { color: C.white }, line: { type: "none" } });
  s.addText(x.nombre, { x: colsSem[0].x, y: rowy, w: 2.55, h: 0.32, fontFace: FT, fontSize: 14, bold: true, color: C.ink, margin: 0 });
  s.addText(x.cliente + (x.obra ? "" : " · no es obra"), { x: colsSem[0].x, y: rowy + 0.33, w: 2.6, h: 0.28, fontFace: FT, fontSize: 9.5, color: C.ink2, margin: 0 });
  const cy = rowy + 0.08;
  const celda = (idx, txt, color) => s.addText(txt, { x: colsSem[idx].x, y: cy, w: colsSem[idx].w, h: 0.4,
    fontFace: FT, fontSize: 13, color: color || C.ink, align: "right", margin: 0, valign: "middle" });
  celda(1, nf(v.total));
  celda(2, nf(v.cerradas));
  celda(3, nf(v.abiertas), v.abiertas > 15 ? C.crit : C.ink);
  celda(4, pct(v.cerradas, v.total));
  celda(5, nf(v.internas));
  // La del cliente en cobre cuando supera lo que Besalco detecta por su cuenta.
  celda(6, nf(v.cliente), v.cliente > v.internas ? C.copper : C.ink);
  celda(7, nf(v.subcontrato));
  celda(8, nf(Math.round(v.costo)));
  s.addShape(p.ShapeType.roundRect, { x: colsSem[9].x, y: rowy + 0.1, w: 1.0, h: 0.4,
    rectRadius: 0.2, fill: { color: COLOR_EST[est[0]] }, line: { type: "none" } });
  s.addText(est[1], { x: colsSem[9].x, y: rowy + 0.11, w: 1.0, h: 0.38, fontFace: FT,
    fontSize: 10.5, bold: true, color: C.white, align: "center", margin: 0 });
  s.addShape(p.ShapeType.line, { x: 0.55, y: rowy + rh - 0.05, w: 12.3, h: 0, line: { color: "D9DFE6", width: 0.75 } });
  rowy += rh;
});
// Total de obra: es lo comparable con los otros módulos de la suite.
s.addText("Los 3 proyectos", { x: colsSem[0].x, y: rowy + 0.02, w: 2.55, h: 0.32, fontFace: FT, fontSize: 14, bold: true, color: C.ink, margin: 0 });
s.addText("sin Oficina Central", { x: colsSem[0].x, y: rowy + 0.35, w: 2.6, h: 0.28, fontFace: FT, fontSize: 9.5, color: C.ink2, margin: 0 });
[[1, nf(OBRA.total)], [2, nf(OBRA.cerradas)], [3, nf(OBRA.abiertas)], [4, pct(OBRA.cerradas, OBRA.total)],
 [5, nf(OBRA.internas)], [6, nf(OBRA.cliente)], [7, nf(OBRA.subcontrato)],
 [8, nf(Math.round(OBRA.costo))]].forEach(([i, t]) =>
  s.addText(t, { x: colsSem[i].x, y: rowy + 0.1, w: colsSem[i].w, h: 0.4, fontFace: FT,
    fontSize: 13, bold: true, color: C.ink, align: "right", margin: 0, valign: "middle" }));
s.addText(`Mediana de cierre global ${nf(R.medianaCierre)} días · el hallazgo más lento tomó ${nf(R.maxCierre)} días · ` +
  `${nf(K.sinEspecialidad)} hallazgos sin especialidad declarada.`,
  { x: 0.5, y: 6.55, w: 12.3, h: 0.3, fontFace: FT, fontSize: 9.5, italic: true, color: C.ink2, margin: 0 });
footer(s, false, CAP, "consolidado");

// =====================================================================
// 4 — Levantados en la última semana
// =====================================================================
s = p.addSlide(); bg(s, C.paper); logo(s, false);
kicker(s, "Novedades del período", 0.5, 0.45);
titulo(s, "Levantados en la última semana");
bajada(s, SEM.total
  ? `Entre el ${SEM.desde} y el ${SEM.hasta} se ${SEM.total === 1 ? "levantó" : "levantaron"} ${nf(SEM.total)} ` +
    `${SEM.total === 1 ? "hallazgo" : "hallazgos"}: ${desg(SEM)}. ` +
    `${SEM.abiertas === SEM.total ? (SEM.total === 1 ? "Sigue abierto" : "Todos siguen abiertos") : `${nf(SEM.abiertas)} siguen abiertos`}.`
  : `Entre el ${SEM.desde} y el ${SEM.hasta} no se levantó ningún hallazgo en ningún frente.`);

// Ritmo: 4 hallazgos no dicen nada sin saber el ritmo previo. Las semanas
// recientes van una a una y las anteriores acumuladas; esa barra se rotula con
// las semanas que suma para que no se lea como una sola.
const ritLbl = (w) => (w.acumulada ? `${w.semanas} sem.` : w.desde);
s.addChart(p.ChartType.bar, [
  { name: "Aún abiertos", labels: SEM.tendencia.map(ritLbl), values: SEM.tendencia.map((w) => z(w.abiertas)) },
  { name: "Ya cerrados", labels: SEM.tendencia.map(ritLbl), values: SEM.tendencia.map((w) => z(w.total - w.abiertas)) },
], {
  ...BARRAS, x: 0.5, y: 1.95, w: 6.4, h: 2.05, chartColors: [C.crit, C.blue],
  title: `Ritmo semanal — la barra «${SEM.tendencia[0].semanas} sem.» acumula las semanas más viejas`,
  catAxisLabelFontSize: 10, barGapWidthPct: 40,
});

// Tarjeta con los cortes de la semana: disciplina, frente, origen y tipo.
// 1.95 -> 4.00; las cuatro filas cierran en 3.89.
s.addShape(p.ShapeType.roundRect, { x: 7.1, y: 1.95, w: 5.73, h: 2.05, rectRadius: 0.06, fill: { color: C.navy }, line: { type: "none" } });
s.addText("DE DÓNDE SALIERON", { x: 7.4, y: 2.07, w: 5, h: 0.28, fontFace: FT, fontSize: 10.5, bold: true, color: C.copperL, charSpacing: 1.5, margin: 0 });
const bloques = [
  ["Disciplina", semDisc.map(([e, v]) => `${e} ${nf(v.total)}`).join("  ·  ") || "—"],
  ["Frente", Object.entries(SEM.porProyecto).map(([e, v]) => `${e} ${nf(v.total)}`).join("  ·  ") || "—"],
  ["Origen", Object.entries(SEM.porEmision).map(([k, v]) => `${emi(k)} ${nf(v)}`).join("  ·  ") || "—"],
  ["Tipo", Object.entries(SEM.porTipo).map(([e, v]) => `${e} ${nf(v.total)}`).join("  ·  ") || "—"],
];
let by = 2.45;
bloques.forEach(([l, v]) => {
  s.addText(l.toUpperCase(), { x: 7.4, y: by, w: 1.55, h: 0.36, fontFace: FT, fontSize: 9.5,
    bold: true, color: C.steel, charSpacing: 1, margin: 0, valign: "middle" });
  s.addText(v, { x: 8.95, y: by, w: 3.75, h: 0.36, fontFace: FT, fontSize: 11, color: C.white, margin: 0, valign: "middle" });
  by += 0.36;
});

// Detalle uno a uno: es la lista que se revisa en la reunión semanal. Si la
// semana trae más de las que caben, se dice en el título: nunca se recorta en
// silencio.
const MAX_SEM = 6;
const todasSem = SEM.detalle || [];
const detSem = todasSem.slice(0, MAX_SEM);
s.addText(detSem.length
  ? "DETALLE DE LOS HALLAZGOS NUEVOS" +
    (todasSem.length > MAX_SEM ? ` — SE LISTAN ${MAX_SEM} DE ${nf(todasSem.length)}; EL RESTO ESTÁ EN EL PANEL` : "")
  : "SIN HALLAZGOS NUEVOS EN EL PERÍODO",
  { x: 0.55, y: 4.06, w: 11, h: 0.26, fontFace: FT, fontSize: 10.5, bold: true, color: C.copper, charSpacing: 1.2, margin: 0 });
if (detSem.length) {
  tabla(s, [
    { t: "FECHA", x: 0.55, w: 0.9, al: "left" },
    { t: "FRENTE", x: 1.5, w: 1.15, al: "left" },
    { t: "N°", x: 2.7, w: 1.15, al: "left" },
    { t: "TIPO", x: 3.9, w: 1.5, al: "left" },
    { t: "ORIGEN", x: 5.45, w: 1.55, al: "left" },
    { t: "DISCIPLINA", x: 7.05, w: 1.2, al: "left" },
    { t: "RESPONSABLE", x: 8.3, w: 1.5, al: "left" },
    { t: "ESTADO", x: 9.85, w: 0.95, al: "left" },
    { t: "TÍTULO", x: 10.85, w: 2.0, al: "left" },
  ], detSem.map((x) => [
    x.fecha, x.proy,
    x.codigoExterno ? `${x.n} · ${corta(x.codigoExterno, 11)}` : String(x.n || "—"),
    corta(x.tipo, 19),
    { txt: emi(x.emision), color: x.origen === "Interna" ? C.blue : C.copper, bold: true, size: 10.5 },
    corta(x.especialidad, 13), corta(x.responsable, 17),
    { txt: corta(x.estatus, 11), color: x.estado === "Abierta" ? C.crit : C.goodD, size: 10.5 },
    { txt: corta(x.titulo, 28), size: 10 },
  ]), 4.36, 0.36, { size: 10.5 });   // 6 filas desde 4.84 cierran en 6.94
} else {
  s.addText(`No se levantó ningún hallazgo entre el ${SEM.desde} y el ${SEM.hasta} en ningún frente.`,
    { x: 0.55, y: 4.45, w: 10, h: 0.4, fontFace: FT, fontSize: 13, color: C.ink2, margin: 0 });
}
footer(s, false, CAP, "consolidado");

// =====================================================================
// 5 — Origen y tipo del hallazgo
// =====================================================================
s = p.addSlide(); bg(s, C.paper); logo(s, false);
kicker(s, "Origen del hallazgo", 0.5, 0.45);
titulo(s, "Quién levanta las no conformidades");
bajada(s, `Interna es autodetección: la levanta Besalco antes que nadie. Externa la levanta el cliente ` +
  `—compromete el contrato— o un subcontrato. Autodetección global: ${pc1(G.autodeteccion)}.`);

const serieEmi = EMISIONES.filter((k) => PROY.some((x) => (x.porEmision[k] || {}).total > 0));
s.addChart(p.ChartType.bar, serieEmi.map((k) => ({
  name: emi(k), labels: PROY.map((x) => x.nombre),
  values: PROY.map((x) => z((x.porEmision[k] || {}).total || 0)),
})), {
  ...BARRAS, x: 0.5, y: 2.0, w: 6.6, h: 2.45, chartColors: serieEmi.map((k) => COL_EMI[k]),
  title: "Hallazgos por frente y vía de emisión", catAxisLabelFontSize: 12, barGapWidthPct: 55,
});

// Estado de cada vía: no basta con cuántas entran por cada una, hay que ver
// cuántas siguen abiertas y qué tan viejas son. El atraso se declara en la
// lámina siguiente —no es calculable con este archivo—, así que aquí la
// severidad de lo abierto la da la antigüedad.
s.addText(`ESTADO DE CADA VÍA  ·  atrasada = abierta que pasó los ${nf(K.plazoRespuesta)} días de plazo  ·  ` +
  `promedio en respuesta = días de emisión a cierre, solo sobre las cerradas`,
  { x: 0.55, y: 4.55, w: 7.0, h: 0.3, fontFace: FT,
    fontSize: 9, bold: true, color: C.copper, charSpacing: 0.6, margin: 0 });
const viasEmi = EMISIONES.filter((k) => (G.porEmision[k] || {}).total > 0);
tabla(s, [
  { t: "VÍA DE EMISIÓN", x: 0.55, w: 1.6, al: "left" },
  { t: "LEVANT.", x: 2.2, w: 0.8 },
  { t: "CERRADAS", x: 3.05, w: 0.8 },
  { t: "% CIERRE", x: 3.9, w: 0.8 },
  { t: "ABIERTAS", x: 4.75, w: 0.8 },
  { t: "ATRASADAS", x: 5.6, w: 0.85 },
  { t: "PROM. RESPUESTA", x: 6.35, w: 1.15 },
], viasEmi.concat(["TOTAL"]).map((k) => {
  const tot = k === "TOTAL";
  const v = tot ? R : G.porEmision[k];
  return [
    { txt: tot ? "Todas las vías" : emi(k), bold: tot, color: tot ? C.ink : COL_EMI[k] },
    { txt: nf(v.total), bold: tot },
    { txt: nf(v.cerradas), bold: tot },
    { txt: pct(v.cerradas, v.total), bold: tot },
    { txt: v.abiertas ? nf(v.abiertas) : "—", bold: tot, color: v.abiertas ? C.crit : C.ink2 },
    { txt: v.abiertas ? nf(v.atrasadas) : "—", bold: tot, color: v.atrasadas ? C.crit : C.ink2 },
    { txt: v.promedioCierre == null ? "—" : `${nf(v.promedioCierre)} días`, bold: tot },
  ];
}), 4.88, 0.31, { size: 10.5, ancho: 6.95 });   // 5 filas desde 5.36 cierran en 6.85

const tiposOrd = Object.entries(G.porTipo).sort((a, b) => b[1].total - a[1].total);
s.addChart(p.ChartType.bar, [
  { name: "Cerrados", labels: tiposOrd.map(([t]) => corta(t, 22)), values: tiposOrd.map(([, v]) => z(v.cerradas)) },
  { name: "Abiertos", labels: tiposOrd.map(([t]) => corta(t, 22)), values: tiposOrd.map(([, v]) => z(v.abiertas)) },
], {
  ...BARRAS, x: 7.3, y: 2.0, w: 5.53, h: 2.85, chartColors: [C.blue, C.crit],
  title: "Hallazgos por tipo", catAxisLabelFontSize: 10.5, barGapWidthPct: 45,
});

s.addShape(p.ShapeType.roundRect, { x: 7.3, y: 5.0, w: 5.53, h: 1.4, rectRadius: 0.06, fill: { color: C.navy }, line: { type: "none" } });
s.addText("AUTODETECCIÓN POR FRENTE", { x: 7.6, y: 5.14, w: 5, h: 0.3, fontFace: FT, fontSize: 10.5, bold: true, color: C.copperL, charSpacing: 1.2, margin: 0 });
let ax = 7.6;
OBRAS.forEach((x) => {
  s.addText(pc1(x.autodeteccion), { x: ax, y: 5.46, w: 1.6, h: 0.5, fontFace: FH, fontSize: 24, bold: true,
    color: x.autodeteccion < 30 ? C.copperL : C.white, margin: 0 });
  s.addText(x.nombre, { x: ax, y: 5.96, w: 1.6, h: 0.3, fontFace: FT, fontSize: 10.5, color: C.ice, margin: 0 });
  ax += 1.72;
});
footer(s, false, CAP, "consolidado");

// =====================================================================
// 6 — Antigüedad y velocidad de cierre
// =====================================================================
s = p.addSlide(); bg(s, C.paper); logo(s, false);
kicker(s, "Lo que sigue abierto", 0.5, 0.45);
titulo(s, "Antigüedad y velocidad de corrección");
bajada(s, `De los ${nf(R.abiertas)} hallazgos abiertos, ${nf(viejas)} llevan más de 180 días desde que se levantaron ` +
  `y ${nf(anioMas)} más de un año. La mediana de cierre global es de ${nf(R.medianaCierre)} días.`);

const COL_ANT = { "1-30 días": C.goodD, "31-90 días": C.blue, "91-180 días": C.warnD,
                  "181-365 días": C.copperD, "Más de un año": C.crit };
const tramos = Object.entries(ANT).filter(([, v]) => v > 0);
s.addChart(p.ChartType.bar, tramos.map(([t, v]) => ({
  name: t, labels: ["Abiertos"], values: [z(v)],
})), {
  ...BARRAS, x: 0.5, y: 2.05, w: 6.6, h: 2.6, chartColors: tramos.map(([t]) => COL_ANT[t]),
  title: `Antigüedad de los ${nf(R.abiertas)} hallazgos abiertos`,
  catAxisLabelFontSize: 12, barGapWidthPct: 70,
});

// La regla de atraso del proyecto y por qué no se puede aplicar: se declara,
// no se reporta 0 atrasados como si fuera un buen resultado.
s.addShape(p.ShapeType.roundRect, { x: 7.3, y: 2.05, w: 5.53, h: 2.6, rectRadius: 0.06,
  fill: { color: C.navy }, line: { type: "none" } });
s.addText(K.atrasoCalculable ? "FUERA DE PLAZO DE RESPUESTA" : "ATRASO — NO CALCULABLE",
  { x: 7.6, y: 2.22, w: 5, h: 0.3, fontFace: FT, fontSize: 11, bold: true,
    color: K.atrasoCalculable ? C.crit : C.copperL, charSpacing: 1.5, margin: 0 });
s.addText([{ text: `La regla del proyecto es: ${nf(K.plazoRespuesta)} días para responder una NC desde que se emite; ` +
                   `del día ${nf(K.plazoRespuesta + 1)} en adelante corre atraso.\n`,
             options: { color: C.ice } },
           K.atrasoCalculable
             ? { text: `Al corte hay ${nf(R.atrasadas)} hallazgos fuera de plazo de los ${nf(R.abiertas)} abiertos.`,
                 options: { color: C.white, bold: true } }
             : { text: `${nf(K.abiertasSinFechaEmision)} abiertos no traen fecha de emisión, así que no se les puede ` +
                       `calcular el atraso.`, options: { color: C.white } }],
  { x: 7.6, y: 2.58, w: 4.95, h: 1.5, fontFace: FT, fontSize: 11.5, lineSpacing: 15, margin: 0, valign: "top" });
s.addText("Se mide sobre la fecha de emisión: el Excel no declara una fecha comprometida de cierre —esa columna solo se llena al cerrar—.",
  { x: 7.6, y: 4.05, w: 4.95, h: 0.5, fontFace: FT, fontSize: 10, italic: true, color: C.steel, margin: 0, valign: "top" });

// Velocidad de cierre por frente.
s.addText("VELOCIDAD DE CIERRE POR FRENTE", { x: 0.55, y: 4.75, w: 8, h: 0.3, fontFace: FT,
  fontSize: 11, bold: true, color: C.copper, charSpacing: 1.5, margin: 0 });
const mejorMed = rapido ? rapido.resumen.medianaCierre : null;
tabla(s, [
  { t: "FRENTE", x: 0.55, w: 2.6, al: "left" },
  { t: "CERRADOS", x: 3.3, w: 1.2 },
  { t: "MEDIANA", x: 4.7, w: 1.2 },
  { t: "MÁS LENTO", x: 6.1, w: 1.3 },
  { t: "ABIERTOS", x: 7.6, w: 1.2 },
  { t: "LECTURA", x: 9.0, w: 3.85, al: "left" },
], PROY.map((x) => {
  const v = x.resumen;
  const lentoX = v.medianaCierre != null && mejorMed != null && v.medianaCierre > 3 * mejorMed;
  return [
    { txt: x.nombre, bold: true },
    nf(v.cerradas),
    v.medianaCierre == null ? "—" : `${nf(v.medianaCierre)} d`,
    v.maxCierre == null ? "—" : `${nf(v.maxCierre)} d`,
    { txt: nf(v.abiertas), color: v.abiertas > 15 ? C.crit : C.ink },
    lentoX ? { txt: `${Math.round(v.medianaCierre / mejorMed)}× más lento que el mejor frente`, color: C.crit }
           : { txt: v.abiertas === 0 ? "Sin pendientes" : "", color: C.ink2 },
  ];
}), 5.1, 0.35, { size: 11.5 });   // 4 filas desde 5.58 cierran en 6.92
footer(s, false, CAP, "consolidado");

// =====================================================================
// 7 — Foco de gestión
// =====================================================================
s = p.addSlide(); bg(s, C.navy); logo(s, true);
s.addShape(p.ShapeType.ellipse, { x: 10.8, y: -1.8, w: 4.6, h: 4.6, fill: { color: C.navy2 }, line: { type: "none" } });
kicker(s, "Foco de gestión", 0.5, 0.5, C.copperL);
s.addText("Dónde poner el esfuerzo esta semana", { x: 0.5, y: 0.78, w: 12, h: 0.6, fontFace: FH, fontSize: 30, bold: true, color: C.white, margin: 0 });

// Quién tiene la deuda antigua. Se cuenta sobre los hallazgos que superan los
// 180 días, no sobre el total de abiertos: son listas distintas.
const respViejas = Object.entries(
  D.abiertas.filter((a) => (a.dias || 0) > 180)
    .reduce((m, a) => { m[a.responsable] = (m[a.responsable] || 0) + 1; return m; }, {}))
  .sort((a, b) => b[1] - a[1]).slice(0, 3);
const bajaAuto = OBRAS.slice().sort((a, b) => a.autodeteccion - b.autodeteccion)[0];
const focus = [
  { n: "01", t: `Cerrar el pendiente de ${foco.nombre}`, c: C.crit,
    d: `${nf(foco.resumen.abiertas)} de los ${nf(R.abiertas)} hallazgos abiertos (${pct(foco.resumen.abiertas, R.abiertas)}) ` +
       `y ${pct(foco.resumen.cerradas, foco.resumen.total)} de cierre, contra ${pct(R.cerradas, R.total)} global. ` +
       `Es la ruta crítica del módulo.` },
  { n: "02", t: `Levantar la autodetección de ${bajaAuto.nombre}`, c: C.copper,
    d: `Solo ${pc1(bajaAuto.autodeteccion)} de sus hallazgos los detecta Besalco; el mejor frente va en ` +
       `${pc1(OBRAS.slice().sort((a, b) => b.autodeteccion - a.autodeteccion)[0].autodeteccion)}. ` +
       `Lo que no se detecta adentro llega como observación del cliente.` },
  { n: "03", t: `Sacar los ${nf(viejas)} hallazgos sobre 180 días`, c: C.warn,
    d: `${nf(anioMas)} de ellos superan el año abiertos. ` +
       (respViejas.length
         ? `Los tienen a cargo ${respViejas.map(([n2, c2]) => `${n2} (${nf(c2)})`).join(", ")}. `
         : "") +
       `Es deuda antigua que no se cierra sola.` },
  { n: "04", t: K.atrasoCalculable ? `Responder las ${nf(R.atrasadas)} que pasaron el plazo` : "Habilitar la medición del atraso", c: C.good,
    d: K.atrasoCalculable
      ? `El plazo es de ${nf(K.plazoRespuesta)} días desde la emisión y hoy lo incumplen ` +
        `${pct(R.atrasadas, R.abiertas)} de los abiertos. Solo ${nf(R.abiertas - R.atrasadas)} siguen dentro de plazo.`
      : `Agregar al Excel la fecha de emisión de las abiertas que no la traen. Sin ella el indicador que pide el ` +
        `procedimiento no se puede calcular y hoy se reemplaza por la antigüedad.` },
];
focus.forEach((f, i) => {
  const px = 0.5 + (i % 2) * (6.0 + 0.33);
  const py = 1.75 + Math.floor(i / 2) * 2.35;
  s.addShape(p.ShapeType.roundRect, { x: px, y: py, w: 6.0, h: 2.1, rectRadius: 0.07, fill: { color: C.navy2 }, line: { type: "none" } });
  s.addText(f.n, { x: px + 0.3, y: py + 0.22, w: 1.1, h: 0.7, fontFace: FH, fontSize: 34, bold: true, color: f.c, margin: 0 });
  s.addText(f.t, { x: px + 1.35, y: py + 0.2, w: 6.0 - 1.6, h: 0.66, fontFace: FT, fontSize: 14, bold: true, color: C.white, margin: 0, valign: "middle" });
  s.addText(f.d, { x: px + 1.35, y: py + 0.88, w: 6.0 - 1.65, h: 1.1, fontFace: FT, fontSize: 11, color: C.ice, margin: 0, lineSpacing: 14, valign: "top" });
});
footer(s, true, CAP, "consolidado");

// =====================================================================
// POR PROYECTO — portada + 2 láminas cada uno.
// Se generan en un bucle: así los tres se leen igual y no se puede
// desincronizar uno de otro.
// =====================================================================
OBRAS.forEach((X, idx) => {
  const v = X.resumen;
  const S2 = X.semana;
  const num = String(idx + 1).padStart(2, "0");
  const esp = Object.entries(X.porEspecialidad).sort((a, b) => b[1].total - a[1].total).slice(0, 8);
  const abiertasX = D.abiertas.filter((a) => a.proy === X.k);
  const antX = X.antiguedad || {};
  const viejasX = (antX["181-365 días"] || 0) + (antX["Más de un año"] || 0);

  // ---------------- portada del proyecto ----------------
  s = p.addSlide();
  // El código del proyecto acompaña al número de portada: es la misma
  // identidad —código, nombre y cliente— que usan los otros dos módulos.
  portada(s, `PROYECTO ${num}${X.codigo && X.codigo !== "—" ? " · " + X.codigo : ""}`,
    X.nombre, `${X.cliente} · ${D.meta.descripcion}`,
    `${nf(v.total)} hallazgos levantados · ${desg(v)} · ` +
    `${Object.keys(X.porEspecialidad).length} especialidades involucradas.\n` +
    `Estado de cierre, antigüedad de lo que sigue abierto y novedades de la última semana.`,
    [[pct(v.cerradas, v.total), "Hallazgos cerrados"],
     [nf(v.abiertas), "Siguen abiertos"],
     [v.medianaCierre == null ? "—" : `${nf(v.medianaCierre)} d`, "Mediana de cierre"],
     [pc1(X.autodeteccion), "Autodetección"]],
    `Corte: ${D.meta.corteTexto}`, CAP);

  // ---------------- lámina A · estado de los hallazgos ----------------
  s = p.addSlide(); bg(s, C.paper); logo(s, false);
  kicker(s, `${X.nombre} · Estado de los hallazgos`, 0.5, 0.45);
  titulo(s, "Cierre, disciplina y origen");

  tarjetas(s, [
    { p: nf(v.total), v: `${nf(v.cerradas)} cerrados`, t: "Hallazgos levantados",
      d: desg(v), c: C.blue },
    { p: pct(v.cerradas, v.total), v: `${nf(v.cerradas)}/${nf(v.total)}`, t: "Cierre de hallazgos",
      d: `Global del módulo: ${pct(R.cerradas, R.total)}`, c: C.goodD },
    { p: nf(v.abiertas), v: v.abiertas === 1 ? "hallazgo abierto" : "hallazgos abiertos",
      t: "Siguen abiertos",
      d: Object.entries(X.porEstatus || {}).map(([k, n]) => `${nf(n)} ${k.toLowerCase()}`).join(" · ") || "ninguno abierto",
      c: v.abiertas ? C.crit : C.good },
    { p: v.medianaCierre == null ? "—" : nf(v.medianaCierre), v: "días (mediana)", t: "Velocidad de cierre",
      d: v.maxCierre == null ? "sin hallazgos cerrados" : `el más lento tomó ${nf(v.maxCierre)} días`, c: C.copper },
  ]);

  s.addChart(p.ChartType.bar, [
    { name: "Cerrados", labels: esp.map(([e]) => corta(e, 18)), values: esp.map(([, x2]) => z(x2.cerradas)) },
    { name: "Abiertos", labels: esp.map(([e]) => corta(e, 18)), values: esp.map(([, x2]) => z(x2.abiertas)) },
  ], {
    ...BARRAS, x: 0.5, y: 4.25, w: 7.6, h: 2.7, chartColors: [C.blue, C.crit],
    title: `Especialidades con más hallazgos (${esp.length} de ${Object.keys(X.porEspecialidad).length})`,
    catAxisLabelFontSize: 10, barGapWidthPct: 35,
  });

  // Tarjeta de origen: 4.25 -> 6.95. La autodetección va arriba y las vías de
  // emisión debajo, porque son hasta cuatro y crecen hacia el pie.
  s.addShape(p.ShapeType.roundRect, { x: 8.35, y: 4.25, w: 4.48, h: 2.7, rectRadius: 0.06, fill: { color: C.navy }, line: { type: "none" } });
  s.addText("ORIGEN DEL HALLAZGO", { x: 8.62, y: 4.38, w: 4, h: 0.3, fontFace: FT, fontSize: 10.5, bold: true, color: C.copperL, charSpacing: 1.2, margin: 0 });
  s.addText(pc1(X.autodeteccion), { x: 8.62, y: 4.66, w: 1.3, h: 0.5, fontFace: FH, fontSize: 26,
    bold: true, color: C.white, margin: 0, valign: "middle" });
  s.addText(`de autodetección\nglobal del módulo ${pc1(G.autodeteccion)}`,
    { x: 9.95, y: 4.66, w: 2.7, h: 0.5, fontFace: FT, fontSize: 10, color: C.ice, margin: 0, lineSpacing: 12, valign: "middle" });
  s.addShape(p.ShapeType.line, { x: 8.62, y: 5.24, w: 4.0, h: 0, line: { color: C.navy2, width: 1 } });
  let oy = 5.34;
  EMISIONES.filter((k) => (X.porEmision[k] || {}).total > 0).forEach((k) => {
    const e = X.porEmision[k];
    s.addShape(p.ShapeType.ellipse, { x: 8.64, y: oy + 0.1, w: 0.18, h: 0.18, fill: { color: COL_EMI[k] }, line: { type: "none" } });
    s.addText(emi(k), { x: 8.92, y: oy, w: 2.35, h: 0.38, fontFace: FT, fontSize: 11, color: C.ice, margin: 0, valign: "middle" });
    s.addText(`${nf(e.abiertas)} ab.`, { x: 11.25, y: oy, w: 0.75, h: 0.38, fontFace: FT, fontSize: 9.5,
      color: e.abiertas ? C.copperL : C.steel, align: "right", margin: 0, valign: "middle" });
    s.addText(nf(e.total), { x: 12.0, y: oy, w: 0.65, h: 0.38, fontFace: FH, fontSize: 16, bold: true,
      color: C.white, align: "right", margin: 0, valign: "middle" });
    oy += 0.4;   // 4 vías desde 5.34 cierran en 6.94, dentro de la tarjeta (6.95)
  });
  footer(s, false, CAP, X.nombre);

  // ---------------- lámina B · abiertos y novedades ----------------
  s = p.addSlide(); bg(s, C.paper); logo(s, false);
  kicker(s, `${X.nombre} · Pendiente y novedades`, 0.5, 0.45);
  titulo(s, "Hallazgos abiertos y lo levantado en la semana");
  bajada(s, v.abiertas
    ? `${nf(v.abiertas)} ${v.abiertas === 1 ? "hallazgo sigue abierto" : "hallazgos siguen abiertos"}; ` +
      `${nf(viejasX)} con más de 180 días. Se listan del más antiguo al más reciente.`
    : `No quedan hallazgos abiertos en ${X.nombre}.`);

  const MAX_AB = 9;
  const listaAb = abiertasX.slice(0, MAX_AB);
  if (listaAb.length) {
    tabla(s, [
      { t: "N°", x: 0.55, w: 0.9, al: "left" },
      { t: "TIPO", x: 1.5, w: 1.65, al: "left" },
      { t: "ORIGEN", x: 3.2, w: 1.6, al: "left" },
      { t: "DISCIPLINA", x: 4.85, w: 1.4, al: "left" },
      { t: "RESPONSABLE", x: 6.3, w: 1.65, al: "left" },
      { t: "ESTATUS", x: 8.0, w: 1.45, al: "left" },
      { t: "DÍAS", x: 9.5, w: 0.7 },
      { t: "TÍTULO", x: 10.35, w: 2.5, al: "left" },
    ], listaAb.map((a) => [
      corta(a.n || "—", 13), corta(a.tipo, 20),
      { txt: emi(a.emision), color: a.origen === "Interna" ? C.blue : C.copper, bold: true, size: 10.5 },
      corta(a.especialidad, 16), corta(a.responsable, 19), corta(a.estatus, 18),
      { txt: nf(a.dias), color: (a.dias || 0) > 180 ? C.crit : C.ink, bold: (a.dias || 0) > 180 },
      { txt: corta(a.titulo, 36), size: 10 },
    ]), 2.05, 0.38, { size: 10.5 });   // 9 filas desde 2.53 cierran en 5.89
    if (abiertasX.length > MAX_AB)
      s.addText(`Se listan los ${MAX_AB} más antiguos de ${nf(abiertasX.length)} abiertos; el detalle completo está en el panel.`,
        { x: 0.55, y: 5.95, w: 9, h: 0.28, fontFace: FT, fontSize: 9.5, italic: true, color: C.ink2, margin: 0 });
  } else {
    s.addText("Sin hallazgos abiertos al corte.", { x: 0.55, y: 2.2, w: 10, h: 0.4,
      fontFace: FT, fontSize: 13, color: C.ink2, margin: 0 });
  }

  // Franja de la semana, al pie de la lámina.
  const yS = 6.35;
  s.addShape(p.ShapeType.roundRect, { x: 0.5, y: yS, w: 12.33, h: 0.62, rectRadius: 0.06,
    fill: { color: S2.total ? C.navy : "E9EDF2" }, line: { type: "none" } });
  s.addText(`SEMANA ${S2.desde} — ${S2.hasta}`, { x: 0.78, y: yS, w: 2.7, h: 0.62, fontFace: FT,
    fontSize: 10, bold: true, color: S2.total ? C.copperL : C.ink2, charSpacing: 1, margin: 0, valign: "middle" });
  s.addText(S2.total
    ? `${nf(S2.total)} ${S2.total === 1 ? "hallazgo nuevo" : "hallazgos nuevos"} · ` +
      `${desg(S2)} · ` +
      Object.entries(S2.porEspecialidad).map(([e, w]) => `${e} ${nf(w.total)}`).join(" · ") +
      ` · ${nf(S2.abiertas)} ${S2.abiertas === 1 ? "sigue abierto" : "siguen abiertos"}`
    : "Sin hallazgos nuevos en el período.",
    { x: 3.5, y: yS, w: 9.1, h: 0.62, fontFace: FT, fontSize: 11.5,
      color: S2.total ? C.white : C.ink2, margin: 0, valign: "middle" });
  footer(s, false, CAP, X.nombre);
});

console.log(`Láminas: ${nSlide - 1} (7 corporativas + ${OBRAS.length} proyectos × 3).`);

// =====================================================================
// Se escribe la PPT y se embebe en index.html para que el panel la ofrezca
// con el botón «Descargar Informe» sin depender de archivos sueltos.
// =====================================================================
const MARCA_INI = "<!-- === PPT:INICIO ===";
const MARCA_FIN = "<!-- === PPT:FIN === -->";

// Reemplaza un bloque marcado dentro de index.html. Devuelve el texto nuevo, o
// el original si las marcas no están (nunca deja el archivo a medias).
function reemplazar(txt, ini, fin, contenido) {
  const i = txt.indexOf(ini), f = txt.indexOf(fin);
  if (i === -1 || f === -1 || f < i) {
    console.log(`Aviso: no se encontraron las marcas ${ini.slice(9)} en index.html.`);
    return txt;
  }
  return txt.slice(0, i) + contenido + txt.slice(f);
}

// Los dos logos viajan embebidos en el panel: la versión blanca para el fondo
// oscuro y la de color para el modo claro. Se buscan igual que el de la PPT.
function bloqueLogos() {
  const uri = (f) => {
    const r = [path.join(AQUI, f), path.join(AQUI, "..", "panel_control_TOP_P1", f)]
      .find((x) => fs.existsSync(x));
    return r ? "data:image/png;base64," + fs.readFileSync(r).toString("base64") : null;
  };
  const L = { blanco: uri("besalco_logo_blanco.png"), color: uri("besalco_logo.png") };
  if (!L.blanco && !L.color) console.log("Aviso: no se encontró ningún logo para embeber.");
  return "<!-- === LOGO:INICIO === (logos embebidos por gen_ppt.js — no editar a mano) -->\n" +
    "<script>\nconst LOGOS = " + JSON.stringify(L) + ";\n<\/script>\n";
}

function embeber(archivo) {
  const html = path.join(AQUI, "index.html");
  if (!fs.existsSync(html)) {
    console.log("Aviso: no se encontró index.html; la PPT quedó solo como archivo.");
    return;
  }
  let txt = fs.readFileSync(html, "utf8");
  txt = reemplazar(txt, "<!-- === LOGO:INICIO ===", "<!-- === LOGO:FIN === -->", bloqueLogos());
  const bin = fs.readFileSync(archivo);
  // El nombre del archivo descargado lleva el corte en DDMMAA, para no pisar
  // informes de semanas distintas en la carpeta de descargas de quien lo abra.
  const [dd, mm, aaaa] = String(D.meta.corteTexto || "").split("-");
  const meta = {
    nombre: `Panel_NC_${dd && mm && aaaa ? dd + mm + aaaa.slice(2) : "actual"}.pptx`,
    corte: D.meta.corteTexto,
    laminas: nSlide - 1,
    peso: `${(bin.length / 1024 / 1024).toFixed(1)} MB`,
    b64: bin.toString("base64"),
  };
  const bloque = MARCA_INI + " (la PPT del corte, embebida por gen_ppt.js — no editar a mano) -->\n" +
    "<script>\nconst PPTX = " + JSON.stringify(meta) + ";\n<\/script>\n";
  fs.writeFileSync(html, reemplazar(txt, MARCA_INI, MARCA_FIN, bloque), "utf8");
  console.log(`Informe y logos embebidos en index.html (${meta.nombre} · ${meta.laminas} láminas · ${meta.peso}).`);
}

const salida = path.join(AQUI, "Panel_No_Conformidades.pptx");
p.writeFile({ fileName: salida })
  .then((f) => { console.log("PPT generada:", f); embeber(salida); })
  .catch((e) => { console.error(e); process.exit(1); });
