/*
 * Genera la PPT ejecutiva desde estatus_datos.json + historial.json.
 * No contiene cifras escritas a mano: todo se lee de los datos.
 *
 * Uso:  node gen_ppt.js     (correr siempre DESPUÉS de actualizar.py)
 */
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");

const AQUI = __dirname;
const D = JSON.parse(fs.readFileSync(path.join(AQUI, "estatus_datos.json"), "utf8"));
const PREV = D.previo || null;

// ---------- formato ----------
const nf = (n) => (n || 0).toLocaleString("es-CL");
const pctN = (a, b) => (b ? Math.round((1000 * a) / b) / 10 : 0);
const pct = (a, b) => pctN(a, b).toLocaleString("es-CL") + "%";
const delta = (act, prev) => {
  if (prev === undefined || prev === null) return null;
  const d = act - prev;
  return { d, txt: d === 0 ? "sin cambio" : (d > 0 ? "+" : "") + nf(d) };
};

const C = {
  navy: "12233B", navy2: "1B3252", steel: "6E8CA8", ice: "DCE6F0",
  paper: "F4F6F9", white: "FFFFFF", ink: "1A2433", ink2: "56606E",
  copper: "C87A32", copperL: "E3A55B",
  good: "2E9E5B", goodD: "1F7A44", warn: "E0A32E", crit: "C0392B",
  blue: "2E6FB5", blueL: "9DBDE0", track: "D9DFE6",
};
const FT = "Calibri", FH = "Cambria";
const CAP = `Fuente: Estatus_Resumen_General_QAQC.xlsx · Corte ${D.meta.corte_texto} · Áreas ${D.areas.join("/")}`;

const p = new pptxgen();
p.defineLayout({ name: "W", width: 13.333, height: 7.5 });
p.layout = "W";
p.author = "Control de Proyecto";
p.title = "Panel de Control — Carpetas TOP · Caminatas · Detalles P1";

const bg = (s, color) => { s.background = { color }; };
function footer(s, n, dark) {
  s.addText(CAP, { x: 0.5, y: 7.08, w: 10.6, h: 0.3, fontFace: FT, fontSize: 8.5,
    color: dark ? C.steel : C.ink2, align: "left", margin: 0 });
  s.addText(String(n), { x: 12.5, y: 7.08, w: 0.4, h: 0.3, fontFace: FT, fontSize: 9,
    color: dark ? C.steel : C.ink2, align: "right", margin: 0 });
}
const kicker = (s, txt, x, y, color) =>
  s.addText(txt.toUpperCase(), { x, y, w: 8, h: 0.28, fontFace: FT, fontSize: 11,
    bold: true, color: color || C.copper, charSpacing: 2, margin: 0, align: "left" });
const dot = (s, x, y, color, d) =>
  s.addShape(p.ShapeType.ellipse, { x, y, w: d || 0.14, h: d || 0.14,
    fill: { color }, line: { type: "none" } });

// ---------- derivados ----------
const cam = D.cam.global, top = D.ctop.global, P1 = D.dt.tipos.P1;
const dCam100 = delta(cam.c100, PREV && PREV.cam100);
const dTop = delta(top.entregadas, PREV && PREV.topEntregadas);
const dRech = delta(top.rech, PREV && PREV.topRech);
const dP1c = delta(P1.cerrados, PREV && PREV.p1Cerrados);
const dP1v = delta(P1.vencidos, PREV && PREV.p1Vencidos);
const dProg = delta(cam.prog, PREV && PREV.camProg);

let areaFoco = D.areas[0];
for (const a of D.areas)
  if (D.dt.p1Area[a].abiertos > D.dt.p1Area[areaFoco].abiertos) areaFoco = a;
const focoAb = D.dt.p1Area[areaFoco].abiertos;

const pendPorArea = {};
D.cam.pend.forEach(([a]) => { pendPorArea[a] = (pendPorArea[a] || 0) + 1; });
const areaCamFoco = Object.entries(pendPorArea).sort((a, b) => b[1] - a[1])[0] || [null, 0];

const espOrden = Object.entries(D.dt.p1Esp).sort((a, b) => b[1].abiertos - a[1].abiertos);
const espTopVenc = Object.entries(D.dt.p1Esp).sort((a, b) => b[1].vencidos - a[1].vencidos);

// =====================================================================
// 1 — Portada
// =====================================================================
let s = p.addSlide(); bg(s, C.navy);
s.addShape(p.ShapeType.ellipse, { x: 11.0, y: -1.6, w: 4.2, h: 4.2, fill: { color: C.navy2 }, line: { type: "none" } });
s.addShape(p.ShapeType.ellipse, { x: 12.1, y: 5.0, w: 3.0, h: 3.0, fill: { color: C.navy2 }, line: { type: "none" } });
dot(s, 0.55, 1.35, C.copper, 0.22);
s.addText("PANEL DE CONTROL DE AVANCE", { x: 0.85, y: 1.2, w: 10, h: 0.4, fontFace: FT,
  fontSize: 13, bold: true, color: C.copperL, charSpacing: 3, margin: 0 });
s.addText([{ text: "Carpetas TOP, Caminatas y", options: { breakLine: true } },
           { text: "Detalles de Terminación P1", options: {} }],
  { x: 0.82, y: 1.75, w: 11.6, h: 2.0, fontFace: FH, fontSize: 44, bold: true,
    color: C.white, lineSpacing: 48, margin: 0 });
s.addText("Estatus de armado y entrega de carpetas TOP · avance de caminatas 80% / 100% · detalles de construcción en condición P1. Vista ejecutiva por área y disciplina.",
  { x: 0.85, y: 3.95, w: 9.2, h: 1.0, fontFace: FT, fontSize: 15, color: C.ice, lineSpacing: 23, margin: 0 });

const kp = [
  [pct(cam.c80, cam.subs), "Caminatas 80%"],
  [pct(cam.c100, cam.subs), "Caminatas 100%"],
  [pct(top.entregadas, top.total), "Carpetas TOP entregadas"],
  [pct(P1.cerrados, P1.total), "Cierre detalles P1"],
];
let kx = 0.85;
kp.forEach(([v, l]) => {
  s.addText(v, { x: kx, y: 5.35, w: 2.75, h: 0.7, fontFace: FH, fontSize: 34, bold: true, color: C.copperL, margin: 0 });
  s.addText(l, { x: kx, y: 6.05, w: 2.75, h: 0.5, fontFace: FT, fontSize: 11.5, color: C.ice, margin: 0 });
  kx += 3.0;
});
s.addText(`Corte: ${D.meta.corte_texto}` + (PREV ? `   ·   Corte anterior: ${PREV.corte_texto}` : ""),
  { x: 0.85, y: 6.75, w: 8, h: 0.3, fontFace: FT, fontSize: 11, italic: true, color: C.steel, margin: 0 });

// =====================================================================
// 2 — Resumen ejecutivo
// =====================================================================
s = p.addSlide(); bg(s, C.paper);
kicker(s, "Resumen ejecutivo", 0.5, 0.45);
s.addText("Estado global de los tres frentes", { x: 0.5, y: 0.72, w: 12, h: 0.6, fontFace: FH, fontSize: 30, bold: true, color: C.ink, margin: 0 });

const cards = [
  { p: pct(cam.c80, cam.subs), v: `${nf(cam.c80)}/${nf(cam.subs)}`, t: "Caminatas 80% realizadas",
    d: cam.c80 === cam.subs ? "Frente cerrado en las 4 áreas" : `${nf(cam.subs - cam.c80)} pendientes`,
    c: cam.c80 === cam.subs ? C.good : C.blue, dd: null },
  { p: pct(cam.c100, cam.subs), v: `${nf(cam.c100)}/${nf(cam.subs)}`, t: "Caminatas 100% realizadas",
    d: `${nf(cam.prox)} próximas · ${nf(cam.prog)} por programar`, c: C.blue, dd: dCam100 },
  { p: pct(top.entregadas, top.total), v: `${nf(top.entregadas)}/${nf(top.total)}`, t: "Carpetas TOP entregadas",
    d: `${nf(top.aprob)} aprobadas · ${nf(top.rech)} rechazadas · ${nf(top.rev)} en revisión`, c: C.copper, dd: dTop },
  { p: pct(P1.cerrados, P1.total), v: `${nf(P1.cerrados)}/${nf(P1.total)}`, t: "Cierre de detalles P1",
    d: `${nf(P1.abiertos)} abiertos · ${nf(P1.vencidos)} vencidos`, c: C.crit, dd: dP1c },
];
let cx = 0.5; const cw = 3.0, gp = 0.13;
cards.forEach((cd) => {
  // Geometría de la tarjeta: 1.6 → 4.22. Todo el texto debe cerrar antes de 4.22.
  s.addShape(p.ShapeType.roundRect, { x: cx, y: 1.6, w: cw, h: 2.62, rectRadius: 0.08,
    fill: { color: C.white }, line: { color: "E4E9EF", width: 1 },
    shadow: { type: "outer", color: "9AA7B5", blur: 7, offset: 2, angle: 90, opacity: 0.28 } });
  dot(s, cx + 0.28, 1.88, cd.c, 0.16);
  s.addText(cd.p, { x: cx + 0.28, y: 2.12, w: cw - 0.5, h: 0.82, fontFace: FH, fontSize: 40, bold: true, color: cd.c, margin: 0 });
  s.addText(cd.v, { x: cx + 0.3, y: 2.95, w: cw - 0.5, h: 0.32, fontFace: FT, fontSize: 14, bold: true, color: C.ink, margin: 0 });
  s.addText(cd.t, { x: cx + 0.3, y: 3.28, w: cw - 0.55, h: 0.32, fontFace: FT, fontSize: 12, color: C.ink2, margin: 0 });
  s.addText(cd.d, { x: cx + 0.3, y: 3.60, w: cw - 0.55, h: 0.28, fontFace: FT, fontSize: 9.5, color: C.ink2, margin: 0 });
  if (cd.dd) {
    const up = cd.dd.d > 0, cero = cd.dd.d === 0;
    s.addText(`${cero ? "=" : up ? "▲" : "▼"} ${cd.dd.txt} vs. corte anterior`,
      { x: cx + 0.3, y: 3.89, w: cw - 0.55, h: 0.26, fontFace: FT, fontSize: 9,
        bold: true, color: cero ? C.ink2 : up ? C.goodD : C.crit, margin: 0 });
  }
  cx += cw + gp;
});

s.addShape(p.ShapeType.roundRect, { x: 0.5, y: 4.42, w: 12.33, h: 2.3, rectRadius: 0.06, fill: { color: C.navy }, line: { type: "none" } });
s.addText("LECTURA DEL PERÍODO", { x: 0.85, y: 4.6, w: 6, h: 0.3, fontFace: FT, fontSize: 11, bold: true, color: C.copperL, charSpacing: 2, margin: 0 });

const reads = [];
if (dProg && dProg.d < 0) {
  reads.push(["Se destrabó la programación de caminatas, pero no la ejecución.",
    `Los subsistemas «por programar» bajaron de ${nf(PREV.camProg)} a ${nf(cam.prog)}, ` +
    `mientras las realizadas al 100% se mantienen en ${nf(cam.c100)} (${pct(cam.c100, cam.subs)}).`]);
} else {
  reads.push(["Avance de caminatas al 100%.",
    `${nf(cam.c100)} de ${nf(cam.subs)} realizadas (${pct(cam.c100, cam.subs)}); ` +
    `${nf(cam.prox)} próximas y ${nf(cam.prog)} por programar.`]);
}
reads.push(["Las carpetas TOP siguen sin aprobaciones" + (dRech && dRech.d > 0 ? " y suben los rechazos." : "."),
  `${nf(top.entregadas)} entregadas de ${nf(top.total)} (${pct(top.entregadas, top.total)}) y ${nf(top.aprob)} aprobadas` +
  (dRech && dRech.d > 0 ? `; las rechazadas pasaron de ${nf(PREV.topRech)} a ${nf(top.rech)}.` : `; ${nf(top.rech)} rechazadas.`)]);
reads.push(["El cierre de detalles P1 está prácticamente detenido y crece la deuda vencida.",
  (dP1c ? `${dP1c.d === 0 ? "Ningún" : dP1c.txt} P1 cerrado desde el corte anterior; ` : "") +
  `${nf(P1.abiertos)} abiertos, de los cuales ${nf(P1.vencidos)} superaron su fecha de vencimiento` +
  (dP1v && dP1v.d > 0 ? ` (${dP1v.txt} en el período)` : "") +
  `. El área ${areaFoco} concentra ${nf(focoAb)} (${pct(focoAb, P1.abiertos)}).`]);

let ry = 4.96;
reads.forEach(([h, d]) => {
  dot(s, 0.9, ry + 0.06, C.copper, 0.12);
  s.addText([{ text: h + "  ", options: { bold: true, color: C.white } },
             { text: d, options: { color: C.ice } }],
    { x: 1.15, y: ry - 0.05, w: 11.4, h: 0.55, fontFace: FT, fontSize: 12.5, lineSpacing: 16, margin: 0, valign: "top" });
  ry += 0.56;
});
footer(s, 2, false);

// =====================================================================
// 3 — Semáforo por área
// =====================================================================
s = p.addSlide(); bg(s, C.paper);
kicker(s, "Estado por área", 0.5, 0.45);
s.addText("Semáforo integrado de los tres frentes", { x: 0.5, y: 0.72, w: 12, h: 0.6, fontFace: FH, fontSize: 30, bold: true, color: C.ink, margin: 0 });

const cols = [
  { t: "ÁREA", x: 0.55, w: 3.2, al: "left" },
  { t: "SUBSIST.", x: 3.8, w: 1.0, al: "right" },
  { t: "CAMINATA 100%", x: 4.85, w: 1.7, al: "right" },
  { t: "TOP ENTREGADAS", x: 6.6, w: 1.9, al: "right" },
  { t: "P1 % CIERRE", x: 8.55, w: 1.4, al: "right" },
  { t: "P1 ABIERTOS (VENC.)", x: 9.95, w: 1.9, al: "right" },
  { t: "ESTADO", x: 11.9, w: 0.95, al: "center" },
];
cols.forEach((c) => s.addText(c.t, { x: c.x, y: 1.72, w: c.w, h: 0.3, fontFace: FT,
  fontSize: 9.5, bold: true, color: C.ink2, align: c.al, charSpacing: 0.5, margin: 0 }));
s.addShape(p.ShapeType.line, { x: 0.55, y: 2.06, w: 12.3, h: 0, line: { color: C.steel, width: 1 } });

const COLOR_EST = { good: C.good, warn: C.warn, crit: C.crit };
let rowy = 2.2; const rh = 1.02;
D.areas.forEach((a, i) => {
  const c = D.cam.byArea[a], t = D.ctop.byArea[a], pp = D.dt.p1Area[a];
  const ent = t.rev + t.obs + t.rech + t.aprob;
  const est = D.semaforo[a] || ["warn", "Atención"];
  if (i % 2 === 1) s.addShape(p.ShapeType.rect, { x: 0.5, y: rowy - 0.05, w: 12.35, h: rh, fill: { color: C.white }, line: { type: "none" } });
  s.addText(a, { x: cols[0].x, y: rowy, w: 2.9, h: 0.35, fontFace: FT, fontSize: 15, bold: true, color: C.ink, margin: 0 });
  s.addText(D.areaNames[a], { x: cols[0].x, y: rowy + 0.36, w: 3.1, h: 0.3, fontFace: FT, fontSize: 10, color: C.ink2, margin: 0 });
  const cy = rowy + 0.14;
  const celda = (idx, txt) => s.addText(txt, { x: cols[idx].x, y: cy, w: cols[idx].w, h: 0.35,
    fontFace: FT, fontSize: 13, color: C.ink, align: "right", margin: 0 });
  celda(1, nf(c.subs));
  celda(2, `${nf(c.c100)}/${nf(c.subs)} · ${pct(c.c100, c.subs)}`);
  celda(3, `${nf(ent)}/${nf(t.total)} · ${pct(ent, t.total)}`);
  celda(4, pct(pp.cerrados, pp.total));
  celda(5, `${nf(pp.abiertos)} (${nf(pp.vencidos)})`);
  s.addShape(p.ShapeType.roundRect, { x: cols[6].x - 0.02, y: rowy + 0.12, w: 1.0, h: 0.42,
    rectRadius: 0.21, fill: { color: COLOR_EST[est[0]] }, line: { type: "none" } });
  s.addText(est[1], { x: cols[6].x - 0.02, y: rowy + 0.13, w: 1.0, h: 0.4, fontFace: FT,
    fontSize: 10.5, bold: true, color: C.white, align: "center", margin: 0 });
  s.addShape(p.ShapeType.line, { x: 0.55, y: rowy + rh - 0.05, w: 12.3, h: 0, line: { color: "D9DFE6", width: 0.75 } });
  rowy += rh;
});
s.addText("Criterio · Crítico: caminata 100% < 60% o cierre P1 < 50%   ·   Atención: algún frente entre 60–90% o carpetas TOP por entregar   ·   Al día: caminatas ≥ 90% y P1 ≥ 90%",
  { x: 0.5, y: 6.5, w: 12.3, h: 0.3, fontFace: FT, fontSize: 9.5, italic: true, color: C.ink2, margin: 0 });
footer(s, 3, false);

// =====================================================================
// 4 — Caminatas
// =====================================================================
s = p.addSlide(); bg(s, C.paper);
kicker(s, "Frente 1 · Caminatas", 0.5, 0.45);
s.addText("Avance de caminatas 80% y 100%", { x: 0.5, y: 0.72, w: 12, h: 0.6, fontFace: FH, fontSize: 30, bold: true, color: C.ink, margin: 0 });
s.addText(`Las caminatas al 80% están cerradas en todas las áreas (${nf(cam.c80)}/${nf(cam.subs)}). ` +
  `El hito 100% avanza a ${pct(cam.c100, cam.subs)}` +
  (areaCamFoco[0] ? `, y el área ${areaCamFoco[0]} concentra ${nf(areaCamFoco[1])} de los ${nf(D.cam.pend.length)} subsistemas pendientes.` : "."),
  { x: 0.5, y: 1.35, w: 12.3, h: 0.5, fontFace: FT, fontSize: 13.5, color: C.ink2, margin: 0 });

s.addChart(p.ChartType.bar, [
  { name: "Realizada", labels: D.areas, values: D.areas.map((a) => D.cam.byArea[a].c100) },
  { name: "Próxima a realizar", labels: D.areas, values: D.areas.map((a) => D.cam.byArea[a].prox) },
  { name: "Por programar", labels: D.areas, values: D.areas.map((a) => D.cam.byArea[a].prog) },
], {
  x: 0.5, y: 2.05, w: 7.6, h: 4.5, barDir: "bar", barGrouping: "stacked",
  chartColors: [C.goodD, C.blue, C.blueL],
  showTitle: true, title: "Subsistemas por estado de caminata 100%", titleFontFace: FT, titleFontSize: 13, titleColor: C.ink,
  showValue: true, dataLabelPosition: "ctr", dataLabelFontFace: FT, dataLabelFontSize: 11, dataLabelColor: C.white, dataLabelFontBold: true,
  showLegend: true, legendPos: "b", legendFontFace: FT, legendFontSize: 11, legendColor: C.ink2,
  catAxisLabelFontFace: FT, catAxisLabelFontSize: 13, catAxisLabelColor: C.ink, catAxisLabelFontBold: true,
  valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
  barGapWidthPct: 55, chartArea: { fill: { color: C.white } },
});

s.addShape(p.ShapeType.roundRect, { x: 8.35, y: 2.05, w: 4.48, h: 4.5, rectRadius: 0.06, fill: { color: C.navy }, line: { type: "none" } });
s.addText("HITO 80%", { x: 8.65, y: 2.3, w: 4, h: 0.3, fontFace: FT, fontSize: 11, bold: true, color: C.copperL, charSpacing: 1.5, margin: 0 });
s.addText(pct(cam.c80, cam.subs), { x: 8.65, y: 2.55, w: 2.6, h: 0.75, fontFace: FH, fontSize: 40, bold: true, color: C.white, margin: 0 });
s.addText(`${nf(cam.c80)} de ${nf(cam.subs)} subsistemas`, { x: 8.65, y: 3.32, w: 4, h: 0.3, fontFace: FT, fontSize: 12, color: C.ice, margin: 0 });
s.addShape(p.ShapeType.line, { x: 8.65, y: 3.85, w: 3.9, h: 0, line: { color: C.navy2, width: 1 } });
s.addText("HITO 100%", { x: 8.65, y: 4.02, w: 4, h: 0.3, fontFace: FT, fontSize: 11, bold: true, color: C.copperL, charSpacing: 1.5, margin: 0 });
s.addText(pct(cam.c100, cam.subs), { x: 8.65, y: 4.27, w: 2.6, h: 0.75, fontFace: FH, fontSize: 40, bold: true, color: C.white, margin: 0 });
s.addText(`${nf(cam.c100)} realizadas · ${nf(D.cam.pend.length)} pendientes`, { x: 8.65, y: 5.04, w: 4, h: 0.3, fontFace: FT, fontSize: 12, color: C.ice, margin: 0 });
s.addText([{ text: "Foco: ", options: { bold: true, color: C.copperL } },
           { text: `${nf(cam.prox)} próximas a realizar y ${nf(cam.prog)} por programar` +
             (dProg && dProg.d < 0 ? `. Las «por programar» bajaron ${nf(Math.abs(dProg.d))} en el período.` : "."),
             options: { color: C.ice } }],
  { x: 8.65, y: 5.55, w: 3.95, h: 0.9, fontFace: FT, fontSize: 11.5, lineSpacing: 15, margin: 0, valign: "top" });
footer(s, 4, false);

// =====================================================================
// 5 — Carpetas TOP
// =====================================================================
s = p.addSlide(); bg(s, C.paper);
kicker(s, "Frente 2 · Carpetas TOP", 0.5, 0.45);
s.addText("Estatus de armado y entrega", { x: 0.5, y: 0.72, w: 12, h: 0.6, fontFace: FH, fontSize: 30, bold: true, color: C.ink, margin: 0 });
s.addText(`${nf(top.entregadas)} carpetas entregadas de ${nf(top.total)} (${pct(top.entregadas, top.total)}). ` +
  `${top.aprob === 0 ? "Ninguna aprobada aún" : `${nf(top.aprob)} aprobadas`}: el flujo se acumula en revisión y hay ` +
  `${nf(top.rech)} rechazos que requieren rearmado.`,
  { x: 0.5, y: 1.35, w: 12.3, h: 0.5, fontFace: FT, fontSize: 13.5, color: C.ink2, margin: 0 });

const serieTop = [
  { name: "Aprobada", key: "aprob", color: C.good },
  { name: "En revisión", key: "rev", color: C.blue },
  { name: "Observada", key: "obs", color: C.warn },
  { name: "Rechazada", key: "rech", color: C.crit },
].filter((x) => D.areas.some((a) => D.ctop.byArea[a][x.key] > 0));
serieTop.push({ name: "Sin entregar", key: "_pend", color: C.track });

s.addChart(p.ChartType.bar, serieTop.map((x) => ({
  name: x.name, labels: D.areas,
  values: D.areas.map((a) => {
    const t = D.ctop.byArea[a];
    return x.key === "_pend" ? t.total - (t.rev + t.obs + t.rech + t.aprob) : t[x.key];
  }),
})), {
  x: 0.5, y: 2.05, w: 7.6, h: 4.5, barDir: "bar", barGrouping: "stacked",
  chartColors: serieTop.map((x) => x.color),
  showTitle: true, title: "Carpetas TOP por área y estatus", titleFontFace: FT, titleFontSize: 13, titleColor: C.ink,
  showValue: true, dataLabelPosition: "ctr", dataLabelFontFace: FT, dataLabelFontSize: 10, dataLabelColor: C.white, dataLabelFontBold: true,
  showLegend: true, legendPos: "b", legendFontFace: FT, legendFontSize: 11, legendColor: C.ink2,
  catAxisLabelFontFace: FT, catAxisLabelFontSize: 13, catAxisLabelColor: C.ink, catAxisLabelFontBold: true,
  valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
  barGapWidthPct: 55, chartArea: { fill: { color: C.white } },
});

s.addShape(p.ShapeType.roundRect, { x: 8.35, y: 2.05, w: 4.48, h: 4.5, rectRadius: 0.06, fill: { color: C.white }, line: { color: "E4E9EF", width: 1 } });
s.addText(`EMBUDO DE ENTREGA (${nf(top.total)})`, { x: 8.6, y: 2.28, w: 4, h: 0.3, fontFace: FT, fontSize: 11, bold: true, color: C.copper, charSpacing: 1, margin: 0 });
const funnel = [
  ["Entregadas", top.entregadas, C.blue, dTop, false],
  ["En revisión", top.rev, C.steel, delta(top.rev, PREV && PREV.topRev), false],
  ["Observadas", top.obs, C.warn, delta(top.obs, PREV && PREV.topObs), true],
  ["Rechazadas", top.rech, C.crit, dRech, true],
  ["Aprobadas", top.aprob, C.good, delta(top.aprob, PREV && PREV.topAprob), false],
];
let fy = 2.68;
funnel.forEach(([l, v, c, dd, subirEsMalo]) => {
  s.addShape(p.ShapeType.roundRect, { x: 8.6, y: fy, w: 3.98, h: 0.62, rectRadius: 0.05, fill: { color: C.paper }, line: { type: "none" } });
  s.addShape(p.ShapeType.ellipse, { x: 8.72, y: fy + 0.19, w: 0.24, h: 0.24, fill: { color: c }, line: { type: "none" } });
  s.addText(l, { x: 9.08, y: fy, w: 1.75, h: 0.62, fontFace: FT, fontSize: 12.5, color: C.ink, valign: "middle", margin: 0 });
  if (dd && dd.d !== 0) {
    const malo = dd.d > 0 ? subirEsMalo : !subirEsMalo;
    s.addText(`${dd.d > 0 ? "▲" : "▼"} ${dd.txt}`, { x: 10.4, y: fy, w: 1.0, h: 0.62, fontFace: FT,
      fontSize: 9.5, bold: true, color: malo ? C.crit : C.goodD, align: "right", valign: "middle", margin: 0 });
  }
  s.addText(nf(v), { x: 11.45, y: fy, w: 0.95, h: 0.62, fontFace: FH, fontSize: 22, bold: true, color: c, align: "right", valign: "middle", margin: 0 });
  fy += 0.68;   // 5 filas desde 2.68 cierran en 6.02, dentro de la tarjeta (6.55)
});
s.addText("Cuello de botella: la revisión del cliente devuelve más carpetas de las que aprueba. Convertir entregas en aprobaciones es el próximo hito.",
  { x: 8.6, y: 6.08, w: 4, h: 0.44, fontFace: FT, fontSize: 10, italic: true, color: C.ink2, margin: 0 });
footer(s, 5, false);

// =====================================================================
// 6 — Detalles P1
// =====================================================================
s = p.addSlide(); bg(s, C.paper);
kicker(s, "Frente 3 · Detalles P1", 0.5, 0.45);
s.addText("Detalles de terminación en condición P1", { x: 0.5, y: 0.72, w: 12.3, h: 0.6, fontFace: FH, fontSize: 30, bold: true, color: C.ink, margin: 0 });
s.addText(`${nf(P1.total)} P1 levantados; ${pct(P1.cerrados, P1.total)} cerrados. ` +
  `El pendiente se concentra en ${espOrden.slice(0, 3).map((x) => x[0]).join(", ")}, y físicamente en el área ${areaFoco}.`,
  { x: 0.5, y: 1.35, w: 12.3, h: 0.5, fontFace: FT, fontSize: 13.5, color: C.ink2, margin: 0 });

const espLbl = espOrden.map(([e]) => (e === "Instrumentación y Control" ? "Instrum. y Control" : e));
s.addChart(p.ChartType.bar, [
  { name: "Vencidos al corte", labels: espLbl, values: espOrden.map(([, v]) => v.vencidos) },
  { name: "En plazo", labels: espLbl, values: espOrden.map(([, v]) => v.abiertos - v.vencidos) },
], {
  x: 0.5, y: 2.05, w: 7.7, h: 4.5, barDir: "bar", barGrouping: "stacked",
  chartColors: [C.crit, C.blue],
  showTitle: true, title: "P1 abiertos por disciplina", titleFontFace: FT, titleFontSize: 13, titleColor: C.ink,
  showValue: true, dataLabelPosition: "ctr", dataLabelFontFace: FT, dataLabelFontSize: 10, dataLabelColor: C.white, dataLabelFontBold: true,
  showLegend: true, legendPos: "b", legendFontFace: FT, legendFontSize: 11, legendColor: C.ink2,
  catAxisLabelFontFace: FT, catAxisLabelFontSize: 11.5, catAxisLabelColor: C.ink, catAxisLabelFontBold: true,
  valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
  barGapWidthPct: 45, chartArea: { fill: { color: C.white } },
});

s.addShape(p.ShapeType.roundRect, { x: 8.4, y: 2.05, w: 4.43, h: 4.5, rectRadius: 0.06, fill: { color: C.navy }, line: { type: "none" } });
s.addText("ESTADO GLOBAL P1", { x: 8.68, y: 2.28, w: 4, h: 0.3, fontFace: FT, fontSize: 11, bold: true, color: C.copperL, charSpacing: 1.2, margin: 0 });
s.addChart(p.ChartType.doughnut, [{ name: "P1", labels: ["Cerrados", "Abiertos", "En trámite"],
  values: [P1.cerrados, P1.abiertos, P1.tramite] }], {
  x: 8.5, y: 2.55, w: 2.2, h: 2.2, holeSize: 62,
  chartColors: [C.good, C.crit, C.copperL],
  showTitle: false, showLegend: false, showValue: false,
  dataBorder: { pct: 1, color: C.navy }, chartArea: { fill: { color: C.navy } },
});
s.addText([{ text: pct(P1.cerrados, P1.total), options: { fontSize: 19, bold: true, color: C.white, breakLine: true } },
           { text: "cerrado", options: { fontSize: 10, color: C.ice } }],
  { x: 8.5, y: 3.28, w: 2.2, h: 0.75, align: "center", valign: "middle", fontFace: FH, margin: 0 });
const dl = [["Cerrados", P1.cerrados, C.good], ["Abiertos", P1.abiertos, C.crit], ["En trámite", P1.tramite, C.copperL]];
let dly = 2.7;
dl.forEach(([l, v, c]) => {
  s.addShape(p.ShapeType.ellipse, { x: 10.95, y: dly + 0.04, w: 0.2, h: 0.2, fill: { color: c }, line: { type: "none" } });
  s.addText(l, { x: 11.2, y: dly - 0.05, w: 1.4, h: 0.35, fontFace: FT, fontSize: 11, color: C.ice, margin: 0 });
  s.addText(nf(v), { x: 11.2, y: dly + 0.2, w: 1.4, h: 0.3, fontFace: FT, fontSize: 11, bold: true, color: C.white, margin: 0 });
  dly += 0.62;
});
s.addShape(p.ShapeType.line, { x: 8.68, y: 4.95, w: 3.9, h: 0, line: { color: C.navy2, width: 1 } });
const casiCerradas = D.areas.filter((a) => pctN(D.dt.p1Area[a].cerrados, D.dt.p1Area[a].total) >= 90);
s.addText([{ text: "Concentración por área\n", options: { bold: true, color: C.copperL, fontSize: 11.5 } },
           { text: `${nf(focoAb)} de ${nf(P1.abiertos)} P1 abiertos (${pct(focoAb, P1.abiertos)}) están en el área ${areaFoco}.` +
             (casiCerradas.length ? ` Áreas ${casiCerradas.join(" y ")} sobre 90% de cierre.` : ""),
             options: { color: C.ice, fontSize: 11.5 } }],
  { x: 8.68, y: 5.1, w: 3.95, h: 1.3, fontFace: FT, lineSpacing: 15, margin: 0, valign: "top" });
footer(s, 6, true);

// =====================================================================
// 7 — Foco de gestión
// =====================================================================
s = p.addSlide(); bg(s, C.navy);
s.addShape(p.ShapeType.ellipse, { x: 10.8, y: -1.8, w: 4.6, h: 4.6, fill: { color: C.navy2 }, line: { type: "none" } });
kicker(s, "Foco de gestión", 0.5, 0.5, C.copperL);
s.addText("Dónde poner el esfuerzo esta semana", { x: 0.5, y: 0.78, w: 12, h: 0.6, fontFace: FH, fontSize: 30, bold: true, color: C.white, margin: 0 });

const focus = [
  { n: "01", t: "Destrabar aprobaciones de carpetas TOP", c: C.copper,
    d: `${nf(top.aprob)} aprobadas de ${nf(top.total)}. Priorizar el cierre de las ${nf(top.rev)} en revisión y ` +
       `el rearmado de las ${nf(top.rech)} rechazadas para convertir entregas en aprobaciones.` },
  { n: "02", t: `Atacar el atraso del área ${areaFoco}`, c: C.crit,
    d: `${nf(focoAb)} P1 abiertos (${pct(focoAb, P1.abiertos)} del total) y ${nf(pendPorArea[areaFoco] || 0)} caminatas 100% pendientes. ` +
       `Es la ruta crítica: sin esta área no mejora ningún indicador global.` },
  { n: "03", t: `Recuperar los ${nf(P1.vencidos)} P1 vencidos`, c: C.warn,
    d: `${pct(P1.vencidos, P1.abiertos)} de los P1 abiertos superó su fecha de vencimiento. Foco en ` +
       espTopVenc.slice(0, 3).map(([e, v]) => `${e} (${nf(v.vencidos)})`).join(", ") + "." },
  { n: "04", t: "Ejecutar las caminatas ya programadas", c: C.good,
    d: `${nf(D.cam.pend.length)} subsistemas pendientes: ${nf(cam.prox)} próximos a realizar y ${nf(cam.prog)} por programar. ` +
       `La programación ya está hecha; falta ejecutarla para habilitar la entrega de carpetas.` },
];
focus.forEach((f, i) => {
  const px = 0.5 + (i % 2) * (6.0 + 0.33);
  const py = 1.75 + Math.floor(i / 2) * 2.35;
  s.addShape(p.ShapeType.roundRect, { x: px, y: py, w: 6.0, h: 2.1, rectRadius: 0.07, fill: { color: C.navy2 }, line: { type: "none" } });
  s.addText(f.n, { x: px + 0.3, y: py + 0.22, w: 1.1, h: 0.7, fontFace: FH, fontSize: 34, bold: true, color: f.c, margin: 0 });
  s.addText(f.t, { x: px + 1.35, y: py + 0.24, w: 6.0 - 1.6, h: 0.6, fontFace: FT, fontSize: 15, bold: true, color: C.white, margin: 0, valign: "middle" });
  s.addText(f.d, { x: px + 1.35, y: py + 0.88, w: 6.0 - 1.65, h: 1.1, fontFace: FT, fontSize: 11.5, color: C.ice, margin: 0, lineSpacing: 15, valign: "top" });
});
footer(s, 7, true);

const salida = path.join(AQUI, "Panel_Control_TOP_P1.pptx");
p.writeFile({ fileName: salida })
  .then((f) => console.log("PPT generada:", f))
  .catch((e) => { console.error(e); process.exit(1); });
