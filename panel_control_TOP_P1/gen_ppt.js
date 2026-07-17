const pptxgen = require("pptxgenjs");
const p = new pptxgen();
p.defineLayout({ name: "W", width: 13.333, height: 7.5 });
p.layout = "W";
p.author = "Control de Proyecto";
p.title = "Panel de Control — Carpetas TOP · Caminatas · Detalles P1";

// ---------- paleta (concentradora de cobre) ----------
const C = {
  navy:   "12233B",   // fondo oscuro
  navy2:  "1B3252",
  steel:  "6E8CA8",
  ice:    "DCE6F0",
  paper:  "F4F6F9",
  white:  "FFFFFF",
  ink:    "1A2433",
  ink2:   "56606E",
  copper: "C87A32",   // acento
  copperL:"E3A55B",
  good:   "2E9E5B",
  goodD:  "1F7A44",
  warn:   "E0A32E",
  crit:   "C0392B",
  blue:   "2E6FB5",
  blueL:  "9DBDE0",
  track:  "D9DFE6",
};
const FT = "Calibri", FH = "Cambria";

const CAP = "Fuente: Estatus.xlsx · Corte 16-jul-2026 · Áreas 2000/3000/4000/8000";

// ---------- helpers ----------
function bg(s, color){ s.background = { color }; }
function footer(s, n, dark){
  s.addText(CAP, { x:0.5, y:7.08, w:9.8, h:0.3, fontFace:FT, fontSize:8.5,
    color: dark ? C.steel : C.ink2, align:"left", margin:0 });
  s.addText(String(n), { x:12.5, y:7.08, w:0.4, h:0.3, fontFace:FT, fontSize:9,
    color: dark ? C.steel : C.ink2, align:"right", margin:0 });
}
function kicker(s, txt, x, y, color){
  s.addText(txt.toUpperCase(), { x, y, w:8, h:0.28, fontFace:FT, fontSize:11, bold:true,
    color: color||C.copper, charSpacing:2, margin:0, align:"left" });
}

// dot marker (motif: círculo de acento)
function dot(s, x, y, color, d){ d=d||0.14; s.addShape(p.ShapeType.ellipse, { x, y, w:d, h:d, fill:{color}, line:{type:"none"} }); }

// =====================================================================
// SLIDE 1 — Portada
// =====================================================================
let s = p.addSlide(); bg(s, C.navy);
s.addShape(p.ShapeType.rect, { x:0, y:0, w:13.333, h:7.5, fill:{ color:C.navy }, line:{type:"none"} });
// bloque de acento sutil (círculos motif, no barras)
s.addShape(p.ShapeType.ellipse, { x:11.0, y:-1.6, w:4.2, h:4.2, fill:{ color:C.navy2 }, line:{type:"none"} });
s.addShape(p.ShapeType.ellipse, { x:12.1, y:5.0, w:3.0, h:3.0, fill:{ color:C.navy2 }, line:{type:"none"} });
dot(s, 0.55, 1.35, C.copper, 0.22);
s.addText("PANEL DE CONTROL DE AVANCE", { x:0.85, y:1.2, w:10, h:0.4, fontFace:FT, fontSize:13, bold:true, color:C.copperL, charSpacing:3, margin:0 });
s.addText([
  { text:"Carpetas TOP, Caminatas y", options:{ breakLine:true } },
  { text:"Detalles de Terminación P1", options:{} },
], { x:0.82, y:1.75, w:11.6, h:2.0, fontFace:FH, fontSize:44, bold:true, color:C.white, lineSpacing:48, margin:0 });
s.addText("Estatus de armado y entrega de carpetas TOP · avance de caminatas 80% / 100% · detalles de construcción en condición P1. Vista ejecutiva por área y disciplina.",
  { x:0.85, y:3.95, w:9.2, h:1.0, fontFace:FT, fontSize:15, color:C.ice, lineSpacing:23, margin:0 });

// mini-KPI strip en portada
const kp = [["100%","Caminatas 80%"],["66,7%","Caminatas 100%"],["19,3%","Carpetas TOP entregadas"],["57,7%","Cierre detalles P1"]];
let kx = 0.85;
kp.forEach(([v,l])=>{
  s.addText(v, { x:kx, y:5.35, w:2.75, h:0.7, fontFace:FH, fontSize:34, bold:true, color:C.copperL, margin:0, align:"left" });
  s.addText(l, { x:kx, y:6.05, w:2.75, h:0.5, fontFace:FT, fontSize:11.5, color:C.ice, margin:0, align:"left" });
  kx += 3.0;
});
s.addText("Corte: 16 de julio de 2026", { x:0.85, y:6.75, w:6, h:0.3, fontFace:FT, fontSize:11, italic:true, color:C.steel, margin:0 });

// =====================================================================
// SLIDE 2 — Resumen ejecutivo (KPIs + lectura)
// =====================================================================
s = p.addSlide(); bg(s, C.paper);
kicker(s, "Resumen ejecutivo", 0.5, 0.45);
s.addText("Estado global de los tres frentes", { x:0.5, y:0.72, w:12, h:0.6, fontFace:FH, fontSize:30, bold:true, color:C.ink, margin:0 });

const cards = [
  { v:"135/135", p:"100%", t:"Caminatas 80% realizadas", d:"Frente cerrado en las 4 áreas", c:C.good },
  { v:"90/135", p:"66,7%", t:"Caminatas 100% realizadas", d:"45 pendientes · 41 en área 3000", c:C.blue },
  { v:"26/135", p:"19,3%", t:"Carpetas TOP entregadas", d:"0 aprobadas · 8 rechazadas · 17 en revisión", c:C.copper },
  { v:"1.147/1.988", p:"57,7%", t:"Cierre de detalles P1", d:"830 abiertos · 401 vencidos al corte", c:C.crit },
];
let cx = 0.5; const cw = 3.0, gap = 0.13;
cards.forEach(cd=>{
  s.addShape(p.ShapeType.roundRect, { x:cx, y:1.6, w:cw, h:2.35, rectRadius:0.08,
    fill:{ color:C.white }, line:{ color:"E4E9EF", width:1 }, shadow:{ type:"outer", color:"9AA7B5", blur:7, offset:2, angle:90, opacity:0.28 } });
  dot(s, cx+0.28, 1.9, cd.c, 0.16);
  s.addText(cd.p, { x:cx+0.28, y:2.15, w:cw-0.5, h:0.85, fontFace:FH, fontSize:40, bold:true, color:cd.c, margin:0, align:"left" });
  s.addText(cd.v, { x:cx+0.3, y:3.0, w:cw-0.5, h:0.35, fontFace:FT, fontSize:14, bold:true, color:C.ink, margin:0, align:"left" });
  s.addText(cd.t, { x:cx+0.3, y:3.36, w:cw-0.55, h:0.35, fontFace:FT, fontSize:12, color:C.ink2, margin:0, align:"left" });
  s.addText(cd.d, { x:cx+0.3, y:3.66, w:cw-0.55, h:0.28, fontFace:FT, fontSize:9.5, color:C.ink2, margin:0, align:"left" });
  cx += cw + gap;
});

// lectura ejecutiva
s.addShape(p.ShapeType.roundRect, { x:0.5, y:4.25, w:12.33, h:2.45, rectRadius:0.06, fill:{ color:C.navy }, line:{type:"none"} });
s.addText("LECTURA EJECUTIVA", { x:0.85, y:4.5, w:6, h:0.3, fontFace:FT, fontSize:11, bold:true, color:C.copperL, charSpacing:2, margin:0 });
const reads = [
  ["Caminatas 80% cerradas; el 100% avanza a dos velocidades.", "Área 3000 explica 41 de los 45 subsistemas pendientes de caminata al 100%."],
  ["La entrega de carpetas TOP es el frente más rezagado.", "Solo 19,3% entregadas y ninguna aprobada aún; 8 rechazadas concentradas en áreas 2000 y 3000."],
  ["El cierre de detalles P1 está a mitad de camino, con deuda vencida.", "401 de 830 P1 abiertos ya superaron su fecha de vencimiento; el 88% del atraso está en el área 3000."],
];
let ry = 4.85;
reads.forEach(([h,d])=>{
  dot(s, 0.9, ry+0.06, C.copper, 0.12);
  s.addText([
    { text:h+"  ", options:{ bold:true, color:C.white } },
    { text:d, options:{ color:C.ice } },
  ], { x:1.15, y:ry-0.05, w:11.4, h:0.55, fontFace:FT, fontSize:13, lineSpacing:17, margin:0, valign:"top" });
  ry += 0.6;
});
footer(s, 2, false);

// =====================================================================
// SLIDE 3 — Semáforo por área
// =====================================================================
s = p.addSlide(); bg(s, C.paper);
kicker(s, "Estado por área", 0.5, 0.45);
s.addText("Semáforo integrado de los tres frentes", { x:0.5, y:0.72, w:12, h:0.6, fontFace:FH, fontSize:30, bold:true, color:C.ink, margin:0 });

const areasS = [
  { a:"2000", n:"Chancado (Área Seca)", sub:35, c100:"100%", top:"17/35 · 48,6%", p1:"91,8%", ab:"38 (15)", st:"Al día", stc:C.good },
  { a:"3000", n:"Molienda–Flotación (Área Húmeda)", sub:76, c100:"46,1%", top:"8/76 · 10,5%", p1:"32,6%", ab:"707 (366)", st:"Crítico", stc:C.crit },
  { a:"4000", n:"Espesamiento y Relaves", sub:22, c100:"86,4%", top:"1/22 · 4,5%", p1:"75,9%", ab:"84 (19)", st:"Atención", stc:C.warn },
  { a:"8000", n:"Salas de Control", sub:2, c100:"50%", top:"0/2 · 0%", p1:"91,7%", ab:"1 (1)", st:"Atención", stc:C.warn },
];
// encabezado tabla
const cols = [
  { t:"ÁREA", x:0.55, w:3.2, al:"left" },
  { t:"SUBSIST.", x:3.8, w:1.0, al:"right" },
  { t:"CAMINATA 100%", x:4.85, w:1.7, al:"right" },
  { t:"TOP ENTREGADAS", x:6.6, w:1.9, al:"right" },
  { t:"P1 % CIERRE", x:8.55, w:1.4, al:"right" },
  { t:"P1 ABIERTOS (VENC.)", x:9.95, w:1.9, al:"right" },
  { t:"ESTADO", x:11.9, w:0.95, al:"center" },
];
let ty = 1.72;
cols.forEach(c=> s.addText(c.t, { x:c.x, y:ty, w:c.w, h:0.3, fontFace:FT, fontSize:9.5, bold:true, color:C.ink2, align:c.al, charSpacing:0.5, margin:0 }));
s.addShape(p.ShapeType.line, { x:0.55, y:2.06, w:12.3, h:0, line:{ color:C.steel, width:1 } });
let rowy = 2.2; const rh = 1.02;
areasS.forEach((r,i)=>{
  if(i%2===1) s.addShape(p.ShapeType.rect, { x:0.5, y:rowy-0.05, w:12.35, h:rh, fill:{ color:C.white }, line:{type:"none"} });
  s.addText(r.a, { x:cols[0].x, y:rowy, w:2.9, h:0.35, fontFace:FT, fontSize:15, bold:true, color:C.ink, align:"left", margin:0 });
  s.addText(r.n, { x:cols[0].x, y:rowy+0.36, w:3.1, h:0.3, fontFace:FT, fontSize:10, color:C.ink2, align:"left", margin:0 });
  const cellY = rowy+0.14;
  s.addText(String(r.sub), { x:cols[1].x, y:cellY, w:cols[1].w, h:0.35, fontFace:FT, fontSize:13, color:C.ink, align:"right", margin:0 });
  s.addText(r.c100, { x:cols[2].x, y:cellY, w:cols[2].w, h:0.35, fontFace:FT, fontSize:13, color:C.ink, align:"right", margin:0 });
  s.addText(r.top, { x:cols[3].x, y:cellY, w:cols[3].w, h:0.35, fontFace:FT, fontSize:13, color:C.ink, align:"right", margin:0 });
  s.addText(r.p1, { x:cols[4].x, y:cellY, w:cols[4].w, h:0.35, fontFace:FT, fontSize:13, color:C.ink, align:"right", margin:0 });
  s.addText(r.ab, { x:cols[5].x, y:cellY, w:cols[5].w, h:0.35, fontFace:FT, fontSize:13, color:C.ink, align:"right", margin:0 });
  // chip estado
  s.addShape(p.ShapeType.roundRect, { x:cols[6].x-0.02, y:rowy+0.12, w:1.0, h:0.42, rectRadius:0.21, fill:{ color:r.stc }, line:{type:"none"} });
  s.addText(r.st, { x:cols[6].x-0.02, y:rowy+0.13, w:1.0, h:0.4, fontFace:FT, fontSize:10.5, bold:true, color:C.white, align:"center", margin:0 });
  s.addShape(p.ShapeType.line, { x:0.55, y:rowy+rh-0.05, w:12.3, h:0, line:{ color:"D9DFE6", width:0.75 } });
  rowy += rh;
});
s.addText("Criterio · Crítico: caminata 100% < 60% o cierre P1 < 50%   ·   Atención: algún frente 60–90% o carpetas TOP sin entregar   ·   Al día: caminatas completas y P1 ≥ 90%",
  { x:0.5, y:6.5, w:12.3, h:0.3, fontFace:FT, fontSize:9.5, italic:true, color:C.ink2, margin:0 });
footer(s, 3, false);

// =====================================================================
// SLIDE 4 — Caminatas
// =====================================================================
s = p.addSlide(); bg(s, C.paper);
kicker(s, "Frente 1 · Caminatas", 0.5, 0.45);
s.addText("Avance de caminatas 80% y 100%", { x:0.5, y:0.72, w:12, h:0.6, fontFace:FH, fontSize:30, bold:true, color:C.ink, margin:0 });
s.addText("Las caminatas al 80% están cerradas en todas las áreas (135/135). El detalle está en el hito 100%, donde el área 3000 arrastra el atraso.",
  { x:0.5, y:1.35, w:12.3, h:0.5, fontFace:FT, fontSize:13.5, color:C.ink2, margin:0 });

// chart: caminata 100% realizada vs pendiente por área (stacked bar horizontal)
s.addChart(p.ChartType.bar, [
  { name:"Realizada", labels:["2000","3000","4000","8000"], values:[35,35,19,1] },
  { name:"Próxima a realizar", labels:["2000","3000","4000","8000"], values:[0,5,3,1] },
  { name:"Por programar", labels:["2000","3000","4000","8000"], values:[0,36,0,0] },
], {
  x:0.5, y:2.05, w:7.6, h:4.5, barDir:"bar", barGrouping:"stacked",
  chartColors:[C.goodD, C.blue, C.blueL],
  showTitle:true, title:"Subsistemas por estado de caminata 100%", titleFontFace:FT, titleFontSize:13, titleColor:C.ink,
  showValue:true, dataLabelPosition:"ctr", dataLabelFontFace:FT, dataLabelFontSize:11, dataLabelColor:C.white, dataLabelFontBold:true,
  showLegend:true, legendPos:"b", legendFontFace:FT, legendFontSize:11, legendColor:C.ink2,
  catAxisLabelFontFace:FT, catAxisLabelFontSize:13, catAxisLabelColor:C.ink, catAxisLabelFontBold:true,
  valAxisHidden:true, valGridLine:{ style:"none" }, catGridLine:{ style:"none" },
  barGapWidthPct:55, chartArea:{ fill:{ color:C.white } }, showValAxisTitle:false,
});

// panel derecho: cifras clave
s.addShape(p.ShapeType.roundRect, { x:8.35, y:2.05, w:4.48, h:4.5, rectRadius:0.06, fill:{ color:C.navy }, line:{type:"none"} });
s.addText("HITO 80%", { x:8.65, y:2.3, w:4, h:0.3, fontFace:FT, fontSize:11, bold:true, color:C.copperL, charSpacing:1.5, margin:0 });
s.addText("100%", { x:8.65, y:2.55, w:2, h:0.75, fontFace:FH, fontSize:40, bold:true, color:C.white, margin:0 });
s.addText("135 de 135 subsistemas", { x:8.65, y:3.32, w:4, h:0.3, fontFace:FT, fontSize:12, color:C.ice, margin:0 });
s.addShape(p.ShapeType.line, { x:8.65, y:3.85, w:3.9, h:0, line:{ color:C.navy2, width:1 } });
s.addText("HITO 100%", { x:8.65, y:4.02, w:4, h:0.3, fontFace:FT, fontSize:11, bold:true, color:C.copperL, charSpacing:1.5, margin:0 });
s.addText("66,7%", { x:8.65, y:4.27, w:2.4, h:0.75, fontFace:FH, fontSize:40, bold:true, color:C.white, margin:0 });
s.addText("90 realizadas · 45 pendientes", { x:8.65, y:5.04, w:4, h:0.3, fontFace:FT, fontSize:12, color:C.ice, margin:0 });
s.addText([
  { text:"Foco: ", options:{ bold:true, color:C.copperL } },
  { text:"9 subsistemas próximos a realizar y 36 por programar — 41 de ellos en el área 3000 (molienda–flotación).", options:{ color:C.ice } },
], { x:8.65, y:5.55, w:3.95, h:0.85, fontFace:FT, fontSize:11.5, lineSpacing:15, margin:0, valign:"top" });
footer(s, 4, false);

// =====================================================================
// SLIDE 5 — Carpetas TOP
// =====================================================================
s = p.addSlide(); bg(s, C.paper);
kicker(s, "Frente 2 · Carpetas TOP", 0.5, 0.45);
s.addText("Estatus de armado y entrega", { x:0.5, y:0.72, w:12, h:0.6, fontFace:FH, fontSize:30, bold:true, color:C.ink, margin:0 });
s.addText("26 carpetas entregadas de 135 (19,3%). Ninguna aprobada aún: el flujo se acumula en revisión y hay 8 rechazos que requieren rearmado.",
  { x:0.5, y:1.35, w:12.3, h:0.5, fontFace:FT, fontSize:13.5, color:C.ink2, margin:0 });

// chart TOP por área (stacked)
s.addChart(p.ChartType.bar, [
  { name:"Rechazada", labels:["2000","3000","4000","8000"], values:[2,6,0,0] },
  { name:"En revisión", labels:["2000","3000","4000","8000"], values:[15,1,1,0] },
  { name:"Observada", labels:["2000","3000","4000","8000"], values:[0,1,0,0] },
  { name:"Sin entregar", labels:["2000","3000","4000","8000"], values:[18,68,21,2] },
], {
  x:0.5, y:2.05, w:7.6, h:4.5, barDir:"bar", barGrouping:"stacked",
  chartColors:[C.crit, C.blue, C.warn, C.track],
  showTitle:true, title:"Carpetas TOP por área y estatus", titleFontFace:FT, titleFontSize:13, titleColor:C.ink,
  showValue:true, dataLabelPosition:"ctr", dataLabelFontFace:FT, dataLabelFontSize:10, dataLabelColor:C.white, dataLabelFontBold:true,
  showLegend:true, legendPos:"b", legendFontFace:FT, legendFontSize:11, legendColor:C.ink2,
  catAxisLabelFontFace:FT, catAxisLabelFontSize:13, catAxisLabelColor:C.ink, catAxisLabelFontBold:true,
  valAxisHidden:true, valGridLine:{ style:"none" }, catGridLine:{ style:"none" },
  barGapWidthPct:55, chartArea:{ fill:{ color:C.white } },
});

// funnel de estatus global (panel derecho)
s.addShape(p.ShapeType.roundRect, { x:8.35, y:2.05, w:4.48, h:4.5, rectRadius:0.06, fill:{ color:C.white }, line:{ color:"E4E9EF", width:1 } });
s.addText("EMBUDO DE ENTREGA (135)", { x:8.6, y:2.28, w:4, h:0.3, fontFace:FT, fontSize:11, bold:true, color:C.copper, charSpacing:1, margin:0 });
const funnel = [
  ["Entregadas", 26, C.blue],
  ["En revisión", 17, C.steel],
  ["Observadas", 1, C.warn],
  ["Rechazadas", 8, C.crit],
  ["Aprobadas", 0, C.good],
];
let fy = 2.68;
funnel.forEach(([l,v,c])=>{
  s.addShape(p.ShapeType.roundRect, { x:8.6, y:fy, w:3.98, h:0.62, rectRadius:0.05, fill:{ color:C.paper }, line:{type:"none"} });
  s.addShape(p.ShapeType.ellipse, { x:8.72, y:fy+0.19, w:0.24, h:0.24, fill:{ color:c }, line:{type:"none"} });
  s.addText(l, { x:9.08, y:fy, w:2.3, h:0.62, fontFace:FT, fontSize:13, color:C.ink, align:"left", valign:"middle", margin:0 });
  s.addText(String(v), { x:11.0, y:fy, w:1.4, h:0.62, fontFace:FH, fontSize:22, bold:true, color:c, align:"right", valign:"middle", margin:0 });
  fy += 0.72;
});
s.addText("Cuello de botella: revisión del cliente y rechazos por rearmar. La aprobación efectiva es el próximo hito a destrabar.",
  { x:8.6, y:6.32, w:4, h:0.5, fontFace:FT, fontSize:10.5, italic:true, color:C.ink2, margin:0 });
footer(s, 5, false);

// =====================================================================
// SLIDE 6 — Detalles P1
// =====================================================================
s = p.addSlide(); bg(s, C.paper);
kicker(s, "Frente 3 · Detalles P1", 0.5, 0.45);
s.addText("Detalles de terminación en condición P1", { x:0.5, y:0.72, w:12.3, h:0.6, fontFace:FH, fontSize:30, bold:true, color:C.ink, margin:0 });
s.addText("1.988 P1 levantados; 57,7% cerrados. El pendiente se concentra en Piping, Eléctrica e Instrumentación, y físicamente en el área 3000.",
  { x:0.5, y:1.35, w:12.3, h:0.5, fontFace:FT, fontSize:13.5, color:C.ink2, margin:0 });

// chart: P1 abiertos por disciplina (vencidos vs en plazo) stacked bar
s.addChart(p.ChartType.bar, [
  { name:"Vencidos al corte", labels:["Piping","Eléctrica","Instrum. y Control","Mecánica","Estructura","Obras Civiles"], values:[74,98,128,84,17,0] },
  { name:"En plazo", labels:["Piping","Eléctrica","Instrum. y Control","Mecánica","Estructura","Obras Civiles"], values:[177,118,53,42,19,20] },
], {
  x:0.5, y:2.05, w:7.7, h:4.5, barDir:"bar", barGrouping:"stacked",
  chartColors:[C.crit, C.blue],
  showTitle:true, title:"P1 abiertos por disciplina", titleFontFace:FT, titleFontSize:13, titleColor:C.ink,
  showValue:true, dataLabelPosition:"ctr", dataLabelFontFace:FT, dataLabelFontSize:10, dataLabelColor:C.white, dataLabelFontBold:true,
  showLegend:true, legendPos:"b", legendFontFace:FT, legendFontSize:11, legendColor:C.ink2,
  catAxisLabelFontFace:FT, catAxisLabelFontSize:11.5, catAxisLabelColor:C.ink, catAxisLabelFontBold:true,
  valAxisHidden:true, valGridLine:{ style:"none" }, catGridLine:{ style:"none" },
  barGapWidthPct:45, chartArea:{ fill:{ color:C.white } },
});

// panel: donut global + concentración por área
s.addShape(p.ShapeType.roundRect, { x:8.4, y:2.05, w:4.43, h:4.5, rectRadius:0.06, fill:{ color:C.navy }, line:{type:"none"} });
s.addText("ESTADO GLOBAL P1", { x:8.68, y:2.28, w:4, h:0.3, fontFace:FT, fontSize:11, bold:true, color:C.copperL, charSpacing:1.2, margin:0 });
s.addChart(p.ChartType.doughnut, [
  { name:"P1", labels:["Cerrados","Abiertos","Listo p/ revisión"], values:[1147,830,11] },
], {
  x:8.5, y:2.55, w:2.2, h:2.2, holeSize:62,
  chartColors:[C.good, C.crit, C.copperL],
  showTitle:false, showLegend:false, showValue:false,
  dataBorder:{ pct:1, color:C.navy }, chartArea:{ fill:{ color:C.navy } },
});
s.addText([{ text:"57,7%", options:{ fontSize:19, bold:true, color:C.white, breakLine:true } },
           { text:"cerrado", options:{ fontSize:10, color:C.ice } }],
  { x:8.5, y:3.28, w:2.2, h:0.75, align:"center", valign:"middle", fontFace:FH, margin:0 });
// leyenda donut
const dl = [["Cerrados","1.147",C.good],["Abiertos","830",C.crit],["Listo p/ revisión","11",C.copperL]];
let dly = 2.7;
dl.forEach(([l,v,c])=>{
  s.addShape(p.ShapeType.ellipse, { x:10.95, y:dly+0.04, w:0.2, h:0.2, fill:{color:c}, line:{type:"none"} });
  s.addText(l, { x:11.2, y:dly-0.05, w:1.3, h:0.35, fontFace:FT, fontSize:11, color:C.ice, align:"left", margin:0 });
  s.addText(v, { x:11.2, y:dly+0.2, w:1.3, h:0.3, fontFace:FT, fontSize:11, bold:true, color:C.white, align:"left", margin:0 });
  dly += 0.62;
});
s.addShape(p.ShapeType.line, { x:8.68, y:4.95, w:3.9, h:0, line:{ color:C.navy2, width:1 } });
s.addText([
  { text:"Concentración por área\n", options:{ bold:true, color:C.copperL, fontSize:11.5 } },
  { text:"707 de 830 P1 abiertos (85%) están en el área 3000. Áreas 2000 (91,8%) y 8000 (91,7%) ya casi cerradas.", options:{ color:C.ice, fontSize:11.5 } },
], { x:8.68, y:5.1, w:3.95, h:1.3, fontFace:FT, lineSpacing:15, margin:0, valign:"top" });
footer(s, 6, true);

// =====================================================================
// SLIDE 7 — Foco y próximos pasos
// =====================================================================
s = p.addSlide(); bg(s, C.navy);
s.addShape(p.ShapeType.ellipse, { x:10.8, y:-1.8, w:4.6, h:4.6, fill:{ color:C.navy2 }, line:{type:"none"} });
kicker(s, "Foco de gestión", 0.5, 0.5, C.copperL);
s.addText("Dónde poner el esfuerzo al 31-jul", { x:0.5, y:0.78, w:12, h:0.6, fontFace:FH, fontSize:30, bold:true, color:C.white, margin:0 });

const focus = [
  { n:"01", t:"Destrabar carpetas TOP", c:C.copper,
    d:"0 aprobadas de 135. Priorizar cierre de las 17 en revisión y rearmado de las 8 rechazadas (áreas 2000 y 3000) para convertir entregas en aprobaciones." },
  { n:"02", t:"Atacar el atraso del área 3000", c:C.crit,
    d:"707 P1 abiertos (85% del total) y 41 caminatas 100% pendientes. Es la ruta crítica: sin 3000 no mejora ningún indicador global." },
  { n:"03", t:"Recuperar los 401 P1 vencidos", c:C.warn,
    d:"Casi la mitad de los P1 abiertos superó su fecha de vencimiento. Foco en Instrumentación (128), Eléctrica (98) y Mecánica (84)." },
  { n:"04", t:"Cerrar caminatas 100% restantes", c:C.good,
    d:"45 subsistemas pendientes (9 próximos, 36 por programar). Programarlas habilita el levantamiento final de detalles y la entrega de carpetas." },
];
let fx = 0.5, fw = 6.0, fgap = 0.33;
let fry = 1.75;
focus.forEach((f,i)=>{
  const px = 0.5 + (i%2)*(fw+fgap);
  const py = 1.75 + Math.floor(i/2)*2.35;
  s.addShape(p.ShapeType.roundRect, { x:px, y:py, w:fw, h:2.1, rectRadius:0.07, fill:{ color:C.navy2 }, line:{type:"none"} });
  s.addText(f.n, { x:px+0.3, y:py+0.22, w:1.1, h:0.7, fontFace:FH, fontSize:34, bold:true, color:f.c, margin:0, align:"left" });
  s.addText(f.t, { x:px+1.35, y:py+0.28, w:fw-1.6, h:0.55, fontFace:FT, fontSize:17, bold:true, color:C.white, margin:0, align:"left", valign:"middle" });
  s.addText(f.d, { x:px+1.35, y:py+0.85, w:fw-1.65, h:1.1, fontFace:FT, fontSize:12, color:C.ice, margin:0, align:"left", lineSpacing:16, valign:"top" });
});
footer(s, 7, true);

p.writeFile({ fileName: "/home/user/Inversiones/panel_control_TOP_P1/Panel_Control_TOP_P1.pptx" })
 .then(f => console.log("OK:", f))
 .catch(e => { console.error(e); process.exit(1); });
