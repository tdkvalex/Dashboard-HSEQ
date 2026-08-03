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
  // Gris para el segmento "sin entregar / pendiente / sin programar" DENTRO de
  // los gráficos: tiene que ser lo bastante oscuro para que el número blanco
  // impreso encima se lea (el track claro dejaba la cifra invisible).
  neutral: "5B6B7C",
  // Variantes oscurecidas de blueL y warn: las originales eran demasiado claras
  // y el número blanco impreso encima quedaba ilegible.
  blueM: "3A6FA8", warnD: "9A6B0F",
};

const FT = "Calibri", FH = "Cambria";
// Los valores 0 no se dibujan: si se dejan, PowerPoint imprime un «0» pegado
// al eje que se pisa con la etiqueta vecina y ensucia la lectura.
const z = (v) => (v > 0 ? v : null);

// Etiquetas comunes de los gráficos de barras: número blanco, en negrita y de
// buen tamaño, sobre segmentos que siempre tienen contraste suficiente.
const ETIQ = { showValue: true, dataLabelPosition: "ctr", dataLabelFontFace: FT,
  dataLabelFontSize: 11, dataLabelColor: "FFFFFF", dataLabelFontBold: true };


// Nombre del contrato de Arqueros para la portada del proyecto. Se toma del JSON
// si algún día viene en los datos; si no, se usa este.
// El título de la portada ya dice «Arqueros», así que acá va solo el contrato y
// el cliente, con el mismo formato «descripción · cliente» de los otros dos.
const PROY_MASA = (D.meta && D.meta.proyecto) ||
  "Contrato Electromecánico Planta Concentradora y de Espesado · MASA";
const CAP = `Fuente: Estatus_Resumen_General_QAQC.xlsx · Corte ${D.meta.corte_texto} · Áreas ${D.areas.join("/")}`;

const p = new pptxgen();
p.defineLayout({ name: "W", width: 13.333, height: 7.5 });
p.layout = "W";
p.author = "Control de Proyecto";
p.title = "Panel de Control — Carpetas TOP · Caminatas · Detalles P1";

const bg = (s, color) => { s.background = { color }; };

// ---------- logo corporativo ----------
// Se lee de besalco_logo.png (misma carpeta). Si no está, las láminas salen
// igual, solo sin logo: nunca rompe la generación.
const LOGO = path.join(AQUI, "besalco_logo.png");
const HAY_LOGO = fs.existsSync(LOGO);
if (!HAY_LOGO) console.log("Aviso: no se encontró besalco_logo.png — la PPT sale sin logo.");
// En portadas (fondo azul) el logo va más grande; en láminas de contenido, menor.
function logo(s, portada) {
  if (!HAY_LOGO) return;
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

// ---------- portada de proyecto ----------
// Misma estructura para los tres proyectos, para que el mazo se lea parejo.
function portada(s, num, nombre, descripcion, detalle, kpis, pie, cap) {
  bg(s, C.navy);
  s.addShape(p.ShapeType.ellipse, { x: 11.0, y: -1.6, w: 4.2, h: 4.2, fill: { color: C.navy2 }, line: { type: "none" } });
  s.addShape(p.ShapeType.ellipse, { x: 12.1, y: 5.0, w: 3.0, h: 3.0, fill: { color: C.navy2 }, line: { type: "none" } });
  logo(s, true);
  dot(s, 0.55, 1.35, C.copper, 0.22);
  s.addText(`PROYECTO ${num}`, { x: 0.85, y: 1.2, w: 10, h: 0.4, fontFace: FT, fontSize: 13,
    bold: true, color: C.copperL, charSpacing: 3, margin: 0 });
  s.addText(nombre, { x: 0.82, y: 1.72, w: 11.6, h: 1.1, fontFace: FH, fontSize: 46,
    bold: true, color: C.white, margin: 0 });
  s.addText(descripcion, { x: 0.85, y: 2.9, w: 10.5, h: 0.5, fontFace: FT, fontSize: 16,
    color: C.ice, margin: 0 });
  s.addText(detalle, { x: 0.85, y: 3.5, w: 9.6, h: 1.3, fontFace: FT, fontSize: 13.5,
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
// PROYECTO DESALADORA — se agrega solo si existe datos_desaladora.json
// (lo genera desaladora.py). Las cifras salen íntegras de ese archivo.
// =====================================================================
// `s` es la lámina en curso: se declara acá porque los tres proyectos la
// reasignan, y el orden del mazo lo fija el orden de estos bloques.
let s;

const rutaDes = path.join(AQUI, "datos_desaladora.json");
if (fs.existsSync(rutaDes)) {
  const X = JSON.parse(fs.readFileSync(rutaDes, "utf8"));
  const XS = X.subsistemas, XC = X.caminatasPorTipo, XP = X.carpetas, XU = X.punch;
  const xp1 = XU.porCategoria.P1 || { total: 0, cerrados: 0, abiertos: 0, atrasados: 0 };
  const xg = XU.global;
  const CAPX = `Fuente: REPORTE_GERENCIAL + Listado de Puntos Punch Consolidado · ` +
    `Corte ${X.meta.corteTexto} · Atrasos al ${X.meta.hoy}`;
  const opV = XC.Operable ? XC.Operable[XC.Operable.vigente] : null;
  const faV = XC.Facility ? XC.Facility[XC.Facility.vigente] : null;
  const zonas = Object.keys(XS.porZona);
  const discOrden = Object.entries(XU.p1PorDisciplina).sort((a, b) => b[1].abiertos - a[1].abiertos);
  const discAbiertas = discOrden.filter(([, v]) => v.abiertos > 0);

  // ---------------- Portada de proyecto ----------------
  s = p.addSlide(); bg(s, C.navy); logo(s, true);
  s.addShape(p.ShapeType.ellipse, { x: 11.0, y: -1.6, w: 4.2, h: 4.2, fill: { color: C.navy2 }, line: { type: "none" } });
  s.addShape(p.ShapeType.ellipse, { x: 12.1, y: 5.0, w: 3.0, h: 3.0, fill: { color: C.navy2 }, line: { type: "none" } });
  dot(s, 0.55, 1.35, C.copper, 0.22);
  s.addText("PROYECTO 01 · P2416", { x: 0.85, y: 1.2, w: 10, h: 0.4, fontFace: FT, fontSize: 13,
    bold: true, color: C.copperL, charSpacing: 3, margin: 0 });
  s.addText("Desaladora", { x: 0.82, y: 1.72, w: 11.6, h: 1.1, fontFace: FH, fontSize: 46,
    bold: true, color: C.white, margin: 0 });
  s.addText(`${X.meta.descripcion} · ${X.meta.cliente}`, { x: 0.85, y: 2.9, w: 10.5, h: 0.5,
    fontFace: FT, fontSize: 16, color: C.ice, margin: 0 });
  // Los componentes no suman al universo (son entregas parciales de un
  // subsistema ya contado): se dice en la portada para que la cifra no baile
  // contra el reporte gerencial, que trae dos filas más.
  const XCO = X.componentes || { total: 0 };
  s.addText(`Contrato ${X.meta.contrato} · ${nf(XS.total)} subsistemas ` +
    `(${Object.entries(XS.porTipo).map(([k, v]) => `${nf(v)} ${k}`).join(" · ")}) en ${zonas.join(" y ")}.` +
    (XCO.total ? `\nNo suman al universo ${nf(XCO.total)} componentes ya entregados: ` +
      `son entregas parciales de un subsistema que ya está contado.` : ""),
    { x: 0.85, y: 3.45, w: 9.6, h: 1.1, fontFace: FT, fontSize: 13.5, color: C.steel, lineSpacing: 20, margin: 0 });

  const kpx = [
    [opV ? pct(opV.realizada, opV.aplica) : "—", `Caminata ${XC.Operable ? XC.Operable.vigente : 2} · Operables`],
    [faV ? pct(faV.realizada, faV.aplica) : "—", `Caminata ${XC.Facility ? XC.Facility.vigente : 1} · Facility`],
    [pct(XP.entregadas, XP.total), "Carpetas entregadas"],
    [pct(xp1.cerrados, xp1.total), "Cierre de punch P1"],
  ];
  let kxx = 0.85;
  kpx.forEach(([v, l]) => {
    s.addText(v, { x: kxx, y: 5.3, w: 2.9, h: 0.7, fontFace: FH, fontSize: 34, bold: true, color: C.copperL, margin: 0 });
    s.addText(l, { x: kxx, y: 6.0, w: 2.9, h: 0.55, fontFace: FT, fontSize: 11.5, color: C.ice, margin: 0 });
    kxx += 3.0;
  });
  s.addText(`Corte: ${X.meta.corteTexto}`, { x: 0.85, y: 6.75, w: 8, h: 0.3, fontFace: FT,
    fontSize: 11, italic: true, color: C.steel, margin: 0 });
  footer(s, true, CAPX);

  // ---------------- Resumen ejecutivo ----------------
  s = p.addSlide(); bg(s, C.paper); logo(s, false);
  kicker(s, "Desaladora · Resumen ejecutivo", 0.5, 0.45);
  s.addText("Estado global de los tres frentes", { x: 0.5, y: 0.72, w: 12, h: 0.6, fontFace: FH, fontSize: 30, bold: true, color: C.ink, margin: 0 });

  const cardsX = [
    { p: opV ? pct(opV.realizada, opV.aplica) : "—", v: opV ? `${nf(opV.realizada)}/${nf(opV.aplica)}` : "—",
      t: `Caminata ${XC.Operable ? XC.Operable.vigente : 2} · Operables`,
      d: opV ? `${nf(opV.pendiente)} subsistemas por caminar` : "", c: C.blue },
    { p: faV ? pct(faV.realizada, faV.aplica) : "—", v: faV ? `${nf(faV.realizada)}/${nf(faV.aplica)}` : "—",
      t: `Caminata ${XC.Facility ? XC.Facility.vigente : 1} · Facility`,
      d: faV ? `${nf(faV.pendiente)} facilities por caminar` : "", c: C.blue },
    { p: pct(XP.entregadas, XP.total), v: `${nf(XP.entregadas)}/${nf(XP.total)}`, t: "Carpetas entregadas",
      d: `${nf(XP.aprobadas)} aprobadas · ${nf(XP.rechazadas)} rechazadas · ${nf(XP.enPreparacion)} en preparación`, c: C.copper },
    { p: pct(xp1.cerrados, xp1.total), v: `${nf(xp1.cerrados)}/${nf(xp1.total)}`, t: "Cierre de punch P1",
      d: `${nf(xp1.abiertos)} abiertos · ${nf(xp1.atrasados)} atrasados`, c: C.crit },
  ];
  let cxx = 0.5;
  cardsX.forEach((cd) => {
    s.addShape(p.ShapeType.roundRect, { x: cxx, y: 1.6, w: 3.0, h: 2.5, rectRadius: 0.08,
      fill: { color: C.white }, line: { color: "E4E9EF", width: 1 },
      shadow: { type: "outer", color: "9AA7B5", blur: 7, offset: 2, angle: 90, opacity: 0.28 } });
    dot(s, cxx + 0.28, 1.88, cd.c, 0.16);
    s.addText(cd.p, { x: cxx + 0.28, y: 2.12, w: 2.5, h: 0.82, fontFace: FH, fontSize: 38, bold: true, color: cd.c, margin: 0 });
    s.addText(cd.v, { x: cxx + 0.3, y: 2.95, w: 2.5, h: 0.32, fontFace: FT, fontSize: 14, bold: true, color: C.ink, margin: 0 });
    s.addText(cd.t, { x: cxx + 0.3, y: 3.27, w: 2.45, h: 0.32, fontFace: FT, fontSize: 12, color: C.ink2, margin: 0 });
    s.addText(cd.d, { x: cxx + 0.3, y: 3.58, w: 2.45, h: 0.46, fontFace: FT, fontSize: 9.5, color: C.ink2, lineSpacing: 12, margin: 0 });
    cxx += 3.13;
  });

  s.addShape(p.ShapeType.roundRect, { x: 0.5, y: 4.32, w: 12.33, h: 2.42, rectRadius: 0.06, fill: { color: C.navy }, line: { type: "none" } });
  s.addText("LECTURA DEL PERÍODO", { x: 0.85, y: 4.5, w: 6, h: 0.3, fontFace: FT, fontSize: 11, bold: true, color: C.copperL, charSpacing: 2, margin: 0 });
  const topDisc = discAbiertas[0];
  const readsX = [
    ["El punch cierra a buen ritmo, pero la deuda vencida es la mitad de lo abierto.",
      `${nf(xg.cerrados)} de ${nf(xg.total)} ítems cerrados (${pct(xg.cerrados, xg.total)}); de los ${nf(xg.abiertos)} abiertos, ` +
      `${nf(xg.atrasados)} ya pasaron su fecha requerida de cierre (${pct(xg.atrasados, xg.abiertos)}).`],
    [`Las caminatas avanzan a distinta velocidad según el tipo de subsistema.`,
      (opV ? `Operables ${nf(opV.realizada)}/${nf(opV.aplica)} por Caminata ${XC.Operable.vigente} (${pct(opV.realizada, opV.aplica)})` : "") +
      (faV ? `; Facility ${nf(faV.realizada)}/${nf(faV.aplica)} por Caminata ${XC.Facility.vigente} (${pct(faV.realizada, faV.aplica)}).` : ".")],
    [`Las carpetas están a ${pct(XP.entregadas, XP.total)} de entrega y el cuello está en el armado interno.`,
      `${nf(XP.entregadas)} entregadas de ${nf(XP.total)}: ${nf(XP.aprobadas)} aprobadas, ${nf(XP.enRevision)} en revisión y ` +
      `${nf(XP.rechazadas)} rechazadas. Quedan ${nf(XP.enPreparacion)} en preparación y ${nf(XP.sinEntregar)} sin iniciar.`],
  ];
  if (topDisc) readsX.push([`${topDisc[0]} concentra el pendiente de P1.`,
    `${nf(topDisc[1].abiertos)} de los ${nf(xp1.abiertos)} P1 abiertos (${pct(topDisc[1].abiertos, xp1.abiertos)}), ` +
    `con ${nf(topDisc[1].atrasados)} atrasados.`]);
  let ryx = 4.86;
  readsX.forEach(([h, d]) => {
    dot(s, 0.9, ryx + 0.06, C.copper, 0.12);
    s.addText([{ text: h + "  ", options: { bold: true, color: C.white } },
               { text: d, options: { color: C.ice } }],
      { x: 1.15, y: ryx - 0.05, w: 11.4, h: 0.5, fontFace: FT, fontSize: 11.5, lineSpacing: 15, margin: 0, valign: "top" });
    ryx += 0.46;
  });
  footer(s, false, CAPX, "Desaladora");

  // ---------------- Caminatas y carpetas ----------------
  s = p.addSlide(); bg(s, C.paper); logo(s, false);
  kicker(s, "Desaladora · Frentes 1 y 2", 0.5, 0.45);
  s.addText("Caminatas y certificados de entrega", { x: 0.5, y: 0.72, w: 12.3, h: 0.6, fontFace: FH, fontSize: 30, bold: true, color: C.ink, margin: 0 });
  s.addText(`Cada tipo de subsistema se mide con su propia caminata: los Operables por la Caminata ` +
    `${XC.Operable ? XC.Operable.vigente : 2} y los Facility por la Caminata ${XC.Facility ? XC.Facility.vigente : 1}, ` +
    `tal como lo declara el propio reporte gerencial.`,
    { x: 0.5, y: 1.35, w: 12.3, h: 0.5, fontFace: FT, fontSize: 13.5, color: C.ink2, margin: 0 });

  const tiposCam = Object.entries(XC).filter(([, v]) => v[v.vigente] && v[v.vigente].aplica);
  const lblCam = tiposCam.map(([t, v]) => `${t} (C${v.vigente})`);
  s.addChart(p.ChartType.bar, [
    { name: "Realizada", labels: lblCam, values: tiposCam.map(([, v]) => z(v[v.vigente].realizada)) },
    { name: "Pendiente", labels: lblCam, values: tiposCam.map(([, v]) => z(v[v.vigente].pendiente)) },
  ], {
    x: 0.5, y: 2.05, w: 6.1, h: 4.5, barDir: "bar", barGrouping: "stacked",
    chartColors: [C.goodD, C.neutral],
    showTitle: true, title: "Subsistemas por estado de su caminata", titleFontFace: FT, titleFontSize: 13, titleColor: C.ink,
    ...ETIQ,
    showLegend: true, legendPos: "b", legendFontFace: FT, legendFontSize: 11, legendColor: C.ink2,
    catAxisLabelFontFace: FT, catAxisLabelFontSize: 12, catAxisLabelColor: C.ink, catAxisLabelFontBold: true,
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    barGapWidthPct: 55, chartArea: { fill: { color: C.white } },
  });

  s.addShape(p.ShapeType.roundRect, { x: 6.85, y: 2.05, w: 5.98, h: 4.5, rectRadius: 0.06, fill: { color: C.white }, line: { color: "E4E9EF", width: 1 } });
  s.addText(`CERTIFICADO DE ENTREGA (${nf(XP.total)} SUBSISTEMAS)`, { x: 7.1, y: 2.28, w: 5.5, h: 0.3,
    fontFace: FT, fontSize: 11, bold: true, color: C.copper, charSpacing: 1, margin: 0 });
  const escalera = [
    ["Aprobada", XP.aprobadas, C.good],
    ["En revisión del cliente", XP.enRevision, C.blue],
    ["Rechazada", XP.rechazadas, C.crit],
    ["En preparación (interna)", XP.enPreparacion, C.warn],
    ["Sin entregar", XP.sinEntregar, C.steel],
  ];
  let fyx = 2.7;
  escalera.forEach(([l, v, c]) => {
    s.addShape(p.ShapeType.roundRect, { x: 7.1, y: fyx, w: 5.48, h: 0.62, rectRadius: 0.05, fill: { color: C.paper }, line: { type: "none" } });
    s.addShape(p.ShapeType.ellipse, { x: 7.22, y: fyx + 0.19, w: 0.24, h: 0.24, fill: { color: c }, line: { type: "none" } });
    s.addText(l, { x: 7.58, y: fyx, w: 2.9, h: 0.62, fontFace: FT, fontSize: 12.5, color: C.ink, valign: "middle", margin: 0 });
    s.addText(pct(v, XP.total), { x: 10.4, y: fyx, w: 1.1, h: 0.62, fontFace: FT, fontSize: 11, color: C.ink2, align: "right", valign: "middle", margin: 0 });
    s.addText(nf(v), { x: 11.55, y: fyx, w: 0.9, h: 0.62, fontFace: FH, fontSize: 22, bold: true, color: c, align: "right", valign: "middle", margin: 0 });
    fyx += 0.68;
  });
  s.addText(`${nf(XP.entregadas)} carpetas ya están en manos del cliente (${pct(XP.entregadas, XP.total)}) y ` +
    `${nf(XP.tarjetaVerde)} subsistemas tienen tarjeta verde. El grueso del pendiente todavía no sale de Besalco.`,
    { x: 7.1, y: 6.12, w: 5.5, h: 0.4, fontFace: FT, fontSize: 10, italic: true, color: C.ink2, lineSpacing: 12, margin: 0 });
  footer(s, false, CAPX, "Desaladora");

  // ---------------- Punch P1 ----------------
  s = p.addSlide(); bg(s, C.paper); logo(s, false);
  kicker(s, "Desaladora · Frente 3", 0.5, 0.45);
  s.addText("Punch list en condición P1", { x: 0.5, y: 0.72, w: 12.3, h: 0.6, fontFace: FH, fontSize: 30, bold: true, color: C.ink, margin: 0 });
  s.addText(`${nf(xp1.total)} P1 levantados; ${pct(xp1.cerrados, xp1.total)} cerrados. ` +
    `Atrasado = ítem abierto cuya fecha requerida de cierre ya venció al ${X.meta.hoy}.`,
    { x: 0.5, y: 1.35, w: 12.3, h: 0.5, fontFace: FT, fontSize: 13.5, color: C.ink2, margin: 0 });

  const lblD = discAbiertas.map(([d]) => (X.discCorta && X.discCorta[d]) || d);
  s.addChart(p.ChartType.bar, [
    { name: "Atrasados", labels: lblD, values: discAbiertas.map(([, v]) => z(v.atrasados)) },
    { name: "En plazo o sin fecha", labels: lblD, values: discAbiertas.map(([, v]) => z(v.abiertos - v.atrasados)) },
  ], {
    x: 0.5, y: 2.05, w: 7.7, h: 4.5, barDir: "bar", barGrouping: "stacked",
    chartColors: [C.crit, C.blue],
    showTitle: true, title: "P1 abiertos por disciplina", titleFontFace: FT, titleFontSize: 13, titleColor: C.ink,
    ...ETIQ,
    showLegend: true, legendPos: "b", legendFontFace: FT, legendFontSize: 11, legendColor: C.ink2,
    catAxisLabelFontFace: FT, catAxisLabelFontSize: 11.5, catAxisLabelColor: C.ink, catAxisLabelFontBold: true,
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    barGapWidthPct: 45, chartArea: { fill: { color: C.white } },
  });

  s.addShape(p.ShapeType.roundRect, { x: 8.4, y: 2.05, w: 4.43, h: 4.5, rectRadius: 0.06, fill: { color: C.navy }, line: { type: "none" } });
  s.addText("ESTADO GLOBAL P1", { x: 8.68, y: 2.28, w: 4, h: 0.3, fontFace: FT, fontSize: 11, bold: true, color: C.copperL, charSpacing: 1.2, margin: 0 });
  s.addChart(p.ChartType.doughnut, [{ name: "P1", labels: ["Cerrados", "Atrasados", "Abiertos en plazo"],
    values: [xp1.cerrados, xp1.atrasados, xp1.abiertos - xp1.atrasados] }], {
    x: 8.5, y: 2.55, w: 2.2, h: 2.2, holeSize: 62,
    chartColors: [C.good, C.crit, C.copperL],
    showTitle: false, showLegend: false, showValue: false,
    dataBorder: { pct: 1, color: C.navy }, chartArea: { fill: { color: C.navy } },
  });
  s.addText([{ text: pct(xp1.cerrados, xp1.total), options: { fontSize: 19, bold: true, color: C.white, breakLine: true } },
             { text: "cerrado", options: { fontSize: 10, color: C.ice } }],
    { x: 8.5, y: 3.28, w: 2.2, h: 0.75, align: "center", valign: "middle", fontFace: FH, margin: 0 });
  const dlx = [["Cerrados", xp1.cerrados, C.good], ["Atrasados", xp1.atrasados, C.crit],
               ["Abiertos en plazo", xp1.abiertos - xp1.atrasados, C.copperL]];
  let dlyx = 2.7;
  dlx.forEach(([l, v, c]) => {
    s.addShape(p.ShapeType.ellipse, { x: 10.95, y: dlyx + 0.04, w: 0.2, h: 0.2, fill: { color: c }, line: { type: "none" } });
    s.addText(l, { x: 11.2, y: dlyx - 0.05, w: 1.55, h: 0.35, fontFace: FT, fontSize: 10.5, color: C.ice, margin: 0 });
    s.addText(nf(v), { x: 11.2, y: dlyx + 0.22, w: 1.5, h: 0.3, fontFace: FT, fontSize: 11, bold: true, color: C.white, margin: 0 });
    dlyx += 0.62;
  });
  s.addShape(p.ShapeType.line, { x: 8.68, y: 4.95, w: 3.9, h: 0, line: { color: C.navy2, width: 1 } });
  const ant = XU.antiguedad;
  s.addText([{ text: "Antigüedad del atraso\n", options: { bold: true, color: C.copperL, fontSize: 11.5 } },
             { text: ant.max
                 ? `Mediana de ${nf(ant.mediana)} días vencidos; el más antiguo lleva ${nf(ant.max)}. ` +
                   Object.entries(ant.tramos).filter(([, v]) => v).map(([k, v]) => `${nf(v)} entre ${k}`).join(", ") + "."
                 : "Sin ítems atrasados al corte.",
               options: { color: C.ice, fontSize: 11.5 } }],
    { x: 8.68, y: 5.1, w: 3.95, h: 1.3, fontFace: FT, lineSpacing: 15, margin: 0, valign: "top" });
  footer(s, true, CAPX, "Desaladora");

  // ---------------- Foco de gestión ----------------
  s = p.addSlide(); bg(s, C.navy); logo(s, true);
  s.addShape(p.ShapeType.ellipse, { x: 10.8, y: -1.8, w: 4.6, h: 4.6, fill: { color: C.navy2 }, line: { type: "none" } });
  kicker(s, "Desaladora · Foco de gestión", 0.5, 0.5, C.copperL);
  s.addText("Dónde poner el esfuerzo esta semana", { x: 0.5, y: 0.78, w: 12, h: 0.6, fontFace: FH, fontSize: 30, bold: true, color: C.white, margin: 0 });

  const zonaFoco = zonas
    .map((z) => [z, (X.semaforo[z] && X.semaforo[z][2]) || { abiertos: 0, atrasados: 0 }])
    .sort((a, b) => b[1].atrasados - a[1].atrasados)[0];
  const focusX = [
    { n: "01", t: `Recuperar los ${nf(xg.atrasados)} ítems atrasados`, c: C.crit,
      d: `${pct(xg.atrasados, xg.abiertos)} de los ${nf(xg.abiertos)} abiertos ya venció su fecha requerida. ` +
         (discAbiertas.length ? `Foco en ${discAbiertas.slice(0, 3).map(([d, v]) => `${d} (${nf(v.atrasados)})`).join(", ")}.` : "") },
    { n: "02", t: opV ? `Ejecutar las ${nf(opV.pendiente)} caminatas de Operables pendientes` : "Ejecutar las caminatas pendientes", c: C.copper,
      d: opV ? `La Caminata ${XC.Operable.vigente} va en ${pct(opV.realizada, opV.aplica)} (${nf(opV.realizada)}/${nf(opV.aplica)}). ` +
         `Sin caminata no se levanta el punch final ni se habilita la entrega de la carpeta.` : "" },
    { n: "03", t: `Sacar del armado interno las ${nf(XP.enPreparacion + XP.sinEntregar)} carpetas retenidas`, c: C.warn,
      d: `${nf(XP.enPreparacion)} están en preparación y ${nf(XP.sinEntregar)} sin iniciar. ` +
         `Solo ${nf(XP.entregadas)} de ${nf(XP.total)} (${pct(XP.entregadas, XP.total)}) llegaron al cliente; ` +
         `${nf(XP.rechazadas)} volvieron rechazadas y hay que rearmarlas.` },
    { n: "04", t: zonaFoco ? `Atender la zona ${zonaFoco[0]}` : "Atender la zona crítica", c: C.good,
      d: zonaFoco ? `Concentra ${nf(zonaFoco[1].abiertos)} P1 abiertos y ${nf(zonaFoco[1].atrasados)} atrasados ` +
         `sobre ${nf(XS.porZona[zonaFoco[0]])} subsistemas. Es la ruta crítica del proyecto.` : "" },
  ];
  focusX.forEach((f, i) => {
    const px = 0.5 + (i % 2) * (6.0 + 0.33);
    const py = 1.75 + Math.floor(i / 2) * 2.35;
    s.addShape(p.ShapeType.roundRect, { x: px, y: py, w: 6.0, h: 2.1, rectRadius: 0.07, fill: { color: C.navy2 }, line: { type: "none" } });
    s.addText(f.n, { x: px + 0.3, y: py + 0.22, w: 1.1, h: 0.7, fontFace: FH, fontSize: 34, bold: true, color: f.c, margin: 0 });
    s.addText(f.t, { x: px + 1.35, y: py + 0.2, w: 6.0 - 1.6, h: 0.66, fontFace: FT, fontSize: 14, bold: true, color: C.white, margin: 0, valign: "middle" });
    s.addText(f.d, { x: px + 1.35, y: py + 0.88, w: 6.0 - 1.65, h: 1.1, fontFace: FT, fontSize: 11, color: C.ice, margin: 0, lineSpacing: 14, valign: "top" });
  });
  footer(s, true, CAPX, "Desaladora");

  console.log(`Desaladora: 5 láminas agregadas (corte ${X.meta.corteTexto}).`);
} else {
  console.log("No se encontró datos_desaladora.json — la PPT sale sin Desaladora.");
}

// =====================================================================
// PROYECTO TALABRE — se agrega solo si existe datos_talabre.json
// (lo genera talabre.py). Las cifras salen íntegras de ese archivo.
// =====================================================================
const rutaTal = path.join(AQUI, "datos_talabre.json");
if (fs.existsSync(rutaTal)) {
  const Y = JSON.parse(fs.readFileSync(rutaTal, "utf8"));
  const YS = Y.subsistemas, YC = Y.caminatas, YP = Y.carpetas, YT = Y.dt;
  const yp1 = YT.porPrioridad.P1 || { total: 0, cerrados: 0, abiertos: 0, atrasados: 0 };
  const yg = YT.global;
  // P1+P2 es lo que se gestiona: los P3/P4 no se abordan mientras el cliente no
  // formalice su ejecución. El detalle por prioridad va completo más abajo.
  const yg12 = ["P1", "P2"].reduce((o, k) => { const v = YT.porPrioridad[k]; if (v) {
    o.total += v.total; o.cerrados += v.cerrados; o.abiertos += v.abiertos; o.atrasados += v.atrasados; }
    return o; }, { total: 0, cerrados: 0, abiertos: 0, atrasados: 0 });
  const CAPY = `Fuente: STATUS_SUBSISTEMAS_TALABRE + Detalle_de_Terminaciones · ` +
    `${nf(YS.total)} subsistemas · ${nf(yg.total)} DT · Atrasos al ${Y.meta.hoy}`;
  const c1 = YC["1"], c2 = YC["2"];
  const areasY = Y.areas;
  const discOrden = Object.entries(YT.porDisciplina).sort((a, b) => b[1].abiertos - a[1].abiertos);
  const areaFocoY = Object.entries(YT.porArea).sort((a, b) => b[1].abiertos - a[1].abiertos)[0];

  // ---------------- Portada de proyecto ----------------
  s = p.addSlide(); bg(s, C.navy); logo(s, true);
  s.addShape(p.ShapeType.ellipse, { x: 11.0, y: -1.6, w: 4.2, h: 4.2, fill: { color: C.navy2 }, line: { type: "none" } });
  s.addShape(p.ShapeType.ellipse, { x: 12.1, y: 5.0, w: 3.0, h: 3.0, fill: { color: C.navy2 }, line: { type: "none" } });
  dot(s, 0.55, 1.35, C.copper, 0.22);
  s.addText("PROYECTO 02 · P2407", { x: 0.85, y: 1.2, w: 10, h: 0.4, fontFace: FT, fontSize: 13,
    bold: true, color: C.copperL, charSpacing: 3, margin: 0 });
  s.addText("Talabre", { x: 0.82, y: 1.72, w: 11.6, h: 1.1, fontFace: FH, fontSize: 46,
    bold: true, color: C.white, margin: 0 });
  s.addText(`${Y.meta.descripcion} · ${Y.meta.cliente}`, { x: 0.85, y: 2.9, w: 10.5, h: 0.5,
    fontFace: FT, fontSize: 16, color: C.ice, margin: 0 });
  s.addText(`Contrato ${Y.meta.contrato} · ${nf(YS.total)} subsistemas en ${areasY.length} áreas · ` +
    `${nf(yg.total)} detalles de terminación levantados.`,
    { x: 0.85, y: 3.45, w: 9.6, h: 0.9, fontFace: FT, fontSize: 13.5, color: C.steel, lineSpacing: 20, margin: 0 });

  const kpy = [
    [pct(c1.realizada, c1.exigible), "Caminata 1 · exigible"],
    [pct(c2.realizada, c2.exigible), "Caminata 2 · exigible"],
    [`${YP.promedio.toLocaleString("es-CL")}%`, "Avance PEC promedio"],
    [pct(yg12.cerrados, yg12.total), "Cierre de DT P1+P2"],
  ];
  let kxy = 0.85;
  kpy.forEach(([v, l]) => {
    s.addText(v, { x: kxy, y: 5.3, w: 2.9, h: 0.7, fontFace: FH, fontSize: 34, bold: true, color: C.copperL, margin: 0 });
    s.addText(l, { x: kxy, y: 6.0, w: 2.9, h: 0.55, fontFace: FT, fontSize: 11.5, color: C.ice, margin: 0 });
    kxy += 3.0;
  });
  s.addText(`Atrasos calculados al ${Y.meta.hoy}`, { x: 0.85, y: 6.75, w: 8, h: 0.3, fontFace: FT,
    fontSize: 11, italic: true, color: C.steel, margin: 0 });
  footer(s, true, CAPY);

  // ---------------- Resumen ejecutivo ----------------
  s = p.addSlide(); bg(s, C.paper); logo(s, false);
  kicker(s, "Talabre · Resumen ejecutivo", 0.5, 0.45);
  s.addText("Estado global de los tres frentes", { x: 0.5, y: 0.72, w: 12, h: 0.6, fontFace: FH, fontSize: 30, bold: true, color: C.ink, margin: 0 });

  const extCam = (c) => `${pct(c.realizada, c.total)} del universo · ${nf(c.programada)} programadas en plazo` +
    (c.vencida ? ` · ${nf(c.vencida)} con agenda vencida` : "") + ` · ${nf(c.pendiente)} sin programar`;
  const cardsY = [
    { p: pct(c1.realizada, c1.exigible), v: `${nf(c1.realizada)}/${nf(c1.exigible)} exigibles`, t: "Caminata 1 realizada",
      d: extCam(c1), c: C.blue },
    { p: pct(c2.realizada, c2.exigible), v: `${nf(c2.realizada)}/${nf(c2.exigible)} exigibles`, t: "Caminata 2 realizada",
      d: extCam(c2), c: C.blue },
    { p: `${YP.promedio.toLocaleString("es-CL")}%`, v: `${nf(YP.sobre95)}/${nf(YP.total)} sobre 95%`,
      t: "Avance PEC de carpetas", d: `${nf(YP.bajo80)} subsistemas bajo 80% de avance`, c: C.copper },
    { p: pct(yp1.cerrados, yp1.total), v: `${nf(yp1.cerrados)}/${nf(yp1.total)}`, t: "Cierre de DT P1",
      d: `${nf(yp1.abiertos)} abiertos · ${nf(yp1.atrasados)} atrasados`, c: C.crit },
  ];
  let cxy = 0.5;
  cardsY.forEach((cd) => {
    s.addShape(p.ShapeType.roundRect, { x: cxy, y: 1.6, w: 3.0, h: 2.5, rectRadius: 0.08,
      fill: { color: C.white }, line: { color: "E4E9EF", width: 1 },
      shadow: { type: "outer", color: "9AA7B5", blur: 7, offset: 2, angle: 90, opacity: 0.28 } });
    dot(s, cxy + 0.28, 1.88, cd.c, 0.16);
    s.addText(cd.p, { x: cxy + 0.28, y: 2.12, w: 2.5, h: 0.82, fontFace: FH, fontSize: 38, bold: true, color: cd.c, margin: 0 });
    s.addText(cd.v, { x: cxy + 0.3, y: 2.95, w: 2.5, h: 0.32, fontFace: FT, fontSize: 13, bold: true, color: C.ink, margin: 0 });
    s.addText(cd.t, { x: cxy + 0.3, y: 3.27, w: 2.45, h: 0.32, fontFace: FT, fontSize: 12, color: C.ink2, margin: 0 });
    s.addText(cd.d, { x: cxy + 0.3, y: 3.58, w: 2.45, h: 0.46, fontFace: FT, fontSize: 9.5, color: C.ink2, lineSpacing: 12, margin: 0 });
    cxy += 3.13;
  });

  s.addShape(p.ShapeType.roundRect, { x: 0.5, y: 4.32, w: 12.33, h: 2.42, rectRadius: 0.06, fill: { color: C.navy }, line: { type: "none" } });
  s.addText("LECTURA DEL PERÍODO", { x: 0.85, y: 4.5, w: 6, h: 0.3, fontFace: FT, fontSize: 11, bold: true, color: C.copperL, charSpacing: 2, margin: 0 });
  const pMenores = ["P2", "P3", "P4"].map((k) => YT.porPrioridad[k]).filter(Boolean);
  const abMenores = pMenores.reduce((a, v) => a + v.abiertos, 0);
  const readsY = [
    [`P1 está prácticamente cerrado; el volumen pendiente está en las prioridades menores.`,
      `${pct(yp1.cerrados, yp1.total)} de cierre en P1 (${nf(yp1.abiertos)} abiertos), pero ${nf(abMenores)} DT ` +
      `abiertos entre P2, P3 y P4 sobre un total de ${nf(yg.abiertos)}.`],
    [`El atraso es más antiguo que el de los otros proyectos.`,
      `${nf(yg.atrasados)} DT con la fecha de compromiso vencida; mediana de ${nf(YT.antiguedad.mediana)} días ` +
      `y el más antiguo lleva ${nf(YT.antiguedad.max)}. Los ${nf(yg.sinCompromiso)} abiertos sin fecha de compromiso no se cuentan como atrasados.`],
    [`Caminatas: ${pct(c1.realizada, c1.exigible)} y ${pct(c2.realizada, c2.exigible)} de lo exigible (C1 y C2).`,
      `Caminata 1: ${nf(c1.realizada)}/${nf(c1.exigible)} exigibles` +
      (c1.vencida ? `, ${nf(c1.vencida)} con agenda vencida` : "") +
      `. Caminata 2: ${nf(c2.realizada)}/${nf(c2.exigible)}, ` +
      `con ${nf(c2.programada)} agendadas en plazo y ${nf(c2.pendiente)} sin programar.`],
    [`Las carpetas avanzan bien salvo un grupo acotado.`,
      `Avance promedio ${YP.promedio.toLocaleString("es-CL")}%: ${nf(YP.sobre95)} subsistemas sobre 95% y solo ` +
      `${nf(YP.bajo80)} bajo 80%.` + (YP.enRevisionCliente !== null && YP.enRevisionCliente !== undefined
        ? ` El archivo declara ${nf(YP.enRevisionCliente)} carpetas en revisión del cliente.` : "")],
  ];
  let ryy = 4.86;
  readsY.forEach(([h, d]) => {
    dot(s, 0.9, ryy + 0.06, C.copper, 0.12);
    s.addText([{ text: h + "  ", options: { bold: true, color: C.white } },
               { text: d, options: { color: C.ice } }],
      { x: 1.15, y: ryy - 0.05, w: 11.4, h: 0.5, fontFace: FT, fontSize: 11.5, lineSpacing: 15, margin: 0, valign: "top" });
    ryy += 0.46;
  });
  footer(s, false, CAPY, "Talabre");

  // ---------------- Caminatas y carpetas ----------------
  s = p.addSlide(); bg(s, C.paper); logo(s, false);
  kicker(s, "Talabre · Frentes 1 y 2", 0.5, 0.45);
  s.addText("Caminatas y avance de carpetas", { x: 0.5, y: 0.72, w: 12.3, h: 0.6, fontFace: FH, fontSize: 30, bold: true, color: C.ink, margin: 0 });
  s.addText(`Caminata 2: ${pct(c2.realizada, c2.exigible)} de lo exigible (${nf(c2.realizada)}/${nf(c2.exigible)}) — ` +
    `se descuentan ${nf(c2.programada)} agendadas a fecha futura, que no son incumplimiento al corte. ` +
    `Talabre mide la carpeta como % de avance (PEC), no como estado de aprobación.`,
    { x: 0.5, y: 1.35, w: 12.3, h: 0.5, fontFace: FT, fontSize: 13.5, color: C.ink2, margin: 0 });

  // La serie «Agenda vencida» solo se dibuja si existe: una leyenda con una
  // serie siempre en cero es ruido.
  const hayVencidasC2 = areasY.some((a) => (Y.caminatasPorArea[a]["2"].vencida || 0) > 0);
  s.addChart(p.ChartType.bar, [
    { name: "Realizada", labels: areasY, values: areasY.map((a) => z(Y.caminatasPorArea[a]["2"].realizada)) },
    { name: "Agendada en plazo", labels: areasY, values: areasY.map((a) => z(Y.caminatasPorArea[a]["2"].programada)) },
    ...(hayVencidasC2 ? [{ name: "Agenda vencida", labels: areasY,
        values: areasY.map((a) => z(Y.caminatasPorArea[a]["2"].vencida || 0)) }] : []),
    { name: "Sin programar", labels: areasY, values: areasY.map((a) => z(Y.caminatasPorArea[a]["2"].pendiente)) },
  ], {
    x: 0.5, y: 2.05, w: 6.4, h: 4.5, barDir: "bar", barGrouping: "stacked",
    chartColors: hayVencidasC2 ? [C.goodD, C.blue, C.warnD, C.neutral] : [C.goodD, C.blue, C.neutral],
    showTitle: true, title: "Caminata 2 por área", titleFontFace: FT, titleFontSize: 13, titleColor: C.ink,
    ...ETIQ,
    showLegend: true, legendPos: "b", legendFontFace: FT, legendFontSize: 10.5, legendColor: C.ink2,
    catAxisLabelFontFace: FT, catAxisLabelFontSize: 9.5, catAxisLabelColor: C.ink, catAxisLabelFontBold: true,
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    barGapWidthPct: 40, chartArea: { fill: { color: C.white } },
  });

  s.addChart(p.ChartType.bar, [
    { name: "95% o más", labels: areasY, values: areasY.map((a) => z(YP.porArea[a].sobre95)) },
    { name: "80% a 95%", labels: areasY, values: areasY.map((a) => z(YP.porArea[a].sobre80 - YP.porArea[a].sobre95)) },
    { name: "Bajo 80%", labels: areasY, values: areasY.map((a) => z(YP.porArea[a].bajo80)) },
  ], {
    x: 7.05, y: 2.05, w: 5.78, h: 4.5, barDir: "bar", barGrouping: "stacked",
    chartColors: [C.goodD, C.blue, C.crit],
    showTitle: true, title: "Avance PEC de carpetas por área", titleFontFace: FT, titleFontSize: 13, titleColor: C.ink,
    ...ETIQ,
    showLegend: true, legendPos: "b", legendFontFace: FT, legendFontSize: 10.5, legendColor: C.ink2,
    catAxisLabelFontFace: FT, catAxisLabelFontSize: 9.5, catAxisLabelColor: C.ink, catAxisLabelFontBold: true,
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    barGapWidthPct: 40, chartArea: { fill: { color: C.white } },
  });
  footer(s, false, CAPY, "Talabre");

  // ---------------- Detalles de terminación ----------------
  s = p.addSlide(); bg(s, C.paper); logo(s, false);
  kicker(s, "Talabre · Frente 3", 0.5, 0.45);
  s.addText("Detalles de terminación", { x: 0.5, y: 0.72, w: 12.3, h: 0.6, fontFace: FH, fontSize: 30, bold: true, color: C.ink, margin: 0 });
  s.addText(`${nf(yg.total)} DT levantados; ${pct(yg.cerrados, yg.total)} cerrados. ` +
    `Atrasado = DT abierto cuya fecha de compromiso de cierre ya venció al ${Y.meta.hoy}.`,
    { x: 0.5, y: 1.35, w: 12.3, h: 0.5, fontFace: FT, fontSize: 13.5, color: C.ink2, margin: 0 });

  const lblDY = discOrden.map(([d]) => (Y.discCorta && Y.discCorta[d]) || d);
  s.addChart(p.ChartType.bar, [
    { name: "Atrasados", labels: lblDY, values: discOrden.map(([, v]) => z(v.atrasados)) },
    { name: "En plazo o sin compromiso", labels: lblDY, values: discOrden.map(([, v]) => z(v.enPlazo)) },
  ], {
    x: 0.5, y: 2.05, w: 7.7, h: 4.5, barDir: "bar", barGrouping: "stacked",
    chartColors: [C.crit, C.blue],
    showTitle: true, title: "DT abiertos por disciplina", titleFontFace: FT, titleFontSize: 13, titleColor: C.ink,
    ...ETIQ,
    showLegend: true, legendPos: "b", legendFontFace: FT, legendFontSize: 11, legendColor: C.ink2,
    catAxisLabelFontFace: FT, catAxisLabelFontSize: 11.5, catAxisLabelColor: C.ink, catAxisLabelFontBold: true,
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    barGapWidthPct: 45, chartArea: { fill: { color: C.white } },
  });

  s.addShape(p.ShapeType.roundRect, { x: 8.4, y: 2.05, w: 4.43, h: 4.5, rectRadius: 0.06, fill: { color: C.navy }, line: { type: "none" } });
  s.addText("CIERRE POR PRIORIDAD", { x: 8.68, y: 2.28, w: 4, h: 0.3, fontFace: FT, fontSize: 11, bold: true, color: C.copperL, charSpacing: 1.2, margin: 0 });
  // Paso de 0.56": con 5 prioridades el bloque termina en y≈5.5 y deja sitio
  // a la antigüedad DENTRO del panel; con 0.62 el texto se montaba sobre el pie.
  let pyy = 2.66;
  YT.prioridades.forEach((k) => {
    const v = YT.porPrioridad[k];
    if (!v || !v.total) return;
    const w = 3.9, pctv = pctN(v.cerrados, v.total);
    s.addText(k, { x: 8.68, y: pyy, w: 1.0, h: 0.26, fontFace: FT, fontSize: 11.5, bold: true, color: C.white, margin: 0 });
    s.addText(`${nf(v.cerrados)}/${nf(v.total)}`, { x: 9.6, y: pyy, w: 1.6, h: 0.26, fontFace: FT,
      fontSize: 10.5, color: C.ice, align: "right", margin: 0 });
    s.addText(`${pctv.toLocaleString("es-CL")}%`, { x: 11.25, y: pyy, w: 1.35, h: 0.26, fontFace: FT,
      fontSize: 11, bold: true, color: C.white, align: "right", margin: 0 });
    s.addShape(p.ShapeType.rect, { x: 8.68, y: pyy + 0.28, w: w, h: 0.13, fill: { color: C.navy2 }, line: { type: "none" } });
    if (pctv > 0) s.addShape(p.ShapeType.rect, { x: 8.68, y: pyy + 0.28, w: Math.max(0.03, w * pctv / 100), h: 0.13,
      fill: { color: pctv >= 90 ? C.good : pctv >= 50 ? C.copperL : C.crit }, line: { type: "none" } });
    pyy += 0.56;
  });
  s.addShape(p.ShapeType.line, { x: 8.68, y: pyy + 0.04, w: 3.9, h: 0, line: { color: C.navy2, width: 1 } });
  s.addText([{ text: "Antigüedad del atraso\n", options: { bold: true, color: C.copperL, fontSize: 11.5 } },
             { text: YT.antiguedad.max
                 ? `Mediana de ${nf(YT.antiguedad.mediana)} días vencidos; el más antiguo lleva ${nf(YT.antiguedad.max)}. ` +
                   Object.entries(YT.antiguedad.tramos).filter(([, v]) => v).map(([k, v]) => `${nf(v)} entre ${k}`).join(", ") + "."
                 : "Sin DT atrasados al corte.",
               options: { color: C.ice, fontSize: 11 } }],
    { x: 8.68, y: pyy + 0.16, w: 3.95, h: 0.9, fontFace: FT, lineSpacing: 14, margin: 0, valign: "top" });
  footer(s, true, CAPY, "Talabre");

  // ---------------- Foco de gestión ----------------
  s = p.addSlide(); bg(s, C.navy); logo(s, true);
  s.addShape(p.ShapeType.ellipse, { x: 10.8, y: -1.8, w: 4.6, h: 4.6, fill: { color: C.navy2 }, line: { type: "none" } });
  kicker(s, "Talabre · Foco de gestión", 0.5, 0.5, C.copperL);
  s.addText("Dónde poner el esfuerzo esta semana", { x: 0.5, y: 0.78, w: 12, h: 0.6, fontFace: FH, fontSize: 30, bold: true, color: C.white, margin: 0 });

  const p2 = YT.porPrioridad.P2 || { total: 0, cerrados: 0, abiertos: 0, atrasados: 0 };
  const p4 = YT.porPrioridad.P4 || { total: 0, cerrados: 0, abiertos: 0 };
  const camCrit = areasY.map((a) => [a, Y.caminatasPorArea[a]["2"]])
    .filter(([, c]) => c.total && c.realizada / c.total < 0.6)
    .sort((a, b) => (a[1].realizada / a[1].total) - (b[1].realizada / b[1].total));
  const focusY = [
    { n: "01", t: `Recuperar los ${nf(yg.atrasados)} DT atrasados`, c: C.crit,
      d: `${nf(p2.atrasados)} son P2 y ${nf(yp1.atrasados)} son P1. La mediana de atraso es de ` +
         `${nf(YT.antiguedad.mediana)} días y ${nf((YT.antiguedad.tramos["91-180 días"] || 0) + (YT.antiguedad.tramos[">180 días"] || 0))} ` +
         `superan los 90 días.` },
    { n: "02", t: camCrit.length ? `Ejecutar la caminata 2 en ${camCrit.slice(0, 2).map(([a]) => a).join(" y ")}` : "Completar la caminata 2", c: C.copper,
      d: camCrit.length
         ? camCrit.slice(0, 2).map(([a, c]) => `${a}: ${nf(c.realizada)}/${nf(c.total)} (${pct(c.realizada, c.total)})`).join(" · ") +
           `. A nivel proyecto la caminata 2 va ${pct(c2.realizada, c2.total)} con ${nf(c2.programada)} ya agendadas.`
         : `Van ${nf(c2.realizada)}/${nf(c2.total)} (${pct(c2.realizada, c2.total)}), con ${nf(c2.programada)} agendadas.` },
    { n: "03", t: `Destrabar P4: ${nf(p4.abiertos)} abiertos y ninguno cerrado`, c: C.warn,
      d: `De ${nf(p4.total)} DT levantados en P4 no se ha cerrado ninguno (${pct(p4.cerrados, p4.total)}). ` +
         `P3 tampoco despega: ${pct((YT.porPrioridad.P3 || {}).cerrados || 0, (YT.porPrioridad.P3 || {}).total || 1)} de cierre.` },
    { n: "04", t: areaFocoY ? `Atender ${areaFocoY[0]}` : "Atender el área crítica", c: C.good,
      d: areaFocoY ? `Concentra ${nf(areaFocoY[1].abiertos)} de los ${nf(yg.abiertos)} DT abiertos ` +
         `(${pct(areaFocoY[1].abiertos, yg.abiertos)}), con ${nf(areaFocoY[1].atrasados)} atrasados. ` +
         `Además hay ${nf(YP.bajo80)} carpetas bajo 80% de avance que conviene cerrar en paralelo.` : "" },
  ];
  focusY.forEach((f, i) => {
    const px = 0.5 + (i % 2) * (6.0 + 0.33);
    const py = 1.75 + Math.floor(i / 2) * 2.35;
    s.addShape(p.ShapeType.roundRect, { x: px, y: py, w: 6.0, h: 2.1, rectRadius: 0.07, fill: { color: C.navy2 }, line: { type: "none" } });
    s.addText(f.n, { x: px + 0.3, y: py + 0.22, w: 1.1, h: 0.7, fontFace: FH, fontSize: 34, bold: true, color: f.c, margin: 0 });
    s.addText(f.t, { x: px + 1.35, y: py + 0.2, w: 6.0 - 1.6, h: 0.66, fontFace: FT, fontSize: 13.5, bold: true, color: C.white, margin: 0, valign: "middle" });
    s.addText(f.d, { x: px + 1.35, y: py + 0.88, w: 6.0 - 1.65, h: 1.1, fontFace: FT, fontSize: 11, color: C.ice, margin: 0, lineSpacing: 14, valign: "top" });
  });
  footer(s, true, CAPY, "Talabre");

  console.log(`Talabre: 5 láminas agregadas (atrasos al ${Y.meta.hoy}).`);
} else {
  console.log("No se encontró datos_talabre.json — la PPT sale sin Talabre.");
}

// =====================================================================
// Portada · Proyecto Arqueros
// =====================================================================
s = p.addSlide();
portada(s, "03 · P2342", "Arqueros", PROY_MASA,
  `${nf(cam.subs)} subsistemas en ${D.areas.length} áreas (${D.areas.join(" · ")}) · ` +
  `${nf(D.dt.global.total)} detalles de terminación levantados.\n` +
  "Estatus de armado y entrega de carpetas TOP · avance de caminatas 80% / 100% · " +
  "detalles de construcción en condición P1.",
  [[pct(cam.c80, cam.subs), "Caminatas 80%"],
   [pct(cam.c100, cam.subs), "Caminatas 100%"],
   [pct(top.entregadas, top.total), "Carpetas TOP entregadas"],
   [pct(P1.cerrados, P1.total), "Cierre detalles P1"]],
  `Corte: ${D.meta.corte_texto}` + (PREV ? `   ·   Corte anterior: ${PREV.corte_texto}` : ""),
  CAP);

// =====================================================================
// Arqueros · Resumen ejecutivo
// =====================================================================
s = p.addSlide(); bg(s, C.paper); logo(s, false);
kicker(s, "Arqueros · Resumen ejecutivo", 0.5, 0.45);
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
footer(s, false, CAP, "Arqueros");

// =====================================================================
// Arqueros · Semáforo por área
// =====================================================================
s = p.addSlide(); bg(s, C.paper); logo(s, false);
kicker(s, "Arqueros · Estado por área", 0.5, 0.45);
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
footer(s, false, CAP, "Arqueros");

// =====================================================================
// Arqueros · Caminatas
// =====================================================================
s = p.addSlide(); bg(s, C.paper); logo(s, false);
kicker(s, "Arqueros · Frente 1 · Caminatas", 0.5, 0.45);
s.addText("Avance de caminatas 80% y 100%", { x: 0.5, y: 0.72, w: 12, h: 0.6, fontFace: FH, fontSize: 30, bold: true, color: C.ink, margin: 0 });
s.addText(`Las caminatas al 80% están cerradas en todas las áreas (${nf(cam.c80)}/${nf(cam.subs)}). ` +
  `El hito 100% avanza a ${pct(cam.c100, cam.subs)}` +
  (areaCamFoco[0] ? `, y el área ${areaCamFoco[0]} concentra ${nf(areaCamFoco[1])} de los ${nf(D.cam.pend.length)} subsistemas pendientes.` : "."),
  { x: 0.5, y: 1.35, w: 12.3, h: 0.5, fontFace: FT, fontSize: 13.5, color: C.ink2, margin: 0 });

s.addChart(p.ChartType.bar, [
  { name: "Realizada", labels: D.areas, values: D.areas.map((a) => z(D.cam.byArea[a].c100)) },
  { name: "Próxima a realizar", labels: D.areas, values: D.areas.map((a) => z(D.cam.byArea[a].prox)) },
  { name: "Por programar", labels: D.areas, values: D.areas.map((a) => z(D.cam.byArea[a].prog)) },
], {
  x: 0.5, y: 2.05, w: 7.6, h: 4.5, barDir: "bar", barGrouping: "stacked",
  chartColors: [C.goodD, C.blue, C.blueM],
  showTitle: true, title: "Subsistemas por estado de caminata 100%", titleFontFace: FT, titleFontSize: 13, titleColor: C.ink,
  ...ETIQ,
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
footer(s, false, CAP, "Arqueros");

// =====================================================================
// Arqueros · Carpetas TOP
// =====================================================================
s = p.addSlide(); bg(s, C.paper); logo(s, false);
kicker(s, "Arqueros · Frente 2 · Carpetas TOP", 0.5, 0.45);
s.addText("Estatus de armado y entrega", { x: 0.5, y: 0.72, w: 12, h: 0.6, fontFace: FH, fontSize: 30, bold: true, color: C.ink, margin: 0 });
s.addText(`${nf(top.entregadas)} carpetas entregadas de ${nf(top.total)} (${pct(top.entregadas, top.total)}). ` +
  `${top.aprob === 0 ? "Ninguna aprobada aún" : `${nf(top.aprob)} aprobadas`}: el flujo se acumula en revisión y hay ` +
  `${nf(top.rech)} rechazos que requieren rearmado.`,
  { x: 0.5, y: 1.35, w: 12.3, h: 0.5, fontFace: FT, fontSize: 13.5, color: C.ink2, margin: 0 });

const serieTop = [
  { name: "Aprobada", key: "aprob", color: C.goodD },
  { name: "En revisión", key: "rev", color: C.blue },
  { name: "Observada", key: "obs", color: C.warnD },
  { name: "Rechazada", key: "rech", color: C.crit },
].filter((x) => D.areas.some((a) => D.ctop.byArea[a][x.key] > 0));
serieTop.push({ name: "Sin entregar", key: "_pend", color: C.neutral });

s.addChart(p.ChartType.bar, serieTop.map((x) => ({
  name: x.name, labels: D.areas,
  values: D.areas.map((a) => {
    const t = D.ctop.byArea[a];
    return z(x.key === "_pend" ? t.total - (t.rev + t.obs + t.rech + t.aprob) : t[x.key]);
  }),
})), {
  x: 0.5, y: 2.05, w: 7.6, h: 4.5, barDir: "bar", barGrouping: "stacked",
  chartColors: serieTop.map((x) => x.color),
  showTitle: true, title: "Carpetas TOP por área y estatus", titleFontFace: FT, titleFontSize: 13, titleColor: C.ink,
  ...ETIQ,
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
footer(s, false, CAP, "Arqueros");

// =====================================================================
// Arqueros · Detalles P1
// =====================================================================
s = p.addSlide(); bg(s, C.paper); logo(s, false);
kicker(s, "Arqueros · Frente 3 · Detalles P1", 0.5, 0.45);
s.addText("Detalles de terminación en condición P1", { x: 0.5, y: 0.72, w: 12.3, h: 0.6, fontFace: FH, fontSize: 30, bold: true, color: C.ink, margin: 0 });
s.addText(`${nf(P1.total)} P1 levantados; ${pct(P1.cerrados, P1.total)} cerrados. ` +
  `El pendiente se concentra en ${espOrden.slice(0, 3).map((x) => x[0]).join(", ")}, y físicamente en el área ${areaFoco}.`,
  { x: 0.5, y: 1.35, w: 12.3, h: 0.5, fontFace: FT, fontSize: 13.5, color: C.ink2, margin: 0 });

const espLbl = espOrden.map(([e]) => (e === "Instrumentación y Control" ? "Instrum. y Control" : e));
s.addChart(p.ChartType.bar, [
  { name: "Vencidos al corte", labels: espLbl, values: espOrden.map(([, v]) => z(v.vencidos)) },
  { name: "En plazo", labels: espLbl, values: espOrden.map(([, v]) => z(v.abiertos - v.vencidos)) },
], {
  x: 0.5, y: 2.05, w: 7.7, h: 4.5, barDir: "bar", barGrouping: "stacked",
  chartColors: [C.crit, C.blue],
  showTitle: true, title: "P1 abiertos por disciplina", titleFontFace: FT, titleFontSize: 13, titleColor: C.ink,
  ...ETIQ,
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
footer(s, true, CAP, "Arqueros");

// =====================================================================
// Arqueros · Foco de gestión
// =====================================================================
s = p.addSlide(); bg(s, C.navy); logo(s, true);
s.addShape(p.ShapeType.ellipse, { x: 10.8, y: -1.8, w: 4.6, h: 4.6, fill: { color: C.navy2 }, line: { type: "none" } });
kicker(s, "Arqueros · Foco de gestión", 0.5, 0.5, C.copperL);
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
footer(s, true, CAP, "Arqueros");

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
// oscuro y la de color para el modo claro.
function bloqueLogos() {
  const uri = (f) => {
    const r = path.join(AQUI, f);
    return fs.existsSync(r) ? "data:image/png;base64," + fs.readFileSync(r).toString("base64") : null;
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
  // El nombre del archivo descargado lleva el corte más reciente de los tres
  // proyectos, en DDMMAA, para no pisar informes de semanas distintas en la
  // carpeta de descargas de quien lo abra.
  const cortes = [D.meta.corte_texto];
  for (const r of ["datos_desaladora.json", "datos_talabre.json"]) {
    const f2 = path.join(AQUI, r);
    if (!fs.existsSync(f2)) continue;
    const j = JSON.parse(fs.readFileSync(f2, "utf8"));
    // «hoy» es la fecha de referencia real del corte; corteTexto puede ser una
    // fecha de agenda futura en algún proyecto, así que se prefiere hoy.
    if (j.meta) cortes.push(j.meta.hoy || j.meta.corteTexto);
  }
  const aIso = (t) => { const [d, m, a] = String(t).split("-"); return `${a}${m}${d}`; };
  const corteMax = cortes.filter(Boolean).sort((a, b) => aIso(a).localeCompare(aIso(b))).pop();
  const [dd, mm, aaaa] = String(corteMax || "").split("-");
  const meta = {
    nombre: `Panel_CRP_${dd && mm && aaaa ? dd + mm + aaaa.slice(2) : "actual"}.pptx`,
    corte: corteMax,
    laminas: nSlide - 1,
    peso: `${(bin.length / 1024 / 1024).toFixed(1)} MB`,
    b64: bin.toString("base64"),
  };
  const bloque = MARCA_INI + " (la PPT del corte, embebida por gen_ppt.js — no editar a mano) -->\n" +
    "<script>\nconst PPTX = " + JSON.stringify(meta) + ";\n<\/script>\n";
  fs.writeFileSync(html, reemplazar(txt, MARCA_INI, MARCA_FIN, bloque), "utf8");
  console.log(`Informe y logos embebidos en index.html (${meta.nombre} · ${meta.laminas} láminas · ${meta.peso}).`);
}

const salida = path.join(AQUI, "Panel_Control_TOP_P1.pptx");
p.writeFile({ fileName: salida })
  .then((f) => { console.log("PPT generada:", f); embeber(salida); })
  .catch((e) => { console.error(e); process.exit(1); });
