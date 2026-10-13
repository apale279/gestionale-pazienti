// Generazione PDF (testo vero, non immagine) per documento generico e referto BMC, con la libreria pdfmake.
import { mergeLayout, headerLines, itDate, imgSize } from "./docgen.js";
import { lines, vitals, visitLine, CONSENSO_1, CONSENSO_2, CONSENSO_3 } from "./bmcdoc.js";

const A4W = 595.28;
const PT = 0.75; // px → pt

let loading;
function loadScript(src) {
  return new Promise((res, rej) => {
    const s = document.createElement("script");
    s.src = src; s.onload = res; s.onerror = () => rej(new Error("Impossibile caricare " + src));
    document.head.append(s);
  });
}
// Le librerie PDF (circa 2 MB) si scaricano solo la prima volta che serve un PDF.
export function ensurePdf() {
  if (window.pdfMake && window.pdfMake.vfs) return Promise.resolve();
  return (loading ||= (async () => {
    await loadScript("vendor/pdfmake.min.js");
    await loadScript("vendor/vfs_fonts.js");
  })().catch((e) => { loading = null; throw e; }));
}

function toBlob(def) {
  return new Promise((res, rej) => { try { pdfMake.createPdf(def).getBlob(res); } catch (e) { rej(e); } });
}
const blankLine = (w, y = 14) => ({ canvas: [{ type: "line", x1: 0, y1: y, x2: w, y2: y, lineWidth: 0.6, lineColor: "#bbbbbb" }], margin: [0, 0, 0, 2] });

// ---------------------------------------------------------------- documento generico
export async function buildPdf(profile, d, layout) {
  await ensurePdf();
  const L = mergeLayout(layout), p = profile;
  const [mt, mb, ml, mr] = { narrow: [36, 36, 45, 45], normal: [50, 50, 65, 65], wide: [70, 70, 85, 85] }[L.margin] || [50, 50, 65, 65];
  const cw = A4W - ml - mr;
  const textLines = headerLines(p, L).map((l) => ({ text: l.t, fontSize: l.size, bold: !!l.bold, color: l.gray ? "#555555" : "#000000", alignment: L.hAlign }));
  const textH = headerLines(p, L).reduce((a, l) => a + l.size * 1.3, 0);

  let logo = null, logoH = 0, logoW = 0;
  if (p.logoData && L.logoPos !== "none") {
    const { w, h } = await imgSize(p.logoData);
    const k = Math.min(L.logoW / w, L.logoW / h);
    logoW = w * k * PT; logoH = h * k * PT;
    logo = { image: p.logoData, width: logoW, height: logoH };
  }
  let head, headH;
  if (logo && L.logoPos === "sopra") { head = { stack: [{ ...logo, alignment: L.hAlign, margin: [0, 0, 0, 3] }, ...textLines] }; headH = logoH + 3 + textH; }
  else if (logo) {
    const cols = [{ width: logoW + 8, stack: [logo] }, { width: "*", stack: textLines }];
    head = { columns: L.logoPos === "destra" ? cols.reverse() : cols, columnGap: 8 }; headH = Math.max(logoH, textH);
  } else { head = { stack: textLines }; headH = textH; }
  const rule = L.line ? { canvas: [{ type: "line", x1: 0, y1: 3, x2: cw, y2: 3, lineWidth: 1.2, lineColor: L.lineColor }] } : { text: "" };
  const headerBlock = { margin: [ml, 22, mr, 0], stack: [head, rule] };
  const topMargin = Math.max(mt, headH + 22 + 18);

  const body = [];
  const title = (d.titolo || d.tipo || "").toUpperCase();
  if (title) body.push({ text: title, bold: true, fontSize: L.bodySize + 3, alignment: L.titleAlign === "left" ? "left" : "center", margin: [0, 8, 0, 12] });
  body.push({ text: [{ text: "Paziente: ", bold: true }, d.paziente || ""], margin: [0, 0, 0, 3] });
  for (const [l, v] of d.extraLines || []) body.push({ text: [{ text: l + ": ", bold: true }, String(v)], margin: [0, 0, 0, 3] });
  body.push({ text: [{ text: "Data: ", bold: true }, itDate(d.data)], margin: [0, 0, 0, 12] });
  for (const t of (d.testo || "").split("\n")) body.push({ text: t || " ", alignment: L.justify ? "justify" : "left", margin: [0, 0, 0, 5] });

  const sig = [];
  sig.push({ text: [p.luogo, itDate(d.data)].filter(Boolean).join(", "), alignment: L.sigAlign, margin: [0, 22, 0, 4] });
  const hasF = !!p.firmaData, hasT = !!p.timbroData;
  if (hasF && hasT) {
    const stamp = { width: "*", stack: [{ image: p.timbroData, fit: [L.stampW * PT, 90 * PT], alignment: L.stampSide === "right" ? "right" : "left" }] };
    const sign = { width: "*", stack: [{ image: p.firmaData, fit: [L.sigW * PT, 80 * PT], alignment: L.stampSide === "right" ? "left" : "right" }] };
    sig.push({ columns: L.stampSide === "right" ? [sign, stamp] : [stamp, sign] });
  } else if (hasF || hasT) {
    sig.push({ image: hasF ? p.firmaData : p.timbroData, fit: [(hasF ? L.sigW : L.stampW) * PT, 90 * PT], alignment: L.sigAlign });
  }
  if (p.nome) sig.push({ text: p.nome, alignment: L.sigAlign, margin: [0, 2, 0, 0] });
  body.push({ unbreakable: true, stack: sig });

  return toBlob({
    pageSize: "A4",
    pageMargins: [ml, topMargin, mr, Math.max(mb, 40)],
    info: { title: title || "Documento", author: p.nome || "" },
    defaultStyle: { fontSize: L.bodySize, lineHeight: 1.15 },
    header: () => headerBlock,
    footer: () => (p.piede ? { text: p.piede, alignment: "center", fontSize: 8, color: "#666666", margin: [ml, 0, mr, 18] } : { text: "" }),
    content: body,
  });
}

// ---------------------------------------------------------------- referto BMC
export async function buildBmcPdf(d, bmc, profile) {
  await ensurePdf();
  const ml = 56, mr = 56, cw = A4W - ml - mr;
  const content = [];
  const B = (t) => ({ text: t, bold: true });
  const section = (label, val) => {
    const has = String(val || "").trim();
    const first = has ? { text: lines(val)[0] || " ", margin: [0, 0, 0, 3] } : blankLine(cw);
    const rest = has ? lines(val).slice(1).map((t) => ({ text: t || " ", margin: [0, 0, 0, 3] })) : [blankLine(cw)];
    return [{ unbreakable: true, stack: [{ text: label, bold: true, margin: [0, 9, 0, 2] }, first] }, ...rest];
  };

  if (bmc.logoData) content.push({ image: bmc.logoData, width: 330, alignment: "center", margin: [0, 0, 0, 8] });
  content.push({ text: bmc.titolo || "MEDICAL REPORT", bold: true, fontSize: 14, alignment: "center", margin: [0, 0, 0, 3] });
  content.push({ canvas: [{ type: "line", x1: 0, y1: 2, x2: cw, y2: 2, lineWidth: 1.2, lineColor: "#444444" }], margin: [0, 0, 0, 8] });
  content.push({ columns: [{ width: "40%", text: [B("Numero incarico: "), d.numero || ""] }, { width: "60%", text: [B("Paz.: "), d.paziente || ""] }], margin: [0, 0, 0, 3] });
  content.push({ text: [B("Data di nascita: "), d.nascita ? itDate(d.nascita) : ""], margin: [0, 0, 0, 3] });
  content.push({ text: [B("Data e luogo visita: "), visitLine(d)], margin: [0, 0, 0, 3] });
  content.push({ text: [B("Medico incaricato: "), d.medico || ""], margin: [0, 0, 0, 3] });

  content.push(...section("Motivo della chiamata:", d.motivo));
  content.push(...section("Anamnesi/Allergie:", d.anamnesi));
  content.push({ unbreakable: true, stack: [{ text: "Parametri vitali all'ingresso:", bold: true, margin: [0, 9, 0, 3] },
    { columns: vitals(d).map((v) => ({ width: "*", text: v, bold: true })) }] });

  content.push({ text: "Esame obiettivo:", bold: true, fontSize: 13, alignment: "center", margin: [0, 14, 0, 6] });
  for (const [l, v] of [["Torace", d.torace], ["Addome", d.addome], ["Arti inferiori", d.arti], ["Altro", d.altro]]) {
    const has = String(v || "").trim();
    content.push({ unbreakable: true, stack: [{ text: [B(`- ${l}: `), lines(v).join("  ")], margin: [0, 0, 0, has ? 5 : 0] }, ...(has ? [] : [blankLine(cw, 6)])] });
  }
  content.push(...section("DIAGNOSI:", d.diagnosi));
  content.push(...section("PROGNOSI:", d.prognosi));
  content.push(...section("Conclusione/ TERAPIA:", d.terapia));
  if (String(d.extra || "").trim()) content.push(...section("Note aggiuntive:", d.extra));

  if (d.conFirma && (profile.firmaData || profile.timbroData)) {
    const imgs = [];
    if (profile.timbroData) imgs.push({ image: profile.timbroData, fit: [140 * PT, 90 * PT] });
    if (profile.firmaData) imgs.push({ image: profile.firmaData, fit: [180 * PT, 80 * PT] });
    content.push({ unbreakable: true, margin: [0, 12, 0, 0], stack: [
      { columns: [{ width: "*", text: "" }, ...imgs.map((im) => ({ width: "auto", stack: [im], margin: [8, 0, 0, 0] }))] },
      ...(d.medico ? [{ text: d.medico, alignment: "right", margin: [0, 3, 0, 0] }] : [])] });
  }
  content.push({ text: "DIREZIONE SANITARIA " + (bmc.direzione || ""), bold: true, margin: [0, 18, 0, 18] });

  content.push({ unbreakable: true, stack: [
    { text: CONSENSO_1 },
    { text: [{ text: d.consensoNome || "________________", bold: !!d.consensoNome }, { text: " (cognome e nome)", italics: true }], margin: [0, 4, 0, 8] },
    { text: CONSENSO_2 },
    { text: CONSENSO_3, fontSize: 9, margin: [0, 2, 0, 0] },
    { text: "Firma del paziente" + (d.firmaNote ? ` (${d.firmaNote})` : ""), fontSize: 12, margin: [0, 12, 0, 22] },
    { text: "_________________________" }] });

  return toBlob({
    pageSize: "A4",
    pageMargins: [ml, 48, mr, 56],
    info: { title: "Medical report", author: d.medico || "" },
    defaultStyle: { fontSize: 11, lineHeight: 1.15 },
    footer: (page) => ({ text: String(page), alignment: "center", fontSize: 9, margin: [0, 18, 0, 0] }),
    content,
  });
}
