// Referto "Medical report" per il cliente BMCh24: stessi campi del modello originale, impaginazione pulita.
import { sanitizeName } from "./indexer.js";
import { imageRun, itDate, todayISO } from "./docgen.js";

export const DEFAULT_BMC = { logoData: "", titolo: "MEDICAL REPORT", direzione: "Dr Antonio Magliocca", medico: "", luogo: "" };

export const EMPTY_BMC_DOC = () => ({
  numero: "", paziente: "", nascita: "", dataVisita: todayISO(), luogoVisita: "", medico: "",
  motivo: "", anamnesi: "", pa1: "", pa2: "", fc: "", temp: "", spo2: "",
  torace: "", addome: "", arti: "", altro: "", diagnosi: "", prognosi: "", terapia: "", extra: "",
  consensoNome: "", firmaNote: "", conFirma: true,
});

// Testi fissi del modello originale
const CONSENSO_1 = "Presa visione dell’informativa di cui all’Art. 13 del Reg. UE 679/2016, il paziente:";
const CONSENSO_2 = "Rilascia il consenso al trattamento dei dati personali, anche sensibili, per finalità amministrative e di cura";
const CONSENSO_3 = "(i dati saranno trattati per l’adeguata erogazione delle prestazioni sanitarie, nonché per il corretto assolvimento degli adempimenti amministrativi, burocratici, fiscali, e per lo svolgimento di tutte le attività inerenti al rapporto in essere).";

export function bmcFileName(d) {
  return `${(d.dataVisita || todayISO()).replace(/-/g, "_")}_BMC Medical report_${sanitizeName(d.paziente || "Paziente")}.docx`;
}

const lines = (s) => String(s || "").split("\n");
const vitals = (d) => [
  `P.A. ${d.pa1 || "____"}/${d.pa2 || "____"} mmHg`,
  `F.C. ${d.fc || "____"} /min`,
  `T°: ${d.temp ? d.temp + " °C" : "____ °C"}`,
  `SpO2: ${d.spo2 ? d.spo2 + " %" : "____ %"}`,
];
const visitLine = (d) => [itDate(d.dataVisita), d.luogoVisita].filter(Boolean).join(", ");

// ---------------------------------------------------------------- Word
export async function buildBmcDocx(d, bmc, profile) {
  const D = docx;
  const NONE = { style: D.BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const noB = { top: NONE, bottom: NONE, left: NONE, right: NONE };
  const rule = (color = "999999") => ({ bottom: { style: D.BorderStyle.SINGLE, size: 6, space: 2, color } });
  const P = (children, o = {}) => new D.Paragraph({ children, spacing: { before: o.before || 0, after: o.after ?? 80 }, alignment: o.align, keepNext: o.keepNext, keepLines: o.keepLines, border: o.border });
  const B = (t, size) => new D.TextRun({ text: t, bold: true, size });
  const T = (t, o = {}) => new D.TextRun({ text: t, ...o });
  const block = (label, val) => {
    const out = [P([B(label)], { before: 120, after: 40, keepNext: true })];
    const ls = lines(val);
    if (!String(val || "").trim()) out.push(P([T(" ")], { border: rule("BBBBBB"), after: 120 }), P([T(" ")], { border: rule("BBBBBB"), after: 120 }));
    else ls.forEach((t) => out.push(P([T(t)], { after: 60 })));
    return out;
  };
  const cellOf = (paras, w) => new D.TableCell({ borders: noB, width: { size: w, type: D.WidthType.PERCENTAGE }, children: paras });
  const children = [];

  if (bmc.logoData) children.push(P([await imageRun(bmc.logoData, 440, 70)], { align: D.AlignmentType.CENTER, after: 160 }));
  children.push(P([B(bmc.titolo || "MEDICAL REPORT", 28)], { align: D.AlignmentType.CENTER, after: 120, border: rule("444444") }));

  children.push(new D.Table({ width: { size: 100, type: D.WidthType.PERCENTAGE }, borders: D.TableBorders.NONE, rows: [
    new D.TableRow({ children: [
      cellOf([P([B("Numero incarico: "), T(d.numero)])], 40),
      cellOf([P([B("Paz.: "), T(d.paziente)])], 60)] })] }));
  children.push(P([B("Data di nascita: "), T(d.nascita ? itDate(d.nascita) : "")]));
  children.push(P([B("Data e luogo visita: "), T(visitLine(d))]));
  children.push(P([B("Medico incaricato: "), T(d.medico)]));

  children.push(...block("Motivo della chiamata:", d.motivo));
  children.push(...block("Anamnesi/Allergie:", d.anamnesi));
  children.push(P([B("Parametri vitali all'ingresso:")], { before: 120, after: 60, keepNext: true }));
  children.push(new D.Table({ width: { size: 100, type: D.WidthType.PERCENTAGE }, borders: D.TableBorders.NONE,
    rows: [new D.TableRow({ children: vitals(d).map((v) => cellOf([P([T(v, { bold: true })])], 25)) })] }));

  children.push(P([B("Esame obiettivo:", 26)], { align: D.AlignmentType.CENTER, before: 240, after: 120, keepNext: true }));
  [["Torace", d.torace], ["Addome", d.addome], ["Arti inferiori", d.arti], ["Altro", d.altro]].forEach(([l, v]) => {
    children.push(P([B(`- ${l}: `), T(lines(v).join("  "))], { border: String(v || "").trim() ? undefined : rule("BBBBBB"), after: 100 }));
  });

  children.push(...block("DIAGNOSI:", d.diagnosi));
  children.push(...block("PROGNOSI:", d.prognosi));
  children.push(...block("Conclusione/ TERAPIA:", d.terapia));
  if (String(d.extra || "").trim()) children.push(...block("Note aggiuntive:", d.extra));

  if (d.conFirma && (profile.firmaData || profile.timbroData)) {
    const imgs = [];
    if (profile.timbroData) imgs.push(await imageRun(profile.timbroData, 140, 90));
    if (profile.firmaData) imgs.push(await imageRun(profile.firmaData, 180, 80));
    children.push(P(imgs.flatMap((r, i) => (i ? [T("   "), r] : [r])), { align: D.AlignmentType.RIGHT, before: 160, after: 0, keepNext: true }));
    if (d.medico) children.push(P([T(d.medico)], { align: D.AlignmentType.RIGHT, after: 120 }));
  }
  children.push(P([B("DIREZIONE SANITARIA " + (bmc.direzione || ""))], { before: 240, after: 240, keepNext: true }));

  children.push(P([T(CONSENSO_1)], { keepNext: true, keepLines: true }));
  children.push(P([T(d.consensoNome || "________________", { bold: !!d.consensoNome }), T(" (cognome e nome)", { italics: true })], { before: 60, keepNext: true }));
  children.push(P([T(CONSENSO_2)], { before: 120, keepNext: true, keepLines: true }));
  children.push(P([T(CONSENSO_3, { size: 18 })], { keepNext: true, keepLines: true }));
  children.push(P([T("Firma del paziente" + (d.firmaNote ? ` (${d.firmaNote})` : ""), { size: 24 })], { before: 160, after: 360, keepNext: true }));
  children.push(P([T("_________________________")]));

  const doc = new D.Document({
    creator: d.medico || "Gestionale",
    styles: { default: { document: { run: { font: "Calibri", size: 22 } } } },
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 1000, left: 1250, right: 1250 } } },
      footers: { default: new D.Footer({ children: [new D.Paragraph({ alignment: D.AlignmentType.CENTER, children: [new D.TextRun({ children: [D.PageNumber.CURRENT], size: 18 })] })] }) },
      children,
    }],
  });
  return D.Packer.toBlob(doc);
}

// ---------------------------------------------------------------- HTML (anteprima + stampa/PDF)
const CSS = `
@page{size:A4;margin:16mm 18mm}
.bmc{font-family:Calibri,'Segoe UI',Arial,sans-serif;font-size:11pt;line-height:1.35;color:#111}
.bmc .logo{text-align:center;margin:0 0 8px}.bmc .logo img{width:11.5cm;max-width:100%}
.bmc h1{font-size:14pt;text-align:center;margin:0 0 8px;padding-bottom:5px;border-bottom:1.5px solid #444}
.bmc .row2{display:flex;gap:12px}.bmc .row2 div:first-child{flex:4}.bmc .row2 div:last-child{flex:6}
.bmc p{margin:0 0 4px}.bmc .h{font-weight:700;margin:9px 0 2px;break-after:avoid}
.bmc .v{white-space:pre-wrap;margin:0 0 4px}.bmc .blank{border-bottom:1px solid #bbb;height:20px;margin-bottom:6px}
.bmc .vit{display:flex;justify-content:space-between;font-weight:700;margin:2px 0}
.bmc h2{font-size:13pt;text-align:center;margin:14px 0 6px;break-after:avoid}
.bmc .it{margin:0 0 5px}.bmc .it.blankline{border-bottom:1px solid #bbb;padding-bottom:2px}
.bmc .sig{text-align:right;margin-top:10px;break-after:avoid}.bmc .sig img{max-height:80px;max-width:45%;vertical-align:middle}
.bmc .dir{font-weight:700;margin:16px 0}
.bmc .cons{break-inside:avoid}.bmc .cons .small{font-size:9pt}.bmc .line{margin-top:26px}
`;

export function bmcHtml(d, bmc, profile, esc) {
  const val = (v) => (String(v || "").trim() ? `<div class="v">${esc(v)}</div>` : `<div class="blank"></div><div class="blank"></div>`);
  const block = (l, v) => `<div class="h">${esc(l)}</div>${val(v)}`;
  const item = (l, v) => `<div class="it${String(v || "").trim() ? "" : " blankline"}"><b>- ${esc(l)}:</b> ${esc(lines(v).join("  "))}</div>`;
  const sig = d.conFirma && (profile.firmaData || profile.timbroData)
    ? `<div class="sig">${profile.timbroData ? `<img src="${profile.timbroData}" alt="timbro"> ` : ""}${profile.firmaData ? `<img src="${profile.firmaData}" alt="firma">` : ""}${d.medico ? `<div>${esc(d.medico)}</div>` : ""}</div>` : "";
  return `<style>${CSS}</style><div class="sheet"><div class="bmc">
    ${bmc.logoData ? `<div class="logo"><img src="${bmc.logoData}" alt="BMCh24"></div>` : ""}
    <h1>${esc(bmc.titolo || "MEDICAL REPORT")}</h1>
    <div class="row2"><div><b>Numero incarico:</b> ${esc(d.numero)}</div><div><b>Paz.:</b> ${esc(d.paziente)}</div></div>
    <p><b>Data di nascita:</b> ${esc(d.nascita ? itDate(d.nascita) : "")}</p>
    <p><b>Data e luogo visita:</b> ${esc(visitLine(d))}</p>
    <p><b>Medico incaricato:</b> ${esc(d.medico)}</p>
    ${block("Motivo della chiamata:", d.motivo)}
    ${block("Anamnesi/Allergie:", d.anamnesi)}
    <div class="h">Parametri vitali all'ingresso:</div>
    <div class="vit">${vitals(d).map((v) => `<span>${esc(v)}</span>`).join("")}</div>
    <h2>Esame obiettivo:</h2>
    ${item("Torace", d.torace)}${item("Addome", d.addome)}${item("Arti inferiori", d.arti)}${item("Altro", d.altro)}
    ${block("DIAGNOSI:", d.diagnosi)}
    ${block("PROGNOSI:", d.prognosi)}
    ${block("Conclusione/ TERAPIA:", d.terapia)}
    ${String(d.extra || "").trim() ? block("Note aggiuntive:", d.extra) : ""}
    ${sig}
    <div class="dir">DIREZIONE SANITARIA ${esc(bmc.direzione || "")}</div>
    <div class="cons">
      <p>${esc(CONSENSO_1)}</p>
      <p>${d.consensoNome ? `<b>${esc(d.consensoNome)}</b>` : "________________"} <i>(cognome e nome)</i></p>
      <p style="margin-top:8px">${esc(CONSENSO_2)}</p>
      <p class="small">${esc(CONSENSO_3)}</p>
      <p style="margin-top:10px;font-size:12pt">Firma del paziente${d.firmaNote ? ` (${esc(d.firmaNote)})` : ""}</p>
      <p class="line">_________________________</p>
    </div></div></div>`;
}

export function printHtml(html, title) {
  const w = window.open("", "_blank");
  if (!w) return false;
  w.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>${title || "Documento"}</title><meta name="viewport" content="width=device-width,initial-scale=1"></head><body style="margin:0">${html}</body></html>`);
  w.document.close();
  setTimeout(() => { w.focus(); w.print(); }, 500);
  return true;
}
