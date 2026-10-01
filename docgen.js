// Generazione del documento Word con carta intestata, firma e timbro (tutto nel browser).
// Il layout (allineamenti, font, dimensioni…) è configurabile dall'editor "Layout referti".
import { sanitizeName } from "./indexer.js";

export const EMPTY_PROFILE = {
  nome: "", qualifica: "", indirizzo: "", telefono: "", email: "", extra: "",
  luogo: "", piede: "", firmaData: "", timbroData: "", logoData: "",
};

export const DEFAULT_LAYOUT = {
  hAlign: "left",          // allineamento intestazione: left | center | right
  logoPos: "none",         // none | sopra | sinistra | destra
  logoW: 90,               // larghezza logo (px)
  show: { qualifica: true, indirizzo: true, telefono: true, email: true, extra: true },
  contactMode: "inline",   // inline: indirizzo · telefono · email su una riga; lines: una riga ciascuno
  nameSize: 16, subSize: 11, infoSize: 9,
  line: true, lineColor: "#444444",
  font: "Calibri", bodySize: 11, justify: true,
  margin: "normal",        // narrow | normal | wide
  titleAlign: "center",    // center | left
  sigAlign: "right",       // allineamento di luogo/data e firma: left | center | right
  stampSide: "left",       // con firma e timbro insieme: timbro a sinistra (left) o a destra (right)
  sigW: 180, stampW: 140,  // larghezze massime (px)
};

export const FONTS = ["Calibri", "Arial", "Times New Roman", "Georgia", "Verdana", "Cambria"];
const MARGINS = { narrow: [720, 720, 900, 900], normal: [1000, 1000, 1300, 1300], wide: [1400, 1400, 1700, 1700] }; // top,bottom,left,right (twips)

export const mergeLayout = (l) => ({ ...DEFAULT_LAYOUT, ...(l || {}), show: { ...DEFAULT_LAYOUT.show, ...((l || {}).show || {}) } });

export const itDate = (iso) => { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || ""); return m ? `${m[3]}/${m[2]}/${m[1]}` : iso || ""; };
export const todayISO = () => { const d = new Date(), p = (n) => String(n).padStart(2, "0"); return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`; };

export function docFileName(d) {
  const day = (d.data || todayISO()).replace(/-/g, "_");
  return `${day}_${sanitizeName(d.tipo || "Documento")}_${sanitizeName(d.paziente || "Paziente")}.docx`;
}

function dataUrlToBytes(url) {
  const b = atob(url.split(",")[1]);
  const u = new Uint8Array(b.length);
  for (let i = 0; i < b.length; i++) u[i] = b.charCodeAt(i);
  return u;
}
function imgSize(url) {
  return new Promise((res) => { const i = new Image(); i.onload = () => res({ w: i.naturalWidth, h: i.naturalHeight }); i.onerror = () => res({ w: 1, h: 1 }); i.src = url; });
}
async function imageRun(url, maxW, maxH) {
  const { w, h } = await imgSize(url);
  const s = Math.min(maxW / w, maxH / h);
  return new docx.ImageRun({ data: dataUrlToBytes(url), type: "png", transformation: { width: Math.max(1, Math.round(w * s)), height: Math.max(1, Math.round(h * s)) } });
}

// Righe di testo dell'intestazione secondo le opzioni di layout
export function headerLines(p, L) {
  const lines = [];
  if (p.nome) lines.push({ t: p.nome, size: L.nameSize, bold: true });
  if (L.show.qualifica && p.qualifica) lines.push({ t: p.qualifica, size: L.subSize });
  const contacts = [L.show.indirizzo && p.indirizzo, L.show.telefono && p.telefono, L.show.email && p.email].filter(Boolean);
  if (contacts.length) {
    if (L.contactMode === "lines") contacts.forEach((c) => lines.push({ t: c, size: L.infoSize, gray: true }));
    else lines.push({ t: contacts.join("  ·  "), size: L.infoSize, gray: true });
  }
  if (L.show.extra && p.extra) lines.push({ t: p.extra, size: L.infoSize, gray: true });
  return lines;
}

const hex = (c) => (c || "#444444").replace("#", "").toUpperCase();

export async function buildDocx(profile, d, layout) {
  const D = docx, L = mergeLayout(layout), p = profile;
  const AL = { left: D.AlignmentType.LEFT, center: D.AlignmentType.CENTER, right: D.AlignmentType.RIGHT };
  const NONE = { style: D.BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const noBorders = { top: NONE, bottom: NONE, left: NONE, right: NONE };

  // --- intestazione ---
  const textParas = headerLines(p, L).map((l) => new D.Paragraph({
    alignment: AL[L.hAlign], spacing: { after: 0 },
    children: [new D.TextRun({ text: l.t, bold: !!l.bold, size: Math.round(l.size * 2), color: l.gray ? "555555" : undefined })],
  }));
  const hdr = [];
  const logoOn = L.logoPos !== "none" && p.logoData;
  if (logoOn && L.logoPos === "sopra") {
    hdr.push(new D.Paragraph({ alignment: AL[L.hAlign], spacing: { after: 60 }, children: [await imageRun(p.logoData, L.logoW, L.logoW)] }));
  }
  if (logoOn && (L.logoPos === "sinistra" || L.logoPos === "destra")) {
    const logoCell = new D.TableCell({ borders: noBorders, width: { size: 25, type: D.WidthType.PERCENTAGE }, verticalAlign: D.VerticalAlign.CENTER,
      children: [new D.Paragraph({ alignment: L.logoPos === "sinistra" ? D.AlignmentType.LEFT : D.AlignmentType.RIGHT, children: [await imageRun(p.logoData, L.logoW, L.logoW)] })] });
    const textCell = new D.TableCell({ borders: noBorders, width: { size: 75, type: D.WidthType.PERCENTAGE }, verticalAlign: D.VerticalAlign.CENTER,
      children: textParas.length ? textParas : [new D.Paragraph({ children: [] })] });
    hdr.push(new D.Table({ width: { size: 100, type: D.WidthType.PERCENTAGE }, borders: D.TableBorders.NONE,
      rows: [new D.TableRow({ children: L.logoPos === "sinistra" ? [logoCell, textCell] : [textCell, logoCell] })] }));
  } else {
    hdr.push(...textParas);
  }
  hdr.push(new D.Paragraph({
    spacing: { before: 60 },
    border: L.line ? { bottom: { style: D.BorderStyle.SINGLE, size: 8, space: 4, color: hex(L.lineColor) } } : undefined,
    children: [],
  }));

  // --- corpo ---
  const body = [];
  const title = (d.titolo || d.tipo || "").toUpperCase();
  if (title) body.push(new D.Paragraph({ alignment: L.titleAlign === "left" ? D.AlignmentType.LEFT : D.AlignmentType.CENTER, spacing: { before: 200, after: 240 }, children: [new D.TextRun({ text: title, bold: true, size: Math.round((L.bodySize + 3) * 2) })] }));
  body.push(new D.Paragraph({ spacing: { after: 60 }, children: [new D.TextRun({ text: "Paziente: ", bold: true }), new D.TextRun(d.paziente || "")] }));
  body.push(new D.Paragraph({ spacing: { after: 240 }, children: [new D.TextRun({ text: "Data: ", bold: true }), new D.TextRun(itDate(d.data))] }));
  for (const t of (d.testo || "").split("\n")) {
    body.push(new D.Paragraph({ spacing: { after: 100 }, alignment: L.justify ? D.AlignmentType.JUSTIFIED : D.AlignmentType.LEFT, children: [new D.TextRun({ text: t })] }));
  }

  // --- luogo, data, timbro e firma ---
  body.push(new D.Paragraph({ spacing: { before: 400 }, alignment: AL[L.sigAlign], children: [new D.TextRun({ text: [p.luogo, itDate(d.data)].filter(Boolean).join(", ") })] }));
  const hasF = !!p.firmaData, hasT = !!p.timbroData;
  if (hasF && hasT) {
    const cell = async (url, w, h, right) => new D.TableCell({ borders: noBorders, width: { size: 50, type: D.WidthType.PERCENTAGE },
      children: [new D.Paragraph({ alignment: right ? D.AlignmentType.RIGHT : D.AlignmentType.LEFT, children: [await imageRun(url, w, h)] })] });
    const stamp = await cell(p.timbroData, L.stampW, 90, L.stampSide === "right");
    const sign = await cell(p.firmaData, L.sigW, 80, L.stampSide !== "right");
    body.push(new D.Table({ width: { size: 100, type: D.WidthType.PERCENTAGE }, borders: D.TableBorders.NONE,
      rows: [new D.TableRow({ children: L.stampSide === "right" ? [sign, stamp] : [stamp, sign] })] }));
  } else if (hasF || hasT) {
    body.push(new D.Paragraph({ alignment: AL[L.sigAlign], children: [await imageRun(hasF ? p.firmaData : p.timbroData, hasF ? L.sigW : L.stampW, 90)] }));
  }
  if (p.nome) body.push(new D.Paragraph({ alignment: AL[L.sigAlign], children: [new D.TextRun({ text: p.nome })] }));

  const footer = new D.Footer({ children: [new D.Paragraph({ alignment: D.AlignmentType.CENTER, children: p.piede ? [new D.TextRun({ text: p.piede, size: 16, color: "666666" })] : [] })] });
  const [mt, mb, ml, mr] = MARGINS[L.margin] || MARGINS.normal;

  const doc = new D.Document({
    creator: p.nome || "Gestionale",
    styles: { default: { document: { run: { font: L.font, size: Math.round(L.bodySize * 2) } } } },
    sections: [{
      properties: { page: { margin: { top: mt, bottom: mb, left: ml, right: mr } } },
      headers: { default: new D.Header({ children: hdr }) },
      footers: { default: footer },
      children: body,
    }],
  });
  return D.Packer.toBlob(doc);
}

// Anteprima a schermo (HTML) coerente con il documento Word
export function previewHtml(profile, d, esc, layout) {
  const p = profile, L = mergeLayout(layout);
  const px = (pt) => (pt * 1.333).toFixed(1) + "px";
  const lines = headerLines(p, L).map((l) =>
    `<div style="font-size:${px(l.size)};${l.bold ? "font-weight:700;" : ""}${l.gray ? "color:#555;" : ""}">${esc(l.t)}</div>`).join("");
  const logo = p.logoData && L.logoPos !== "none" ? `<img src="${p.logoData}" alt="logo" style="width:${L.logoW}px;max-height:${L.logoW}px;object-fit:contain">` : "";
  let head;
  if (logo && L.logoPos === "sopra") head = `<div style="text-align:${L.hAlign}">${logo}${lines}</div>`;
  else if (logo) head = `<div style="display:flex;align-items:center;gap:12px;flex-direction:${L.logoPos === "destra" ? "row-reverse" : "row"}"><div>${logo}</div><div style="flex:1;text-align:${L.hAlign}">${lines}</div></div>`;
  else head = `<div style="text-align:${L.hAlign}">${lines}</div>`;
  const body = (d.testo || "").split("\n").map((t) => `<p style="margin:0 0 6px">${esc(t) || "&nbsp;"}</p>`).join("");
  const pad = { narrow: "14px 16px", normal: "28px 32px", wide: "40px 48px" }[L.margin];
  const sigRow = p.firmaData && p.timbroData
    ? `<div style="display:flex;justify-content:space-between;align-items:center;margin-top:6px;flex-direction:${L.stampSide === "right" ? "row-reverse" : "row"}"><img src="${p.timbroData}" style="max-width:${L.stampW}px;max-height:90px"><img src="${p.firmaData}" style="max-width:${L.sigW}px;max-height:80px"></div>`
    : p.firmaData || p.timbroData
      ? `<div style="text-align:${L.sigAlign};margin-top:6px"><img src="${p.firmaData || p.timbroData}" style="max-width:${p.firmaData ? L.sigW : L.stampW}px;max-height:90px"></div>` : "";
  return `<div class="sheet" style="font-family:'${L.font}',Calibri,sans-serif;font-size:${px(L.bodySize)};padding:${pad}">
    <div style="${L.line ? `border-bottom:2px solid ${L.lineColor};` : ""}padding-bottom:8px;margin-bottom:14px">${head || "&nbsp;"}</div>
    <h3 style="text-align:${L.titleAlign};font-size:${px(L.bodySize + 3)}">${esc((d.titolo || d.tipo || "").toUpperCase())}</h3>
    <p style="margin:0 0 10px"><b>Paziente:</b> ${esc(d.paziente || "")}<br><b>Data:</b> ${esc(itDate(d.data))}</p>
    <div style="text-align:${L.justify ? "justify" : "left"}">${body}</div>
    <p style="text-align:${L.sigAlign};margin:18px 0 0">${esc([p.luogo, itDate(d.data)].filter(Boolean).join(", "))}</p>
    ${sigRow}
    ${p.nome ? `<p style="text-align:${L.sigAlign};margin:6px 0 0">${esc(p.nome)}</p>` : ""}
    ${p.piede ? `<div style="font-size:11px;color:#666;text-align:center;margin-top:24px">${esc(p.piede)}</div>` : ""}
  </div>`;
}
