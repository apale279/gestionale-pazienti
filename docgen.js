// Generazione del documento Word con carta intestata, firma e timbro (tutto nel browser).
import { sanitizeName } from "./indexer.js";

export const EMPTY_PROFILE = {
  nome: "", qualifica: "", indirizzo: "", telefono: "", email: "", extra: "",
  luogo: "", piede: "", firmaData: "", timbroData: "",
};

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
  return new docx.ImageRun({ data: dataUrlToBytes(url), type: "png", transformation: { width: Math.round(w * s), height: Math.round(h * s) } });
}

export async function buildDocx(profile, d) {
  const D = docx;
  const hdr = [];
  const line = (text, o = {}) => new D.Paragraph({
    alignment: D.AlignmentType.CENTER,
    spacing: { after: o.after ?? 0 },
    border: o.border ? { bottom: { style: D.BorderStyle.SINGLE, size: 8, space: 6, color: "444444" } } : undefined,
    children: [new D.TextRun({ text, bold: !!o.bold, size: o.size || 20, color: o.color })],
  });
  if (profile.nome) hdr.push(line(profile.nome, { bold: true, size: 32 }));
  if (profile.qualifica) hdr.push(line(profile.qualifica, { size: 22 }));
  const contact = [profile.indirizzo, profile.telefono, profile.email].filter(Boolean).join("  ·  ");
  if (contact) hdr.push(line(contact, { size: 18, color: "555555" }));
  if (profile.extra) hdr.push(line(profile.extra, { size: 18, color: "555555" }));
  if (hdr.length) hdr.push(new D.Paragraph({ border: { bottom: { style: D.BorderStyle.SINGLE, size: 8, space: 4, color: "444444" } }, children: [] }));
  else hdr.push(new D.Paragraph({ children: [] }));

  const body = [];
  const title = (d.titolo || d.tipo || "").toUpperCase();
  if (title) body.push(new D.Paragraph({ alignment: D.AlignmentType.CENTER, spacing: { before: 200, after: 240 }, children: [new D.TextRun({ text: title, bold: true, size: 28 })] }));
  body.push(new D.Paragraph({ spacing: { after: 60 }, children: [new D.TextRun({ text: "Paziente: ", bold: true }), new D.TextRun(d.paziente || "")] }));
  body.push(new D.Paragraph({ spacing: { after: 240 }, children: [new D.TextRun({ text: "Data: ", bold: true }), new D.TextRun(itDate(d.data))] }));
  for (const t of (d.testo || "").split("\n")) {
    body.push(new D.Paragraph({ spacing: { after: 100 }, alignment: D.AlignmentType.JUSTIFIED, children: [new D.TextRun({ text: t, size: 22 })] }));
  }

  body.push(new D.Paragraph({ spacing: { before: 400 }, alignment: D.AlignmentType.RIGHT, children: [new D.TextRun({ text: [profile.luogo, itDate(d.data)].filter(Boolean).join(", "), size: 22 })] }));
  const cell = async (url, w, h, last) => new D.TableCell({
    borders: { top: { style: D.BorderStyle.NONE }, bottom: { style: D.BorderStyle.NONE }, left: { style: D.BorderStyle.NONE }, right: { style: D.BorderStyle.NONE } },
    width: { size: 50, type: D.WidthType.PERCENTAGE },
    children: [new D.Paragraph({ alignment: last ? D.AlignmentType.RIGHT : D.AlignmentType.LEFT, children: url ? [await imageRun(url, w, h)] : [] })],
  });
  if (profile.firmaData || profile.timbroData) {
    body.push(new D.Table({
      width: { size: 100, type: D.WidthType.PERCENTAGE },
      borders: D.TableBorders.NONE,
      rows: [new D.TableRow({ children: [await cell(profile.timbroData, 140, 90, false), await cell(profile.firmaData, 180, 80, true)] })],
    }));
  }
  if (profile.nome) body.push(new D.Paragraph({ alignment: D.AlignmentType.RIGHT, children: [new D.TextRun({ text: profile.nome, size: 22 })] }));

  const footer = profile.piede
    ? new D.Footer({ children: [new D.Paragraph({ alignment: D.AlignmentType.CENTER, children: [new D.TextRun({ text: profile.piede, size: 16, color: "666666" })] })] })
    : new D.Footer({ children: [new D.Paragraph({ children: [] })] });

  const doc = new D.Document({
    creator: profile.nome || "Gestionale",
    styles: { default: { document: { run: { font: "Calibri", size: 22 } } } },
    sections: [{
      properties: { page: { margin: { top: 1000, bottom: 1000, left: 1300, right: 1300 } } },
      headers: { default: new D.Header({ children: hdr }) },
      footers: { default: footer },
      children: body,
    }],
  });
  return D.Packer.toBlob(doc);
}

// Anteprima a schermo (HTML) coerente con il documento
export function previewHtml(profile, d, esc) {
  const p = profile;
  const contact = [p.indirizzo, p.telefono, p.email].filter(Boolean).map(esc).join(" · ");
  const body = (d.testo || "").split("\n").map((t) => `<p>${esc(t) || "&nbsp;"}</p>`).join("");
  return `<div class="sheet">
    <div class="lh">${p.nome ? `<div class="lh-name">${esc(p.nome)}</div>` : ""}
      ${p.qualifica ? `<div>${esc(p.qualifica)}</div>` : ""}
      ${contact ? `<div class="lh-small">${contact}</div>` : ""}
      ${p.extra ? `<div class="lh-small">${esc(p.extra)}</div>` : ""}</div>
    <h3>${esc((d.titolo || d.tipo || "").toUpperCase())}</h3>
    <p><b>Paziente:</b> ${esc(d.paziente || "")}<br><b>Data:</b> ${esc(itDate(d.data))}</p>
    ${body}
    <p class="sig-date">${esc([p.luogo, itDate(d.data)].filter(Boolean).join(", "))}</p>
    <div class="sig">${p.timbroData ? `<img src="${p.timbroData}" alt="timbro">` : "<span></span>"}${p.firmaData ? `<img src="${p.firmaData}" alt="firma">` : ""}</div>
    ${p.nome ? `<p class="sig-name">${esc(p.nome)}</p>` : ""}
    ${p.piede ? `<div class="lh-small foot">${esc(p.piede)}</div>` : ""}
  </div>`;
}
