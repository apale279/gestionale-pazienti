import { CONFIG } from "./config.js";
import { initAuth, login, logout, getClientId, saveClientId } from "./auth.js";
import { GraphDrive, DemoDrive } from "./drive.js";
import { scan, buildPlan, applyPlan, sanitizeName, extOf, stripExt } from "./indexer.js";
import { buildDocx, docFileName, previewHtml, itDate, todayISO, EMPTY_PROFILE } from "./docgen.js";

// ---------- utilità ----------
const join = (...p) => p.filter(Boolean).join("/");
const norm = (s) => s.normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

function h(tag, attrs = {}, ...kids) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v == null || v === false) continue;
    if (k === "class") e.className = v;
    else if (k.startsWith("on")) e.addEventListener(k.slice(2), v);
    else if (k === "value" || k === "checked" || k === "selected" || k === "disabled") e[k] = v;
    else e.setAttribute(k, v === true ? "" : v);
  }
  for (const c of kids.flat()) { if (c == null || c === false) continue; e.append(c.nodeType ? c : document.createTextNode(c)); }
  return e;
}
const $ = (s) => document.querySelector(s);

function toast(msg, err = false) {
  const t = $("#toast");
  t.textContent = msg; t.className = "show" + (err ? " err" : "");
  clearTimeout(toast.t); toast.t = setTimeout(() => (t.className = ""), err ? 6000 : 2800);
}
function busy(msg) {
  const o = h("div", { class: "overlay" }, h("div", { class: "busybox" }, h("div", { class: "spinner" }), h("div", { class: "busymsg" }, msg)));
  document.body.append(o);
  const done = () => o.remove();
  done.set = (m) => (o.querySelector(".busymsg").textContent = m);
  return done;
}
function modal(title, content, buttons) {
  const m = h("div", { class: "overlay" }, h("div", { class: "modal" }, h("h3", {}, title), content, h("div", { class: "btnrow" }, buttons)));
  document.body.append(m);
  return () => m.remove();
}
const icon = (n) => ({ pdf: "📕", docx: "📝", doc: "📝", dotx: "📋", xlsx: "📊", jpeg: "🖼️", jpg: "🖼️", png: "🖼️", heic: "🖼️", zip: "🗜️" }[extOf(n)] || "📎");
const fmtSize = (b) => (b > 1048576 ? (b / 1048576).toFixed(1) + " MB" : Math.max(1, Math.round(b / 1024)) + " KB");

// ---------- stato ----------
const S = { drive: null, st: { profile: { ...EMPTY_PROFILE }, hiddenFolders: [] }, sc: null, q: "", account: "" };
const settingsPath = () => join(CONFIG.rootPath, CONFIG.appFolder, "impostazioni.json");
const scanCfg = () => ({ ...CONFIG, hiddenFolders: S.st.hiddenFolders });
const patientByName = (n) => S.sc && S.sc.patients.find((p) => p.name === n);

async function loadSettings() {
  try {
    const j = await S.drive.readJson(settingsPath());
    if (j) S.st = { profile: { ...EMPTY_PROFILE, ...(j.profile || {}) }, hiddenFolders: j.hiddenFolders || [] };
  } catch (e) { console.warn("impostazioni", e); }
}
const saveSettings = () => S.drive.writeJson(settingsPath(), S.st);

async function refresh(silent = false) {
  const done = silent ? null : busy("Leggo la cartella pazienti…");
  try {
    S.sc = await scan(S.drive, scanCfg(), (i, n) => done && done.set(`Leggo la cartella pazienti… ${i}/${n}`));
  } catch (e) {
    console.error(e);
    showFatal(e);
  } finally { done && done(); }
  render();
}

// ---------- shell ----------
function buildShell() {
  const root = $("#app");
  root.replaceChildren(
    h("header", { class: "top" },
      h("button", { class: "iconbtn back", title: "Indietro", onclick: () => (location.hash = "#/") }, "‹"),
      h("a", { class: "brand", href: "#/" }, "Pazienti"),
      h("span", { class: "spacer" }),
      h("button", { class: "iconbtn", title: "Aggiorna", onclick: () => refresh() }, "⟳"),
      h("a", { class: "iconbtn", title: "Impostazioni", href: "#/impostazioni" }, "⚙︎")),
    S.drive.isDemo ? h("div", { class: "demo" }, "Modalità demo: dati inventati, nulla viene salvato su OneDrive. ", h("a", { href: location.pathname }, "Esci dalla demo")) : null,
    h("div", { class: "cols" }, h("aside", { id: "list" }), h("main", { id: "detail" })));
}

function route() {
  const p = location.hash.slice(1).split("/").filter(Boolean).map(decodeURIComponent);
  return { name: p[0] || "home", arg: p[1] };
}

function render() {
  if (!S.sc) return;
  const r = route();
  document.body.dataset.view = r.name === "home" ? "home" : "detail";
  renderList();
  const d = $("#detail");
  d.replaceChildren();
  if (r.name === "p") d.append(viewPatient(r.arg));
  else if (r.name === "nuovo") d.append(viewNewDoc(r.arg));
  else if (r.name === "riordina") d.append(viewReorder());
  else if (r.name === "impostazioni") d.append(viewSettings());
  else if (r.name === "paziente") d.append(viewNewPatient());
  else d.append(viewHome());
  window.scrollTo(0, 0);
}
window.addEventListener("hashchange", render);

// ---------- elenco ----------
function renderList() {
  const list = $("#list");
  const r = route();
  if (!list.firstChild) {
    list.append(
      h("div", { class: "searchrow" },
        h("input", { id: "q", type: "search", placeholder: "Cerca paziente o documento…", autocomplete: "off", value: S.q, oninput: (e) => { S.q = e.target.value; renderItems(); } }),
        h("a", { class: "btn primary", href: "#/paziente" }, "+ Nuovo")),
      h("div", { id: "banner" }), h("div", { id: "items" }));
  }
  const nl = S.sc.loose.length;
  $("#banner").replaceChildren(nl ? h("a", { class: "banner", href: "#/riordina" }, `⚠︎ ${nl} file sciolti da sistemare ›`) : "");
  renderItems(r.name === "p" ? r.arg : null);
}
function renderItems(active) {
  if (active === undefined) { const r = route(); active = r.name === "p" ? r.arg : null; }
  const q = norm(S.q.trim());
  const ps = S.sc.patients.filter((p) => !q || norm(p.name).includes(q) || p.files.some((f) => norm(f.name).includes(q)));
  $("#items").replaceChildren(...[
    ...ps.map((p) => h("a", { class: "pitem" + (p.name === active ? " active" : ""), href: "#/p/" + encodeURIComponent(p.name) },
      h("div", { class: "pname" }, p.name),
      h("div", { class: "pmeta" }, `${p.files.length} ${p.files.length === 1 ? "documento" : "documenti"}${p.lastDate ? " · ultimo " + itDate(p.lastDate) : ""}`))),
    ps.length ? null : h("p", { class: "muted pad" }, q ? "Nessun risultato." : "Nessun paziente."),
  ].filter(Boolean));
}

// ---------- home (desktop, nessun paziente selezionato) ----------
function viewHome() {
  const n = S.sc.patients.reduce((a, p) => a + p.files.length, 0);
  return h("section", { class: "card homecard" },
    h("h2", {}, "Gestionale pazienti"),
    h("p", {}, `${S.sc.patients.length} pazienti · ${n} documenti indicizzati.`),
    h("p", { class: "muted" }, "Seleziona un paziente dall'elenco, oppure creane uno nuovo."),
    h("a", { class: "btn primary", href: "#/paziente" }, "+ Nuovo paziente"));
}

// ---------- scheda paziente ----------
async function openFile(path) {
  const w = window.open("", "_blank");
  try {
    const url = await S.drive.downloadUrl(path);
    if (w) w.location = url; else location.href = url;
  } catch (e) { if (w) w.close(); toast("Impossibile aprire il file: " + e.message, true); }
}
async function openFolderOnline(path) {
  const w = window.open("", "_blank");
  try {
    const it = await S.drive.stat(path);
    if (it && it.webUrl && w) w.location = it.webUrl; else { if (w) w.close(); toast("Link OneDrive non disponibile" + (S.drive.isDemo ? " in demo" : "")); }
  } catch (e) { if (w) w.close(); toast(e.message, true); }
}

function viewPatient(name) {
  const p = patientByName(name);
  if (!p) return h("section", { class: "card" }, h("p", {}, "Paziente non trovato."), h("a", { class: "btn", href: "#/" }, "Torna all'elenco"));
  const pick = (cap) => {
    const inp = h("input", { type: "file", multiple: true, accept: cap ? "image/*" : "", capture: cap ? "environment" : null, hidden: true,
      onchange: () => inp.files.length && attachFlow(p, [...inp.files]) });
    document.body.append(inp); inp.click(); setTimeout(() => inp.remove(), 60000);
  };
  const groups = new Map();
  for (const f of p.files) { if (!groups.has(f.sub)) groups.set(f.sub, []); groups.get(f.sub).push(f); }
  const keys = [...groups.keys()].sort((a, b) => (a === "" ? -1 : b === "" ? 1 : a.localeCompare(b, "it")));
  return h("section", {},
    h("div", { class: "card" },
      h("h2", {}, p.name),
      h("p", { class: "muted" }, `${p.files.length} documenti`),
      h("div", { class: "btnrow wrap" },
        h("a", { class: "btn primary", href: "#/nuovo/" + encodeURIComponent(p.name) }, "📝 Nuovo documento"),
        h("button", { class: "btn", onclick: () => pick(false) }, "📎 Allega file"),
        h("button", { class: "btn", onclick: () => pick(true) }, "📷 Scatta foto"),
        h("button", { class: "btn", onclick: () => openFolderOnline(p.path) }, "☁︎ Cartella su OneDrive"))),
    p.files.length ? keys.map((k) => h("div", { class: "card" },
      k ? h("h4", {}, "📁 " + k) : h("h4", {}, "Documenti"),
      h("ul", { class: "files" }, groups.get(k).map((f) =>
        h("li", {}, h("button", { class: "file", onclick: () => openFile(f.path) },
          h("span", { class: "ficon" }, icon(f.name)),
          h("span", { class: "fbody" }, h("span", { class: "fname" }, f.name), h("span", { class: "fmeta" }, [f.date ? itDate(f.date) : "", fmtSize(f.size)].filter(Boolean).join(" · "))))))))) :
      h("div", { class: "card" }, h("p", { class: "muted" }, "Cartella vuota. Crea un documento o allega un file.")),
    h("div", { class: "card subtle" },
      h("button", { class: "btn small", onclick: () => hideFolder(p) }, "Nascondi dall'elenco"),
      h("span", { class: "muted small" }, " Non cancella nulla: la cartella resta su OneDrive.")));
}

async function hideFolder(p) {
  S.st.hiddenFolders = [...new Set([...S.st.hiddenFolders, ...p.folders.map((f) => f.split("/").pop())])];
  try { await saveSettings(); } catch (e) { toast("Impostazioni non salvate: " + e.message, true); }
  location.hash = "#/"; await refresh(true);
}

// ---------- allegati ----------
function attachFlow(p, files) {
  let rename = true;
  const hasDate = (n) => /^\d{4}[_\-.]\d{2}[_\-.]\d{2}/.test(n);
  const target = (f) => {
    let n = f.name;
    if (/^(image|img|photo|foto)[_\-\s]*\d*\.\w+$/i.test(n)) n = "Foto" + n.slice(n.lastIndexOf("."));
    return rename && !hasDate(n) ? `${todayISO().replace(/-/g, "_")}_${n}` : n;
  };
  const listEl = h("ul", { class: "files" });
  const upd = () => listEl.replaceChildren(...files.map((f) => h("li", { class: "plain" }, icon(f.name) + " " + target(f), h("span", { class: "muted small" }, " · " + fmtSize(f.size)))));
  upd();
  const cb = h("input", { type: "checkbox", checked: true, onchange: (e) => { rename = e.target.checked; upd(); } });
  const close = modal(`Allega a ${p.name}`,
    h("div", {}, listEl, h("label", { class: "check" }, cb, " Aggiungi la data davanti al nome (AAAA_MM_GG_…)")),
    [h("button", { class: "btn", onclick: () => close() }, "Annulla"),
     h("button", { class: "btn primary", onclick: async () => {
        close();
        const done = busy("Carico i file…");
        try {
          let i = 0;
          for (const f of files) {
            done.set(`Carico ${++i}/${files.length}: ${f.name}`);
            await S.drive.upload(p.path, sanitizeName(target(f)).replace(/ (\.\w+)$/, "$1"), f, "rename", (x) => done.set(`Carico ${i}/${files.length}: ${Math.round(x * 100)}%`));
          }
          toast("File salvati nella cartella del paziente");
        } catch (e) { toast("Errore nel caricamento: " + e.message, true); }
        done(); await refresh(true);
      } }, "Allega")]);
}

// ---------- nuovo paziente ----------
function viewNewPatient() {
  const cog = h("input", { id: "cog", placeholder: "Cognome", autocapitalize: "words", autocomplete: "off" });
  const nom = h("input", { id: "nom", placeholder: "Nome", autocapitalize: "words", autocomplete: "off" });
  return h("section", { class: "card" },
    h("h2", {}, "Nuovo paziente"),
    h("p", { class: "muted" }, "Verrà creata la cartella “Cognome Nome” dentro la cartella Pazienti."),
    h("label", {}, "Cognome", cog), h("label", {}, "Nome", nom),
    h("div", { class: "btnrow" },
      h("a", { class: "btn", href: "#/" }, "Annulla"),
      h("button", { class: "btn primary", onclick: async () => {
        const name = sanitizeName(`${cog.value} ${nom.value}`);
        if (!name) return toast("Inserisci almeno cognome o nome", true);
        const dup = S.sc.patients.find((p) => norm(p.name) === norm(name));
        if (dup) { toast("Esiste già: apro la scheda"); location.hash = "#/p/" + encodeURIComponent(dup.name); return; }
        const done = busy("Creo la cartella…");
        try { await S.drive.mkdir(CONFIG.rootPath, name); } catch (e) { done(); return toast("Errore: " + e.message, true); }
        done(); await refresh(true);
        location.hash = "#/p/" + encodeURIComponent(name);
      } }, "Crea paziente")));
}

// ---------- nuovo documento ----------
const TEXTS = {
  Visita: "Motivo della visita:\n\nAnamnesi:\n\nEsame obiettivo:\n\nConclusioni e terapia:\n",
  Certificato: "Si certifica che il/la paziente sopra indicato/a è stato/a visitato/a in data odierna e presenta:\n\n\nIl presente certificato viene rilasciato su richiesta dell'interessato/a per gli usi consentiti dalla legge.",
  Referto: "Quesito clinico:\n\nEsame eseguito:\n\nReperti:\n\nConclusioni:\n",
  Relazione: "Egregio Collega,\n\nsi relaziona in merito al/alla paziente sopra indicato/a.\n\n\nCordiali saluti.",
  Anamnesi: "Anamnesi familiare:\n\nAnamnesi fisiologica:\n\nAnamnesi patologica remota:\n\nAnamnesi patologica prossima:\n",
  Ricetta: "Si prescrive:\n\n",
  Consenso: "Il/La sottoscritto/a, dopo aver ricevuto adeguata informazione, acconsenta al trattamento proposto e al trattamento dei dati personali ai sensi del Reg. UE 2016/679.",
  Lettera: "", Nota: "",
};

function viewNewDoc(name) {
  const p = patientByName(name);
  if (!p) return h("section", { class: "card" }, h("p", {}, "Paziente non trovato."));
  const d = { tipo: CONFIG.docTypes[0], data: todayISO(), titolo: "", testo: TEXTS[CONFIG.docTypes[0]] || "", paziente: p.name };
  let tpl = d.testo;
  const prev = h("div", { class: "preview" });
  const upd = () => { prev.innerHTML = previewHtml(S.st.profile, d, esc); };
  const tipo = h("select", { onchange: (e) => { d.tipo = e.target.value; if (d.testo === tpl || !d.testo.trim()) { tpl = TEXTS[d.tipo] || ""; d.testo = tpl; ta.value = tpl; } upd(); } },
    CONFIG.docTypes.map((t) => h("option", { value: t }, t)));
  const data = h("input", { type: "date", value: d.data, oninput: (e) => { d.data = e.target.value; upd(); } });
  const tit = h("input", { placeholder: "Titolo (se vuoto: tipo di documento)", oninput: (e) => { d.titolo = e.target.value; upd(); } });
  const ta = h("textarea", { rows: 14, value: d.testo, oninput: (e) => { d.testo = e.target.value; upd(); } });
  upd();
  const noProfile = !S.st.profile.nome;
  const make = async () => {
    const done = busy("Creo il documento…");
    try { const blob = await buildDocx(S.st.profile, d); done(); return blob; } catch (e) { done(); console.error(e); toast("Errore nella creazione: " + e.message, true); return null; }
  };
  return h("section", {},
    h("div", { class: "card" },
      h("h2", {}, "Nuovo documento"), h("p", { class: "muted" }, "Paziente: " + p.name),
      noProfile ? h("a", { class: "banner", href: "#/impostazioni" }, "⚠︎ Intestazione, firma e timbro non ancora impostati › ") : null,
      h("div", { class: "grid2" }, h("label", {}, "Tipo", tipo), h("label", {}, "Data", data)),
      h("label", {}, "Titolo", tit), h("label", {}, "Testo", ta),
      h("div", { class: "btnrow wrap" },
        h("a", { class: "btn", href: "#/p/" + encodeURIComponent(p.name) }, "Annulla"),
        h("button", { class: "btn", onclick: async () => { const b = await make(); if (b) saveBlob(b, docFileName(d)); } }, "⬇︎ Scarica"),
        h("button", { class: "btn primary", onclick: async () => {
          const b = await make(); if (!b) return;
          const done = busy("Salvo nella cartella del paziente…");
          try { await S.drive.upload(p.path, docFileName(d), b); toast("Documento salvato"); }
          catch (e) { done(); return toast("Errore nel salvataggio: " + e.message, true); }
          done(); await refresh(true); location.hash = "#/p/" + encodeURIComponent(p.name);
        } }, "Salva su OneDrive"))),
    h("div", { class: "card" }, h("h4", {}, "Anteprima"), prev));
}
function saveBlob(blob, name) {
  const a = h("a", { href: URL.createObjectURL(blob), download: name });
  document.body.append(a); a.click(); a.remove();
}

// ---------- riordino file sciolti ----------
function viewReorder() {
  const rows = buildPlan(S.sc);
  const names = S.sc.patients.map((p) => p.name);
  const rowEls = rows.map((r) => {
    const sel = h("select", {},
      h("option", { value: "skip", selected: r.kind === "skip" }, "— lascia dov'è —"),
      h("option", { value: "new", selected: r.kind === "new" }, "➕ Nuovo paziente…"),
      h("optgroup", { label: "Pazienti esistenti" }, names.map((n) => h("option", { value: "e:" + n, selected: r.kind === "existing" && r.target === n }, n))));
    const txt = h("input", { value: r.kind === "new" ? r.target : "", placeholder: "Cognome Nome" });
    const sync = () => { txt.style.display = sel.value === "new" ? "" : "none"; r.kind = sel.value === "new" ? "new" : sel.value === "skip" ? "skip" : "existing"; if (sel.value.startsWith("e:")) r.target = sel.value.slice(2); };
    sel.addEventListener("change", sync); txt.addEventListener("input", () => (r.target = txt.value)); sync();
    return h("div", { class: "rrow" },
      h("div", { class: "rfile" }, h("span", {}, icon(r.file.name) + " " + r.file.name), h("span", { class: "tag " + r.confidence }, r.confidence === "nessuna" ? "da assegnare" : "affid. " + r.confidence)),
      h("div", { class: "rtarget" }, sel, txt));
  });
  const apply = () => {
    const todo = rows.filter((r) => r.kind !== "skip" && r.target.trim());
    if (!todo.length) return toast("Niente da spostare");
    const newN = new Set(todo.filter((r) => r.kind === "new" && !S.sc.patients.some((p) => norm(p.name) === norm(sanitizeName(r.target)))).map((r) => norm(sanitizeName(r.target)))).size;
    const close = modal("Confermi?", h("p", {}, `Sposto ${todo.length} file nelle cartelle dei pazienti e creo ${newN} nuove cartelle. Nessun file viene cancellato o sovrascritto.`),
      [h("button", { class: "btn", onclick: () => close() }, "Annulla"),
       h("button", { class: "btn primary", onclick: async () => {
          close(); const done = busy("Riordino…");
          const log = await applyPlan(S.drive, scanCfg(), S.sc, todo, (i, n) => done.set(`Sposto ${i}/${n}…`));
          done();
          const bad = log.filter((l) => !l.ok);
          toast(bad.length ? `${log.length - bad.length} spostati, ${bad.length} errori: ${bad[0].name}: ${bad[0].error}` : `${log.length} file sistemati`, !!bad.length);
          await refresh(true);
        } }, "Sposta")]);
  };
  return h("section", {},
    h("div", { class: "card" },
      h("h2", {}, "Riordina file sciolti"),
      h("p", { class: "muted" }, "Ho abbinato i file ai pazienti dal nome. Controlla e correggi, poi premi “Applica”. Nulla viene spostato prima."),
      rows.length ? h("div", { class: "rlist" }, rowEls) : h("p", {}, "✅ Nessun file sciolto."),
      rows.length ? h("div", { class: "btnrow" }, h("button", { class: "btn primary", onclick: apply }, "Applica")) : null),
    S.sc.models.length ? h("div", { class: "card" }, h("h4", {}, "Modelli e carte intestate (restano dove sono)"),
      h("ul", { class: "files" }, S.sc.models.map((f) => h("li", {}, h("button", { class: "file", onclick: () => openFile(f.path) }, h("span", { class: "ficon" }, icon(f.name)), h("span", { class: "fbody" }, h("span", { class: "fname" }, f.name))))))) : null);
}

// ---------- impostazioni ----------
function toPngDataUrl(file, maxW, removeWhite) {
  return new Promise((res, rej) => {
    const img = new Image();
    img.onload = () => {
      const k = Math.min(1, maxW / img.naturalWidth);
      const c = document.createElement("canvas");
      c.width = Math.round(img.naturalWidth * k); c.height = Math.round(img.naturalHeight * k);
      const x = c.getContext("2d"); x.drawImage(img, 0, 0, c.width, c.height);
      if (removeWhite) {
        const id = x.getImageData(0, 0, c.width, c.height), d = id.data;
        for (let i = 0; i < d.length; i += 4) {
          const m = Math.min(d[i], d[i + 1], d[i + 2]);
          d[i + 3] = m >= 235 ? 0 : m > 190 ? Math.round(255 * (235 - m) / 45) : 255;
        }
        x.putImageData(id, 0, 0);
      }
      res(c.toDataURL("image/png"));
    };
    img.onerror = () => rej(new Error("Immagine non leggibile"));
    img.src = URL.createObjectURL(file);
  });
}

function viewSettings() {
  const P = { ...S.st.profile };
  const field = (k, label, ph, tag = "input") => h("label", {}, label, h(tag, { value: P[k], placeholder: ph || "", oninput: (e) => (P[k] = e.target.value) }));
  const imgBox = (k, label, maxW) => {
    const box = h("div", { class: "imgbox" });
    const show = () => box.replaceChildren(P[k] ? h("img", { src: P[k], alt: label }) : h("span", { class: "muted small" }, "nessuna immagine"));
    show();
    const inp = h("input", { type: "file", accept: "image/*", hidden: true, onchange: async () => {
      if (!inp.files[0]) return;
      try { P[k] = await toPngDataUrl(inp.files[0], maxW, true); show(); } catch (e) { toast(e.message, true); }
    } });
    return h("div", { class: "card inner" }, h("h4", {}, label), box,
      h("div", { class: "btnrow wrap" }, inp,
        h("button", { class: "btn small", onclick: () => inp.click() }, "Carica foto / immagine"),
        P[k] ? h("button", { class: "btn small", onclick: () => { P[k] = ""; show(); } }, "Rimuovi") : null),
      h("p", { class: "muted small" }, "Meglio una firma o un timbro su foglio bianco: lo sfondo bianco viene reso trasparente."));
  };
  const hiddenList = S.st.hiddenFolders.length ? h("ul", { class: "files" }, S.st.hiddenFolders.map((n) =>
    h("li", { class: "plain" }, n, " ", h("button", { class: "btn small", onclick: async () => { S.st.hiddenFolders = S.st.hiddenFolders.filter((x) => x !== n); await saveSettings(); await refresh(true); } }, "Ripristina")))) : h("p", { class: "muted small" }, "Nessuna.");
  return h("section", {},
    h("div", { class: "card" },
      h("h2", {}, "Intestazione dei documenti"),
      h("div", { class: "grid2" }, field("nome", "Nome e cognome", "Dott. Mario Rossi"), field("qualifica", "Qualifica", "Medico Chirurgo – Specialista in …")),
      field("indirizzo", "Indirizzo", "Via … , Città"),
      h("div", { class: "grid2" }, field("telefono", "Telefono"), field("email", "Email")),
      field("extra", "Albo / P.IVA / altro", "Iscr. Albo n. …  ·  P.IVA …"),
      h("div", { class: "grid2" }, field("luogo", "Luogo (accanto alla data)", "Milano"), field("piede", "Nota a piè di pagina")),
      imgBox("firmaData", "Firma", 500), imgBox("timbroData", "Timbro", 500),
      h("div", { class: "btnrow" }, h("button", { class: "btn primary", onclick: async () => {
        const done = busy("Salvo…");
        S.st.profile = P;
        try { await saveSettings(); toast("Impostazioni salvate"); } catch (e) { toast("Errore: " + e.message, true); }
        done();
      } }, "Salva"))),
    h("div", { class: "card" }, h("h4", {}, "Cartelle nascoste dall'elenco"), hiddenList),
    h("div", { class: "card" }, h("h4", {}, "Account"),
      h("p", { class: "muted small" }, S.drive.isDemo ? "Modalità demo" : `Collegato a OneDrive${S.account ? " come " + S.account : ""}. Cartella: ${CONFIG.rootPath}`),
      S.drive.isDemo ? null : h("button", { class: "btn", onclick: () => logout() }, "Esci")));
}

// ---------- avvio ----------
function showFatal(e) {
  $("#app").replaceChildren(h("div", { class: "center" }, h("div", { class: "card" },
    h("h2", {}, "Qualcosa non va"),
    h("p", {}, e && e.status === 404 ? `Non trovo la cartella “${CONFIG.rootPath}” nel tuo OneDrive. Controlla il percorso in config.js.` : String(e && e.message || e)),
    h("button", { class: "btn primary", onclick: () => location.reload() }, "Riprova"))));
}

function screen(...kids) { $("#app").replaceChildren(h("div", { class: "center" }, h("div", { class: "card login" }, ...kids))); }

async function startWith(drive) {
  S.drive = drive;
  buildShell();
  const done = busy("Avvio…");
  await loadSettings();
  done();
  await refresh();
}

async function boot() {
  const demo = new URLSearchParams(location.search).has("demo");
  if (demo) return startWith(new DemoDrive(CONFIG.rootPath));
  if ("serviceWorker" in navigator && (location.protocol === "https:" || location.hostname === "localhost")) navigator.serviceWorker.register("sw.js").catch(() => {});
  let a;
  try { a = await initAuth(); } catch (e) { return showFatal(e); }
  const demoBtn = h("a", { class: "btn", href: "?demo=1" }, "Prova la demo");
  if (a.state === "noclient") {
    const id = h("input", { placeholder: "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx", autocomplete: "off" });
    return screen(h("h2", {}, "Gestionale pazienti"),
      h("p", {}, "Prima di collegare OneDrive serve registrare l'app su Microsoft (una volta sola, gratis). Segui la GUIDA.md e incolla qui l'“ID applicazione (client)”."),
      h("label", {}, "ID applicazione", id),
      h("div", { class: "btnrow wrap" }, demoBtn, h("button", { class: "btn primary", onclick: () => { if (id.value.trim()) { saveClientId(id.value); location.reload(); } } }, "Continua")));
  }
  if (a.state === "loggedout") {
    return screen(h("h2", {}, "Gestionale pazienti"), h("p", {}, "Accedi con il tuo account Microsoft per collegare la cartella dei pazienti su OneDrive."),
      h("div", { class: "btnrow wrap" }, demoBtn, h("button", { class: "btn primary", onclick: () => login() }, "Accedi con Microsoft")));
  }
  S.account = a.account.username || "";
  return startWith(new GraphDrive());
}

boot();
