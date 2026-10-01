import { CONFIG } from "./config.js";
import { initAuth, login, logout, getClientId, saveClientId } from "./auth.js";
import { GraphDrive, DemoDrive } from "./drive.js";
import { scan, buildPlan, applyPlan, sanitizeName, extOf, stripExt } from "./indexer.js";
import { DEFAULT_BMC, EMPTY_BMC_DOC, buildBmcDocx, bmcFileName, bmcHtml, printHtml } from "./bmcdoc.js";
import { buildDocx, docFileName, previewHtml, itDate, todayISO, EMPTY_PROFILE, DEFAULT_LAYOUT, FONTS, mergeLayout } from "./docgen.js";

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
const S = { drive: null, st: { profile: { ...EMPTY_PROFILE }, layout: mergeLayout(), hiddenFolders: [] }, bmc: { ...DEFAULT_BMC }, meta: {}, filter: "all", sc: null, q: "", account: "" };
const settingsPath = () => join(CONFIG.rootPath, CONFIG.appFolder, "impostazioni.json");
const scanCfg = () => ({ ...CONFIG, hiddenFolders: S.st.hiddenFolders });
const allPatients = () => (S.sc ? [...S.sc.patients, ...S.sc.bmc] : []);
const patientByKey = (k) => allPatients().find((p) => p.path === k);

async function loadSettings() {
  try {
    const j = await S.drive.readJson(settingsPath());
    if (j) S.st = { profile: { ...EMPTY_PROFILE, ...(j.profile || {}) }, layout: mergeLayout(j.layout), hiddenFolders: j.hiddenFolders || [] };
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
  root.replaceChildren(...[
    h("header", { class: "top" },
      h("button", { class: "iconbtn back", title: "Indietro", onclick: () => (location.hash = "#/") }, "‹"),
      h("a", { class: "brand", href: "#/" }, "Pazienti"),
      h("span", { class: "spacer" }),
      h("button", { class: "iconbtn", title: "Aggiorna", onclick: () => refresh() }, "⟳"),
      h("a", { class: "iconbtn", title: "Layout referti", href: "#/layout" }, "✎"),
      h("a", { class: "iconbtn", title: "Impostazioni", href: "#/impostazioni" }, "⚙︎")),
    S.drive.isDemo ? h("div", { class: "demo" }, "Modalità demo: dati inventati, nulla viene salvato su OneDrive. ", h("a", { href: location.pathname }, "Esci dalla demo")) : null,
    h("div", { class: "cols" }, h("aside", { id: "list" }), h("main", { id: "detail" })),
  ].filter(Boolean));
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
  else if (r.name === "layout") d.append(viewLayout());
  else if (r.name === "paziente") d.append(viewNewPatient(r.arg === "bmc"));
  else if (r.name === "bmc") d.append(viewBmcDoc(r.arg));
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
        h("a", { class: "btn primary", href: "#/paziente" }, "+ Nuovo"),
        h("a", { class: "btn", href: "#/paziente/bmc", title: "Nuovo paziente BMC" }, "+ BMC")),
      h("div", { id: "chips", class: "chips" }), h("div", { id: "banner" }), h("div", { id: "items" }));
  }
  const nl = S.sc.loose.length;
  $("#banner").replaceChildren(nl ? h("a", { class: "banner", href: "#/riordina" }, `⚠︎ ${nl} file sciolti da sistemare ›`) : "");
  renderItems(r.name === "p" ? r.arg : null);
}
function renderItems(active) {
  if (active === undefined) { const r = route(); active = r.name === "p" ? r.arg : null; }
  const q = norm(S.q.trim());
  $("#chips").replaceChildren(...[["all", "Tutti"], ["studio", "Studio"], ["bmc", "BMC"]].map(([v, t]) =>
    h("button", { class: "chip" + (S.filter === v ? " on" : ""), onclick: () => { S.filter = v; renderItems(); } }, t)));
  const pool = S.filter === "studio" ? S.sc.patients : S.filter === "bmc" ? S.sc.bmc : allPatients().sort((a, b) => a.name.localeCompare(b.name, "it"));
  const ps = pool.filter((p) => !q || norm(p.name).includes(q) || p.files.some((f) => norm(f.name).includes(q)));
  $("#items").replaceChildren(...[
    ...ps.map((p) => h("a", { class: "pitem" + (p.path === active ? " active" : ""), href: "#/p/" + encodeURIComponent(p.path) },
      h("div", { class: "pname" }, p.name, p.group === "BMC" ? h("span", { class: "badge" }, "BMC") : null),
      h("div", { class: "pmeta" }, `${p.files.length} ${p.files.length === 1 ? "documento" : "documenti"}${p.lastDate ? " · ultimo " + itDate(p.lastDate) : ""}`))),
    ps.length ? null : h("p", { class: "muted pad" }, q ? "Nessun risultato." : "Nessun paziente."),
  ].filter(Boolean));
}

// ---------- home (desktop, nessun paziente selezionato) ----------
function viewHome() {
  const n = allPatients().reduce((a, p) => a + p.files.length, 0);
  return h("section", { class: "card homecard" },
    h("h2", {}, "Gestionale pazienti"),
    h("p", {}, `${allPatients().length} pazienti (${S.sc.bmc.length} BMC) · ${n} documenti indicizzati.`),
    h("p", { class: "muted" }, "Seleziona un paziente dall'elenco, oppure creane uno nuovo."),
    h("div", { class: "btnrow wrap", style: "justify-content:center" }, h("a", { class: "btn primary", href: "#/paziente" }, "+ Nuovo paziente"), h("a", { class: "btn", href: "#/paziente/bmc" }, "+ Nuovo paziente BMC")));
}

// ---------- scheda paziente ----------
const VIEW_MIME = { pdf: "application/pdf", jpg: "image/jpeg", jpeg: "image/jpeg", png: "image/png", gif: "image/gif", webp: "image/webp", txt: "text/plain" };
// PDF e immagini: si aprono nel visualizzatore del browser. Word/Excel: nel visualizzatore online di OneDrive.
async function openFile(f) {
  const w = window.open("", "_blank");
  const go = (url) => { if (w) w.location = url; else location.href = url; };
  try {
    const type = VIEW_MIME[extOf(f.name)];
    if (type) go(URL.createObjectURL(new Blob([await S.drive.blob(f.path)], { type })));
    else if (f.webUrl) go(f.webUrl);
    else go(await S.drive.downloadUrl(f.path));
  } catch (e) { if (w) w.close(); toast("Impossibile aprire il file: " + e.message, true); }
}
// Apre il file nell'app desktop (Word/Excel) partendo dalla copia su OneDrive. Sperimentale.
function appLink(f) {
  const scheme = { docx: "ms-word", doc: "ms-word", dotx: "ms-word", xlsx: "ms-excel", xls: "ms-excel" }[extOf(f.name)];
  if (!scheme || !f.driveId || S.drive.isDemo) return null;
  return { label: scheme === "ms-word" ? "Word" : "Excel", href: `${scheme}:ofe|u|https://d.docs.live.net/${f.driveId.toLowerCase()}/${f.path.split("/").map(encodeURIComponent).join("/")}` };
}
function fileRow(f, withMeta = true) {
  const ap = appLink(f);
  return h("li", { class: "frow" },
    h("button", { class: "file", onclick: () => openFile(f) },
      h("span", { class: "ficon" }, icon(f.name)),
      h("span", { class: "fbody" }, h("span", { class: "fname" }, f.name), withMeta ? h("span", { class: "fmeta" }, [f.date ? itDate(f.date) : "", fmtSize(f.size)].filter(Boolean).join(" · ")) : null)),
    ap ? h("a", { class: "btn small appbtn", href: ap.href, title: "Apri nell'app " + ap.label + " (sperimentale)" }, ap.label) : null);
}
async function openFolderOnline(path) {
  const w = window.open("", "_blank");
  try {
    const it = await S.drive.stat(path);
    if (it && it.webUrl && w) w.location = it.webUrl; else { if (w) w.close(); toast("Link OneDrive non disponibile" + (S.drive.isDemo ? " in demo" : "")); }
  } catch (e) { if (w) w.close(); toast(e.message, true); }
}

function viewPatient(key) {
  const p = patientByKey(key);
  const pk = p && encodeURIComponent(p.path);
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
      h("h2", {}, p.name, p.group === "BMC" ? h("span", { class: "badge" }, "BMC") : null),
      h("p", { class: "muted" }, `${p.files.length} documenti`),
      h("div", { class: "btnrow wrap" },
        p.group === "BMC" ? h("a", { class: "btn primary", href: "#/bmc/" + pk }, "📋 Nuovo referto BMC") : null,
        h("a", { class: p.group === "BMC" ? "btn" : "btn primary", href: "#/nuovo/" + pk }, "📝 Nuovo documento"),
        h("button", { class: "btn", onclick: () => pick(false) }, "📎 Allega file"),
        h("button", { class: "btn", onclick: () => pick(true) }, "📷 Scatta foto"),
        h("button", { class: "btn", onclick: () => openFolderOnline(p.path) }, "☁︎ Cartella su OneDrive"))),
    cardAnagrafica(p), cardNote(p),
    p.files.length ? keys.map((k) => h("div", { class: "card" },
      k ? h("h4", {}, "📁 " + k) : h("h4", {}, "Documenti"),
      h("ul", { class: "files" }, groups.get(k).map((f) => fileRow(f))))) :
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
function viewNewPatient(isBmc) {
  const root = isBmc ? CONFIG.bmcPath : CONFIG.rootPath;
  const pool = isBmc ? S.sc.bmc : S.sc.patients;
  const cog = h("input", { id: "cog", placeholder: "Cognome", autocapitalize: "words", autocomplete: "off" });
  const nom = h("input", { id: "nom", placeholder: "Nome", autocapitalize: "words", autocomplete: "off" });
  return h("section", { class: "card" },
    h("h2", {}, isBmc ? "Nuovo paziente BMC" : "Nuovo paziente"),
    h("p", { class: "muted" }, isBmc ? "Verrà creata la cartella “Cognome Nome” dentro la cartella BMC." : "Verrà creata la cartella “Cognome Nome” dentro la cartella Pazienti."),
    h("label", {}, "Cognome", cog), h("label", {}, "Nome", nom),
    h("div", { class: "btnrow" },
      h("a", { class: "btn", href: "#/" }, "Annulla"),
      h("button", { class: "btn primary", onclick: async () => {
        const name = sanitizeName(`${cog.value} ${nom.value}`);
        if (!name) return toast("Inserisci almeno cognome o nome", true);
        const dup = pool.find((p) => norm(p.name) === norm(name));
        if (dup) { toast("Esiste già: apro la scheda"); location.hash = "#/p/" + encodeURIComponent(dup.path); return; }
        const done = busy("Creo la cartella…");
        try { await S.drive.ensureFolder(root); await S.drive.mkdir(root, name); } catch (e) { done(); return toast("Errore: " + e.message, true); }
        done(); await refresh(true);
        location.hash = "#/p/" + encodeURIComponent(join(root, name));
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

function viewNewDoc(key) {
  const p = patientByKey(key);
  if (!p) return h("section", { class: "card" }, h("p", {}, "Paziente non trovato."));
  const d = { tipo: CONFIG.docTypes[0], data: todayISO(), titolo: "", testo: TEXTS[CONFIG.docTypes[0]] || "", paziente: p.name, extraLines: [] };
  let tpl = d.testo;
  const prev = h("div", { class: "preview" });
  const upd = () => { prev.innerHTML = previewHtml(S.st.profile, d, esc, S.st.layout); };
  const tipo = h("select", { onchange: (e) => { d.tipo = e.target.value; if (d.testo === tpl || !d.testo.trim()) { tpl = TEXTS[d.tipo] || ""; d.testo = tpl; ta.value = tpl; } upd(); } },
    CONFIG.docTypes.map((t) => h("option", { value: t }, t)));
  const data = h("input", { type: "date", value: d.data, oninput: (e) => { d.data = e.target.value; upd(); } });
  const tit = h("input", { placeholder: "Titolo (se vuoto: tipo di documento)", oninput: (e) => { d.titolo = e.target.value; upd(); } });
  const ta = h("textarea", { rows: 14, value: d.testo, oninput: (e) => { d.testo = e.target.value; upd(); } });
  upd();
  // Dati anagrafici del paziente: se inseriti, si possono riportare nel documento
  const incl = { nascita: true, luogoNascita: true, cf: true };
  let anag = {};
  const anagBox = h("div", { class: "checks" });
  const syncLines = () => { d.extraLines = ANAG_DOC.filter(([k]) => incl[k] && anag[k]).map(([k, l]) => [l, k === "nascita" ? itDate(anag[k]) : anag[k]]); upd(); };
  loadMeta(p).then((meta) => {
    anag = meta.anagrafica; d.paziente = patientFullName(p, meta);
    const avail = ANAG_DOC.filter(([k]) => anag[k]);
    anagBox.replaceChildren(...(avail.length ? [h("div", { class: "muted small full" }, "Dati del paziente da riportare nel documento:"), ...avail.map(([k, l]) =>
      h("label", { class: "check" }, h("input", { type: "checkbox", checked: !!incl[k], onchange: (e) => { incl[k] = e.target.checked; syncLines(); } }), " " + l))] : []));
    syncLines();
  });
  const noProfile = !S.st.profile.nome;
  const make = async () => {
    const done = busy("Creo il documento…");
    try { const blob = await buildDocx(S.st.profile, d, S.st.layout); done(); return blob; } catch (e) { done(); console.error(e); toast("Errore nella creazione: " + e.message, true); return null; }
  };
  return h("section", {},
    h("div", { class: "card" },
      h("h2", {}, "Nuovo documento"), h("p", { class: "muted" }, "Paziente: " + p.name),
      noProfile ? h("a", { class: "banner", href: "#/layout" }, "⚠︎ Intestazione, firma e timbro non ancora impostati › ") : h("a", { class: "small", href: "#/layout" }, "✎ Modifica layout"),
      h("div", { class: "grid2" }, h("label", {}, "Tipo", tipo), h("label", {}, "Data", data)),
      h("label", {}, "Titolo", tit), anagBox, h("label", {}, "Testo", ta),
      h("div", { class: "btnrow wrap" },
        h("a", { class: "btn", href: "#/p/" + encodeURIComponent(p.path) }, "Annulla"),
        h("button", { class: "btn", onclick: async () => { const b = await make(); if (b) saveBlob(b, docFileName(d)); } }, "⬇︎ Scarica"),
        h("button", { class: "btn primary", onclick: async () => {
          const b = await make(); if (!b) return;
          const done = busy("Salvo nella cartella del paziente…");
          try { await S.drive.upload(p.path, docFileName(d), b); toast("Documento salvato"); }
          catch (e) { done(); return toast("Errore nel salvataggio: " + e.message, true); }
          done(); await refresh(true); location.hash = "#/p/" + encodeURIComponent(p.path);
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
      h("ul", { class: "files" }, S.sc.models.map((f) => fileRow(f, false)))) : null);
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
          if (d[i + 3] === 0) continue; // già trasparente: resta trasparente
          const m = Math.min(d[i], d[i + 1], d[i + 2]);
          const k = m >= 235 ? 0 : m > 190 ? (235 - m) / 45 : 1; // bianco → trasparente, sfumatura sui bordi
          d[i + 3] = Math.round(d[i + 3] * k);
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
  const hiddenList = S.st.hiddenFolders.length ? h("ul", { class: "files" }, S.st.hiddenFolders.map((n) =>
    h("li", { class: "plain" }, n, " ", h("button", { class: "btn small", onclick: async () => { S.st.hiddenFolders = S.st.hiddenFolders.filter((x) => x !== n); await saveSettings(); await refresh(true); } }, "Ripristina")))) : h("p", { class: "muted small" }, "Nessuna.");
  return h("section", {},
    h("div", { class: "card" }, h("h2", {}, "Impostazioni"),
      h("p", { class: "muted" }, "Intestazione, firma, timbro e aspetto dei documenti si modificano nell'editor di layout."),
      h("a", { class: "btn primary", href: "#/layout" }, "✎ Layout referti")),
    cardBmcSettings(),
    h("div", { class: "card" }, h("h4", {}, "Cartelle nascoste dall'elenco"), hiddenList),
    h("div", { class: "card" }, h("h4", {}, "Account"),
      h("p", { class: "muted small" }, S.drive.isDemo ? "Modalità demo" : `Collegato a OneDrive${S.account ? " come " + S.account : ""}. Cartella: ${CONFIG.rootPath}`),
      S.drive.isDemo ? null : h("button", { class: "btn", onclick: () => logout() }, "Esci")));
}

// ---------- editor di layout ----------
function viewLayout() {
  const P = { ...S.st.profile };
  const L = mergeLayout(S.st.layout);
  const sample = { tipo: "Visita", data: todayISO(), titolo: "", paziente: "Mario Rossi (esempio)",
    testo: "Motivo della visita:\nControllo periodico.\n\nEsame obiettivo:\nNella norma.\n\nConclusioni e terapia:\nSi consiglia controllo tra sei mesi." };
  const prev = h("div", { class: "preview" });
  const upd = () => { prev.innerHTML = previewHtml(P, sample, esc, L); };

  const text = (k, label, ph) => h("label", {}, label, h("input", { value: P[k], placeholder: ph || "", oninput: (e) => { P[k] = e.target.value; upd(); } }));
  const sel = (key, label, opts, num) => h("label", {}, label,
    h("select", { onchange: (e) => { L[key] = num ? Number(e.target.value) : e.target.value; upd(); } },
      opts.map(([v, t]) => h("option", { value: v, selected: String(L[key]) === String(v) }, t))));
  const range = (key, label, min, max, unit = "") => {
    const out = h("span", { class: "rv" }, L[key] + unit);
    return h("label", {}, h("span", { class: "rl" }, label, out),
      h("input", { type: "range", min, max, step: 1, value: L[key], oninput: (e) => { L[key] = Number(e.target.value); out.textContent = L[key] + unit; upd(); } }));
  };
  const check = (label, get, set) => h("label", { class: "check" }, h("input", { type: "checkbox", checked: get(), onchange: (e) => { set(e.target.checked); upd(); } }), label);

  const imgBox = (k, label, hint) => {
    const box = h("div", { class: "imgbox" });
    const btns = h("div", { class: "btnrow wrap" });
    const show = () => {
      box.replaceChildren(P[k] ? h("img", { src: P[k], alt: label }) : h("span", { class: "muted small" }, "nessuna immagine"));
      btns.replaceChildren(inp, h("button", { class: "btn small", onclick: () => inp.click() }, P[k] ? "Sostituisci" : "Carica foto / immagine"),
        P[k] ? h("button", { class: "btn small", onclick: () => { P[k] = ""; show(); upd(); } }, "Rimuovi") : "");
    };
    const inp = h("input", { type: "file", accept: "image/*", hidden: true, onchange: async () => {
      if (!inp.files[0]) return;
      try { P[k] = await toPngDataUrl(inp.files[0], 600, k !== "logoData"); show(); upd(); } catch (e) { toast(e.message, true); }
    } });
    show();
    return h("div", { class: "card inner" }, h("h4", {}, label), box, btns, h("p", { class: "muted small" }, hint));
  };

  const save = async () => {
    const done = busy("Salvo…");
    S.st.profile = { ...P }; S.st.layout = mergeLayout(L);
    try { await saveSettings(); toast("Layout salvato"); } catch (e) { toast("Errore: " + e.message, true); }
    done();
  };
  const reset = () => { S.st.layout = mergeLayout(); S.st.profile = { ...P }; render(); toast("Layout ripristinato (non ancora salvato)"); };
  const sampleDoc = async () => {
    const done = busy("Creo l'esempio…");
    try { saveBlob(await buildDocx(P, sample, L), "Esempio layout.docx"); } catch (e) { console.error(e); toast("Errore: " + e.message, true); }
    done();
  };

  const controls = h("div", { class: "designer-controls" },
    h("div", { class: "card" }, h("h3", {}, "I tuoi dati (intestazione)"),
      h("div", { class: "grid2" }, text("nome", "Nome e cognome", "Dott. Mario Rossi"), text("qualifica", "Qualifica", "Medico Chirurgo – Specialista in …")),
      text("indirizzo", "Indirizzo", "Via … , Città"),
      h("div", { class: "grid2" }, text("telefono", "Telefono"), text("email", "Email")),
      text("extra", "Albo / P.IVA / altro", "Iscr. Albo n. …  ·  P.IVA …"),
      h("div", { class: "grid2" }, text("luogo", "Luogo (accanto alla data)", "Milano"), text("piede", "Nota a piè di pagina"))),
    h("div", { class: "card" }, h("h3", {}, "Logo, firma e timbro"),
      imgBox("logoData", "Logo (facoltativo)", "Immagine con il tuo logo; resta com'è, senza togliere lo sfondo."),
      imgBox("firmaData", "Firma (o firma + timbro in un'unica immagine)", "Se hai firma e timbro nello stesso PNG, caricalo qui e lascia vuoto il timbro. Le parti già trasparenti restano trasparenti; lo sfondo bianco viene reso trasparente."),
      imgBox("timbroData", "Timbro (solo se separato dalla firma)", "Meglio su foglio bianco: lo sfondo bianco diventa trasparente.")),
    h("div", { class: "card" }, h("h3", {}, "Intestazione"),
      h("div", { class: "grid2" },
        sel("hAlign", "Posizione del testo", [["left", "In alto a sinistra"], ["center", "Al centro"], ["right", "In alto a destra"]]),
        sel("logoPos", "Logo", [["none", "Nessuno"], ["sopra", "Sopra il testo"], ["sinistra", "A sinistra"], ["destra", "A destra"]])),
      range("logoW", "Dimensione logo", 40, 200, " px"),
      sel("contactMode", "Recapiti", [["inline", "Su una sola riga"], ["lines", "Una riga ciascuno"]]),
      h("div", { class: "checks" },
        check("Qualifica", () => L.show.qualifica, (v) => (L.show.qualifica = v)),
        check("Indirizzo", () => L.show.indirizzo, (v) => (L.show.indirizzo = v)),
        check("Telefono", () => L.show.telefono, (v) => (L.show.telefono = v)),
        check("Email", () => L.show.email, (v) => (L.show.email = v)),
        check("Albo / P.IVA", () => L.show.extra, (v) => (L.show.extra = v)),
        check("Linea sotto l'intestazione", () => L.line, (v) => (L.line = v))),
      h("label", {}, "Colore della linea", h("input", { type: "color", value: L.lineColor, oninput: (e) => { L.lineColor = e.target.value; upd(); } })),
      range("nameSize", "Nome (pt)", 10, 28), range("subSize", "Qualifica (pt)", 8, 18), range("infoSize", "Recapiti (pt)", 7, 14)),
    h("div", { class: "card" }, h("h3", {}, "Pagina e testo"),
      h("div", { class: "grid2" }, sel("font", "Carattere", FONTS.map((f) => [f, f])), sel("margin", "Margini", [["narrow", "Stretti"], ["normal", "Normali"], ["wide", "Ampi"]])),
      range("bodySize", "Testo (pt)", 9, 14),
      h("div", { class: "grid2" }, sel("titleAlign", "Titolo", [["center", "Centrato"], ["left", "A sinistra"]]),
        h("div", {}, check("Testo giustificato", () => L.justify, (v) => (L.justify = v))))),
    h("div", { class: "card" }, h("h3", {}, "Firma e timbro"),
      h("div", { class: "grid2" },
        sel("sigAlign", "Luogo, data e firma", [["right", "A destra"], ["center", "Al centro"], ["left", "A sinistra"]]),
        sel("stampSide", "Con firma e timbro insieme", [["left", "Timbro a sinistra, firma a destra"], ["right", "Firma a sinistra, timbro a destra"]])),
      range("sigW", "Larghezza firma (o firma + timbro)", 80, 500, " px"), range("stampW", "Larghezza timbro", 80, 300, " px")));

  upd();
  return h("section", {},
    h("div", { class: "card" }, h("h2", {}, "Layout referti"),
      h("p", { class: "muted" }, "Imposta i tuoi dati e l'aspetto: l'anteprima si aggiorna mentre modifichi. Il layout vale per tutti i nuovi documenti."),
      h("div", { class: "btnrow wrap" },
        h("button", { class: "btn primary", onclick: save }, "Salva layout"),
        h("button", { class: "btn", onclick: sampleDoc }, "⬇︎ Scarica esempio Word"),
        h("button", { class: "btn", onclick: reset }, "Ripristina predefiniti"))),
    h("div", { class: "designer" }, controls, h("div", { class: "designer-preview" }, h("div", { class: "card" }, h("h4", {}, "Anteprima"), prev))));
}

// ---------- anagrafica e note private ----------
// Anagrafica: nel file nascosto ".paziente.json" dentro la cartella del paziente (segue la cartella se la rinomini o sposti).
// Note: in un file separato in _Gestionale/note/, MAI dentro la cartella del paziente e mai inserite nei documenti.
const ANAG_FIELDS = [
  ["cognome", "Cognome"], ["nome", "Nome"], ["nascita", "Data di nascita", "date"], ["luogoNascita", "Luogo di nascita"],
  ["sesso", "Sesso"], ["cf", "Codice fiscale"], ["indirizzo", "Residenza"], ["telefono", "Telefono"], ["email", "Email"],
];
const ANAG_DOC = [["nascita", "Data di nascita"], ["luogoNascita", "Luogo di nascita"], ["cf", "Codice fiscale"], ["indirizzo", "Residenza"], ["telefono", "Telefono"], ["email", "Email"]];
const metaPath = (p) => join(p.path, ".paziente.json");
const notePath = (id) => join(CONFIG.rootPath, CONFIG.appFolder, "note", id + ".json");
const bmcPath = () => join(CONFIG.rootPath, CONFIG.appFolder, "bmc.json");

// Una sola copia condivisa per paziente (scheda anagrafica e note la usano insieme).
function loadMeta(p) {
  if (!S.meta[p.path]) {
    S.meta[p.path] = (async () => {
      const meta = { id: "", anagrafica: {}, readError: "" };
      try {
        const m = await S.drive.readJson(metaPath(p));
        if (m) { meta.id = m.id || ""; meta.anagrafica = { ...(m.anagrafica || {}) }; }
      } catch (e) { console.warn("anagrafica", e); meta.readError = e.message; delete S.meta[p.path]; }
      return meta;
    })();
  }
  return S.meta[p.path];
}
async function saveMeta(p, meta) {
  // Se la lettura dei dati esistenti è fallita, NON si salva: si rischierebbe di sovrascriverli con un file vuoto.
  if (meta.readError) throw new Error("Non riesco a leggere i dati esistenti (" + meta.readError + "). Aggiorna la pagina e riprova.");
  if (!meta.id) meta.id = (crypto.randomUUID ? crypto.randomUUID() : String(Date.now()) + Math.random().toString(16).slice(2));
  await S.drive.writeJson(metaPath(p), { id: meta.id, anagrafica: meta.anagrafica });
}
const patientFullName = (p, meta) => {
  const a = (meta && meta.anagrafica) || {};
  return [a.cognome, a.nome].filter(Boolean).join(" ").trim() || p.name;
};
const age = (iso) => {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || ""); if (!m) return "";
  const n = new Date(); let a = n.getFullYear() - +m[1];
  if (n.getMonth() + 1 < +m[2] || (n.getMonth() + 1 === +m[2] && n.getDate() < +m[3])) a--;
  return a >= 0 && a < 130 ? `${a} anni` : "";
};

function cardAnagrafica(p) {
  const box = h("div", { class: "card" }, h("h4", {}, "Dati anagrafici"), h("div", { class: "spinner" }));
  const show = (meta) => {
    const a = meta.anagrafica;
    const has = ANAG_FIELDS.some(([k]) => a[k]);
    const rows = ANAG_FIELDS.filter(([k]) => a[k]).map(([k, l, t]) => h("div", { class: "kv" }, h("span", { class: "k" }, l), h("span", {}, t === "date" ? itDate(a[k]) + (age(a[k]) ? ` (${age(a[k])})` : "") : a[k])));
    box.replaceChildren(h("h4", {}, "Dati anagrafici"),
      has ? h("div", { class: "kvs" }, rows) : h("p", { class: "muted small" }, "Nessun dato inserito. Se li inserisci compariranno nei documenti che scrivi per questo paziente."),
      h("button", { class: "btn small", onclick: () => edit(meta) }, has ? "Modifica" : "Inserisci dati"));
  };
  const edit = (meta) => {
    const a = { ...meta.anagrafica };
    const inputs = ANAG_FIELDS.map(([k, l, t]) => k === "sesso"
      ? h("label", {}, l, h("select", { onchange: (e) => (a[k] = e.target.value) }, [["", "—"], ["M", "Maschio"], ["F", "Femmina"]].map(([v, tx]) => h("option", { value: v, selected: a[k] === v }, tx))))
      : h("label", {}, l, h("input", { type: t || "text", value: a[k] || "", autocomplete: "off", oninput: (e) => (a[k] = e.target.value) })));
    box.replaceChildren(h("h4", {}, "Dati anagrafici"), h("div", { class: "grid2" }, inputs),
      h("div", { class: "btnrow" },
        h("button", { class: "btn", onclick: () => show(meta) }, "Annulla"),
        h("button", { class: "btn primary", onclick: async () => {
          const done = busy("Salvo…");
          try { meta.anagrafica = a; await saveMeta(p, meta); toast("Dati salvati"); } catch (e) { toast("Errore: " + e.message, true); }
          done(); show(meta);
        } }, "Salva")));
  };
  loadMeta(p).then(show);
  return box;
}

function cardNote(p) {
  const box = h("div", { class: "card note" }, h("h4", {}, "🔒 Note private"), h("div", { class: "spinner" }));
  loadMeta(p).then(async (meta) => {
    let text = "", saved = "";
    if (meta.id) { try { const j = await S.drive.readJson(notePath(meta.id)); text = (j && j.text) || ""; } catch (e) { console.warn("note", e); } }
    saved = text;
    const st = h("span", { class: "muted small" }, "");
    const ta = h("textarea", { rows: 6, placeholder: "Appunti riservati su questo paziente…", value: text, oninput: (e) => { text = e.target.value; st.textContent = text === saved ? "" : "modifiche non salvate"; } });
    box.replaceChildren(h("h4", {}, "🔒 Note private"), ta,
      h("p", { class: "muted small" }, "Solo per te: non compaiono mai nei documenti e non stanno nella cartella del paziente (sono in “_Gestionale/note”)."),
      h("div", { class: "btnrow wrap" }, h("button", { class: "btn small primary", onclick: async () => {
        try {
          if (!meta.id) await saveMeta(p, meta);
          await S.drive.writeJson(notePath(meta.id), { text, updated: new Date().toISOString() });
          saved = text; st.textContent = "salvate ✓";
        } catch (e) { toast("Errore: " + e.message, true); }
      } }, "Salva note"), st));
  });
  return box;
}

// ---------- referto BMC ----------
const BMC_FIELDS = [
  ["numero", "Numero incarico", "text"], ["paziente", "Paz.", "text"], ["nascita", "Data di nascita", "date"],
  ["dataVisita", "Data visita", "date"], ["luogoVisita", "Luogo visita", "text"], ["medico", "Medico incaricato", "text"],
];
function viewBmcDoc(key) {
  const p = patientByKey(key);
  if (!p) return h("section", { class: "card" }, h("p", {}, "Paziente non trovato."));
  const d = { ...EMPTY_BMC_DOC(), paziente: p.name, consensoNome: p.name, medico: S.bmc.medico || S.st.profile.nome || "", luogoVisita: S.bmc.luogo || "", conFirma: !!(S.st.profile.firmaData || S.st.profile.timbroData) };
  const prev = h("div", { class: "preview" });
  const upd = () => { prev.innerHTML = bmcHtml(d, S.bmc, S.st.profile, esc); };
  const el = {};
  const inp = (k, label, type = "text") => (el[k] = h("input", { type, value: d[k], autocomplete: "off", oninput: (e) => { d[k] = e.target.value; if (k === "paziente" && d.consensoNome === prevName) { d.consensoNome = d.paziente; el.consensoNome.value = d.paziente; } if (k === "paziente") prevName = d.paziente; upd(); } }));
  let prevName = d.paziente;
  const area = (k, rows = 3) => (el[k] = h("textarea", { rows, value: d[k], oninput: (e) => { d[k] = e.target.value; upd(); } }));
  const small = (k, ph) => (el[k] = h("input", { value: d[k], placeholder: ph, inputmode: "decimal", oninput: (e) => { d[k] = e.target.value; upd(); } }));
  const L = (t, node) => h("label", {}, t, node);

  const form = h("div", { class: "card" },
    h("h2", {}, "Medical report BMCh24"), h("p", { class: "muted" }, "Paziente: " + p.name),
    h("div", { class: "grid2" }, L("Numero incarico", inp("numero")), L("Paz.", inp("paziente"))),
    h("div", { class: "grid2" }, L("Data di nascita", inp("nascita", "", "date")), L("Medico incaricato", inp("medico"))),
    h("div", { class: "grid2" }, L("Data visita", inp("dataVisita", "", "date")), L("Luogo visita", inp("luogoVisita"))),
    L("Motivo della chiamata", area("motivo", 3)), L("Anamnesi/Allergie", area("anamnesi", 5)),
    h("h4", {}, "Parametri vitali all'ingresso"),
    h("div", { class: "grid4" }, L("P.A. max", small("pa1", "120")), L("P.A. min (mmHg)", small("pa2", "80")), L("F.C. (/min)", small("fc", "72")), L("T° (°C)", small("temp", "36.5")), L("SpO2 (%)", small("spo2", "98"))),
    h("h4", {}, "Esame obiettivo"),
    L("Torace", area("torace", 2)), L("Addome", area("addome", 2)), L("Arti inferiori", area("arti", 2)), L("Altro", area("altro", 2)),
    L("DIAGNOSI", area("diagnosi", 4)), L("PROGNOSI", area("prognosi", 3)), L("Conclusione / TERAPIA", area("terapia", 6)),
    L("Note aggiuntive (facoltativo, es. traduzione)", area("extra", 4)),
    h("h4", {}, "Consenso e firme"),
    h("div", { class: "grid2" }, L("Paziente (cognome e nome) nel consenso", inp("consensoNome")), L("Firma del paziente — precisazione (facoltativa, es. familiare)", inp("firmaNote"))),
    h("label", { class: "check" }, h("input", { type: "checkbox", checked: d.conFirma, onchange: (e) => { d.conFirma = e.target.checked; upd(); } }), " Inserisci la mia firma e il mio timbro"),
    S.bmc.logoData ? null : h("a", { class: "banner", href: "#/impostazioni" }, "⚠︎ Logo BMCh24 non impostato › "));

  const make = async () => {
    const done = busy("Creo il documento…");
    try { const b = await buildBmcDocx(d, S.bmc, S.st.profile); done(); return b; } catch (e) { done(); console.error(e); toast("Errore nella creazione: " + e.message, true); return null; }
  };
  const buttons = h("div", { class: "btnrow wrap" },
    h("a", { class: "btn", href: "#/p/" + encodeURIComponent(p.path) }, "Annulla"),
    h("button", { class: "btn", onclick: () => { if (!printHtml(bmcHtml(d, S.bmc, S.st.profile, esc), "Medical report")) toast("Il browser ha bloccato la finestra: consenti i popup", true); } }, "🖨 Stampa / PDF"),
    h("button", { class: "btn", onclick: async () => { const b = await make(); if (b) saveBlob(b, bmcFileName(d)); } }, "⬇︎ Scarica Word"),
    h("button", { class: "btn primary", onclick: async () => {
      const b = await make(); if (!b) return;
      const done = busy("Salvo nella cartella del paziente…");
      try { await S.drive.upload(p.path, bmcFileName(d), b); toast("Referto salvato"); }
      catch (e) { done(); return toast("Errore nel salvataggio: " + e.message, true); }
      done(); await refresh(true); location.hash = "#/p/" + encodeURIComponent(p.path);
    } }, "Salva su OneDrive"));
  form.append(buttons);
  upd();

  // Dati anagrafici del paziente (se inseriti) precompilano il modulo
  loadMeta(p).then((meta) => {
    const a = meta.anagrafica, full = patientFullName(p, meta);
    if (full !== d.paziente) { d.paziente = full; el.paziente.value = full; if (d.consensoNome === prevName) { d.consensoNome = full; el.consensoNome.value = full; } prevName = full; }
    if (a.nascita && !d.nascita) { d.nascita = a.nascita; el.nascita.value = a.nascita; }
    upd();
  });
  return h("section", {}, form, h("div", { class: "card" }, h("h4", {}, "Anteprima"), prev));
}

// ---------- impostazioni BMC ----------
function cardBmcSettings() {
  const B = { ...S.bmc };
  const box = h("div", { class: "imgbox" });
  const btns = h("div", { class: "btnrow wrap" });
  const show = () => {
    box.replaceChildren(B.logoData ? h("img", { src: B.logoData, alt: "logo BMC" }) : h("span", { class: "muted small" }, "nessun logo"));
    btns.replaceChildren(fileInp, h("button", { class: "btn small", onclick: () => fileInp.click() }, B.logoData ? "Sostituisci logo" : "Carica logo"));
  };
  const fileInp = h("input", { type: "file", accept: "image/*", hidden: true, onchange: async () => {
    if (!fileInp.files[0]) return;
    try { B.logoData = await toPngDataUrl(fileInp.files[0], 1000, false); show(); } catch (e) { toast(e.message, true); }
  } });
  show();
  const t = (k, label) => h("label", {}, label, h("input", { value: B[k] || "", oninput: (e) => (B[k] = e.target.value) }));
  return h("div", { class: "card" }, h("h4", {}, "Cliente BMCh24"),
    box, btns,
    h("div", { class: "grid2" }, t("titolo", "Titolo del referto"), t("direzione", "Direzione sanitaria")),
    h("div", { class: "grid2" }, t("medico", "Medico incaricato (predefinito)"), t("luogo", "Luogo visita (predefinito)")),
    h("div", { class: "btnrow" }, h("button", { class: "btn primary", onclick: async () => {
      const done = busy("Salvo…");
      try { S.bmc = { ...DEFAULT_BMC, ...B }; await S.drive.writeJson(bmcPath(), S.bmc); toast("Impostazioni BMC salvate"); } catch (e) { toast("Errore: " + e.message, true); }
      done();
    } }, "Salva")));
}

async function loadBmc() {
  try { const j = await S.drive.readJson(bmcPath()); if (j) S.bmc = { ...DEFAULT_BMC, ...j }; } catch (e) { console.warn("bmc", e); }
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
  await loadBmc();
  done();
  await refresh();
}

async function boot() {
  const demo = new URLSearchParams(location.search).has("demo");
  if (demo) return startWith(new DemoDrive(CONFIG.rootPath, CONFIG.bmcPath));
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
