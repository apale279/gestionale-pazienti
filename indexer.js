// Indicizzazione della cartella pazienti e piano di riordino dei file sciolti.
// Funzioni pure o che usano solo l'interfaccia "drive": nessuna cancellazione.

const join = (...p) => p.filter(Boolean).join("/");
const norm = (s) => s.normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();

// Parole che descrivono il tipo di documento e NON fanno parte del nome del paziente.
const STOP = new Set(`vp visita visite anamnesi esami esame referto referti certificato certificati report medical bmc
nuovo nuova documento microsoft word scheda paziente pazienti cartella clinica appunti ecg ee rmn tc eco ecott holter
riassunto relazione consenso privacy template modello carta intestata agg aggiornato aggiornata copia di del della de e
pdf doc docx zip jpeg jpg png xlsx dotx gen feb mar apr mag giu lug ago set ott nov dic v1 v2 v3 telemedicina`.split(/\s+/));

const MODEL_RE = /carta intestata|template|modello|scheda paziente/i;
const MODEL_EXT = /\.(dotx|dot|xltx|potx)$/i;

export function sanitizeName(s) {
  return s.replace(/[\\/:*?"<>|#%]/g, " ").replace(/\s+/g, " ").trim().replace(/[. ]+$/, "");
}

export function stripExt(name) { return name.replace(/\.[^.]{1,5}$/, ""); }
export function extOf(name) { const m = /\.([^.]{1,5})$/.exec(name); return m ? m[1].toLowerCase() : ""; }

export function fileDate(name, modified) {
  const m = /(\d{4})[_\-.](\d{2})[_\-.](\d{2})/.exec(name);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  return (modified || "").slice(0, 10);
}

// Elenco di parole {n: normalizzata, o: originale} di un nome file, senza date/numeri.
function words(filename) {
  const base = stripExt(filename).replace(/\d{4}[_\-.]\d{2}[_\-.]\d{2}/g, " ");
  return base.split(/[\s_\-.()\[\],]+/).filter(Boolean)
    .map((o) => ({ n: norm(o), o }))
    .filter((w) => w.n && !/^\d+$/.test(w.n));
}
const titleCase = (s) => s.charAt(0).toUpperCase() + s.slice(1).toLowerCase();

// ---------------------------------------------------------------------------
// Legge tutte le cartelle-paziente a livelli (una richiesta cumulativa per livello, non una per cartella).
async function collectAll(drive, roots, onProgress) {
  const result = new Map(roots.map((r) => [r, []]));
  let frontier = roots.map((r) => ({ root: r, path: r, sub: "", depth: 0 }));
  let known = frontier.length, done = 0;
  while (frontier.length) {
    const lists = await drive.listMany(frontier.map((f) => f.path));
    const next = [];
    frontier.forEach((fr, i) => {
      for (const x of lists[i] || []) {
        if (x.isFolder) {
          if (fr.depth < 4 && !/^[.]/.test(x.name)) next.push({ root: fr.root, path: x.path, sub: join(fr.sub, x.name), depth: fr.depth + 1 });
        } else if (x.name !== ".paziente.json") {
          result.get(fr.root).push({ ...x, sub: fr.sub, date: fileDate(x.name, x.modified) });
        }
      }
    });
    done += frontier.length; known += next.length;
    onProgress(done, known);
    frontier = next;
  }
  return result;
}

export async function scan(drive, cfg, onProgress = () => {}) {
  const root = cfg.rootPath;
  const hidden = new Set([...(cfg.ignoredFolders || []), ...(cfg.hiddenFolders || [])].map((s) => s.toLowerCase()));
  const containers = new Set((cfg.extraPatientContainers || []).map((s) => s.toLowerCase()));

  const [top, bmcTop] = await drive.listMany([root, cfg.bmcPath].filter(Boolean));
  if (!top) { const e = new Error("Cartella non trovata: " + root); e.status = 404; throw e; }

  const contItems = top.filter((it) => it.isFolder && containers.has(it.name.toLowerCase()));
  const contLists = contItems.length ? await drive.listMany(contItems.map((c) => c.path)) : [];

  const patientFolders = []; // {name, path, container}
  const loose = [];
  for (const it of top) {
    const ln = it.name.toLowerCase();
    if (it.isFolder) {
      if (containers.has(ln)) {
        for (const sub of contLists[contItems.indexOf(it)] || []) {
          if (sub.isFolder) { if (!hidden.has(sub.name.toLowerCase())) patientFolders.push({ name: sub.name, path: sub.path, container: it.name }); }
          else loose.push({ ...sub, container: it.name });
        }
      } else if (!/^[._]/.test(it.name) && !hidden.has(ln)) {
        patientFolders.push({ name: it.name, path: it.path, container: "" });
      }
    } else {
      loose.push({ ...it, container: "" });
    }
  }
  // Pazienti BMC: una sottocartella per paziente (i file sciolti nella cartella BMC, es. contratti, si ignorano)
  const bmcFolders = (bmcTop || []).filter((x) => x.isFolder && !/^[._]/.test(x.name) && !hidden.has(x.name.toLowerCase()));

  const files = await collectAll(drive, [...patientFolders.map((f) => f.path), ...bmcFolders.map((f) => f.path)], onProgress);

  // Unisce cartelle con lo stesso nome (es. una in radice e una in .Pazienti)
  const byKey = new Map();
  for (const pf of patientFolders) {
    const k = norm(pf.name), fl = files.get(pf.path);
    const ex = byKey.get(k);
    if (!ex) byKey.set(k, { name: pf.name, path: pf.path, folders: [pf.path], files: fl, group: "" });
    else { ex.folders.push(pf.path); ex.files.push(...fl); if (pf.container === "") { ex.path = pf.path; ex.name = pf.name; } }
  }
  const byDate = (a, b) => (b.date || "").localeCompare(a.date || "");
  const patients = [...byKey.values()].map((p) => {
    p.files.sort(byDate);
    p.lastDate = p.files.length ? p.files[0].date : "";
    return p;
  }).sort((a, b) => a.name.localeCompare(b.name, "it"));
  const bmc = bmcFolders.map((f) => {
    const fl = files.get(f.path).sort(byDate);
    return { name: f.name, path: f.path, folders: [f.path], files: fl, group: "BMC", lastDate: fl.length ? fl[0].date : "" };
  }).sort((a, b) => a.name.localeCompare(b.name, "it"));

  // Classifica i file sciolti: modelli (carta intestata, template…) oppure candidati ai pazienti
  const models = [], candidates = [];
  for (const f of loose) {
    const looksModel = MODEL_RE.test(f.name) || MODEL_EXT.test(f.name);
    if (looksModel && !matchPatient(f.name, patients)) models.push(f);
    else candidates.push(f);
  }
  return { patients, bmc, models, loose: candidates };
}

// ---------------------------------------------------------------------------
function patientTokens(name) { return words(name).map((w) => w.n); }

function matchPatient(filename, patients) {
  const ws = words(filename);
  const all = new Set(ws.map((w) => w.n));
  const nameT = ws.filter((w) => !STOP.has(w.n)).map((w) => w.n);
  let best = null, bestScore = 0;
  for (const p of patients) {
    const pt = patientTokens(p.name);
    if (!pt.length) continue;
    if (pt.length === 1) {
      if (nameT.length === 1 && nameT[0] === pt[0] && 1 > bestScore) { best = p; bestScore = 1; }
    } else if (pt.every((t) => all.has(t)) && pt.length > bestScore) { best = p; bestScore = pt.length; }
  }
  return best;
}

// Costruisce il piano: per ogni file sciolto propone un paziente (esistente o nuovo).
export function buildPlan(sc) {
  const rows = [];
  const newGroups = new Map(); // chiave ordinata -> {name, tokens}
  const pending = [];

  for (const f of sc.loose) {
    const p = matchPatient(f.name, sc.patients);
    if (p) { rows.push({ file: f, kind: "existing", target: p.name, confidence: "alta" }); continue; }
    const ws = words(f.name).filter((w) => !STOP.has(w.n));
    if (!ws.length) { rows.push({ file: f, kind: "skip", target: "", confidence: "nessuna" }); continue; }
    pending.push({ f, ws });
  }

  // Un solo termine che coincide con una parola del nome di un unico paziente
  const surnameIndex = new Map();
  for (const p of sc.patients) for (const t of patientTokens(p.name)) {
    if (!surnameIndex.has(t)) surnameIndex.set(t, []);
    surnameIndex.get(t).push(p);
  }

  for (const x of pending.filter((x) => x.ws.length >= 2)) {
    const key = x.ws.map((w) => w.n).sort().join(" ");
    if (!newGroups.has(key)) newGroups.set(key, { name: x.ws.map((w) => titleCase(w.o)).join(" "), tokens: new Set(x.ws.map((w) => w.n)) });
  }
  const placed = new Set();
  for (const x of pending) {
    if (x.ws.length >= 2) {
      const g = newGroups.get(x.ws.map((w) => w.n).sort().join(" "));
      rows.push({ file: x.f, kind: "new", target: g.name, confidence: "media" });
      placed.add(x);
    }
  }
  for (const x of pending) {
    if (placed.has(x)) continue;
    const t = x.ws[0].n;
    const g = [...newGroups.values()].filter((g) => g.tokens.has(t));
    const ex = surnameIndex.get(t) || [];
    if (g.length === 1) rows.push({ file: x.f, kind: "new", target: g[0].name, confidence: "media" });
    else if (ex.length === 1) rows.push({ file: x.f, kind: "existing", target: ex[0].name, confidence: "media" });
    else rows.push({ file: x.f, kind: "new", target: titleCase(x.ws[0].o), confidence: "bassa" });
  }
  rows.sort((a, b) => a.file.name.localeCompare(b.file.name, "it"));
  return rows;
}

// Esegue il piano. Non sovrascrive e non cancella: in caso di nome uguale il file viene rinominato.
export async function applyPlan(drive, cfg, sc, rows, onStep = () => {}) {
  const log = [];
  const made = new Set();
  let i = 0;
  for (const r of rows) {
    onStep(++i, rows.length);
    if (r.kind === "skip" || !r.target.trim()) continue;
    try {
      let dest;
      if (r.kind === "existing") {
        dest = (sc.patients.find((p) => p.name === r.target) || {}).path;
        if (!dest) throw new Error("paziente non trovato");
      } else {
        const folder = sanitizeName(r.target);
        if (!folder) throw new Error("nome non valido");
        dest = join(cfg.rootPath, folder);
        if (!made.has(dest)) {
          const existing = sc.patients.find((p) => norm(p.name) === norm(folder));
          if (existing) dest = existing.path; else await drive.mkdir(cfg.rootPath, folder);
          made.add(dest);
        }
      }
      await drive.move(r.file.path, dest);
      log.push({ name: r.file.name, ok: true, to: r.target });
    } catch (e) {
      log.push({ name: r.file.name, ok: false, error: e.message });
    }
  }
  return log;
}
