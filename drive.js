// Strato di accesso ai file. Due implementazioni con la stessa interfaccia:
//  - GraphDrive: OneDrive reale (Microsoft Graph)
//  - DemoDrive: dati inventati in memoria, per provare l'interfaccia
// L'app NON cancella mai nulla: nessuna funzione di eliminazione esiste qui.
import { getToken } from "./auth.js";

const BASE = "https://graph.microsoft.com/v1.0/me/drive";
const enc = (p) => p.split("/").filter(Boolean).map(encodeURIComponent).join("/");
const join = (...parts) => parts.filter(Boolean).join("/");
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function toItem(parentPath, x) {
  return {
    name: x.name,
    isFolder: !!x.folder,
    size: x.size || 0,
    modified: x.lastModifiedDateTime || "",
    webUrl: x.webUrl || "",
    driveId: (x.parentReference && x.parentReference.driveId) || "",
    path: join(parentPath, x.name),
  };
}

export class GraphDrive {
  isDemo = false;

  async api(url, opts = {}, retry = 0) {
    const token = await getToken();
    const headers = { Authorization: "Bearer " + token, ...(opts.headers || {}) };
    const res = await fetch(url, { ...opts, headers });
    if ((res.status === 429 || res.status === 503) && retry < 4) {
      const wait = (Number(res.headers.get("Retry-After")) || 2) * 1000;
      await sleep(wait);
      return this.api(url, opts, retry + 1);
    }
    return res;
  }

  ref(path) { return path ? `${BASE}/root:/${enc(path)}:` : `${BASE}/root`; }

  async json(url, opts) {
    const res = await this.api(url, opts);
    if (!res.ok) {
      let msg = res.statusText;
      try { msg = (await res.json()).error.message || msg; } catch {}
      const e = new Error(`${res.status} ${msg}`);
      e.status = res.status;
      throw e;
    }
    return res.status === 204 ? null : res.json();
  }

  async list(path) {
    const out = [];
    let url = `${this.ref(path)}/children?$top=200&$select=id,name,size,folder,file,webUrl,lastModifiedDateTime,parentReference`;
    while (url) {
      const j = await this.json(url);
      for (const x of j.value) out.push(toItem(path, x));
      url = j["@odata.nextLink"] || null;
    }
    return out;
  }

  async stat(path) {
    try {
      const x = await this.json(`${this.ref(path)}?$select=id,name,size,folder,file,webUrl,lastModifiedDateTime,parentReference`);
      return toItem(path.split("/").slice(0, -1).join("/"), x);
    } catch (e) { if (e.status === 404) return null; throw e; }
  }

  async mkdir(parent, name) {
    try {
      const x = await this.json(`${this.ref(parent)}/children`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name, folder: {}, "@microsoft.graph.conflictBehavior": "fail" }),
      });
      return toItem(parent, x);
    } catch (e) {
      if (e.status === 409) return this.stat(join(parent, name));
      throw e;
    }
  }

  async ensureFolder(path) {
    const parts = path.split("/").filter(Boolean);
    let cur = "";
    for (const p of parts) {
      if (!(await this.stat(join(cur, p)))) await this.mkdir(cur, p);
      cur = join(cur, p);
    }
  }

  // conflict: "rename" (non sovrascrive mai) oppure "replace" (solo per le impostazioni dell'app)
  async upload(parent, name, blob, conflict = "rename", onProgress) {
    const target = join(parent, name);
    if (blob.size < 4 * 1024 * 1024) {
      const x = await this.json(`${this.ref(target)}/content?@microsoft.graph.conflictBehavior=${conflict}`, {
        method: "PUT",
        headers: { "Content-Type": blob.type || "application/octet-stream" },
        body: blob,
      });
      onProgress && onProgress(1);
      return toItem(parent, x);
    }
    const s = await this.json(`${this.ref(target)}/createUploadSession`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": conflict } }),
    });
    const CH = 327680 * 32;
    let pos = 0, last = null;
    while (pos < blob.size) {
      const end = Math.min(pos + CH, blob.size);
      const res = await fetch(s.uploadUrl, {
        method: "PUT",
        headers: { "Content-Range": `bytes ${pos}-${end - 1}/${blob.size}` },
        body: blob.slice(pos, end),
      });
      if (!res.ok && res.status !== 202) throw new Error("Caricamento fallito: " + res.status);
      if (res.status === 200 || res.status === 201) last = await res.json();
      pos = end;
      onProgress && onProgress(pos / blob.size);
    }
    return last ? toItem(parent, last) : null;
  }

  async move(path, destFolder, newName) {
    const dest = await this.stat(destFolder);
    if (!dest) throw new Error("Cartella di destinazione non trovata: " + destFolder);
    const idRes = await this.json(`${this.ref(path)}?$select=id`);
    const dres = await this.json(`${this.ref(destFolder)}?$select=id`);
    const body = { parentReference: { id: dres.id }, "@microsoft.graph.conflictBehavior": "rename" };
    if (newName) body.name = newName;
    const x = await this.json(`${BASE}/items/${idRes.id}`, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    return toItem(destFolder, x);
  }

  async readJson(path) {
    const res = await this.api(`${this.ref(path)}/content`);
    if (res.status === 404) return null;
    if (!res.ok) throw new Error("Lettura fallita: " + res.status);
    return res.json();
  }

  async writeJson(path, obj) {
    const parts = path.split("/");
    const name = parts.pop();
    await this.ensureFolder(parts.join("/"));
    return this.upload(parts.join("/"), name, new Blob([JSON.stringify(obj, null, 1)], { type: "application/json" }), "replace");
  }

  // Contenuto del file come Blob (per mostrarlo nel browser invece di scaricarlo)
  async blob(path) {
    const res = await this.api(`${this.ref(path)}/content`);
    if (!res.ok) throw new Error("Lettura fallita: " + res.status);
    return res.blob();
  }

  async downloadUrl(path) {
    const j = await this.json(`${this.ref(path)}?select=id,@microsoft.graph.downloadUrl`);
    return j["@microsoft.graph.downloadUrl"];
  }
}

// ---------------------------------------------------------------------------
export class DemoDrive {
  isDemo = true;

  constructor(root, bmc) {
    this.root = root;
    this.bmc = bmc;
    this.nodes = new Map(); // path -> {isFolder,size,modified,blob}
    this.seed();
  }

  put(path, isFolder, size = 0, modified = "2026-05-10T10:00:00Z", blob = null) {
    const parts = path.split("/");
    for (let i = 1; i < parts.length; i++) {
      const p = parts.slice(0, i).join("/");
      if (!this.nodes.has(p)) this.nodes.set(p, { isFolder: true, size: 0, modified });
    }
    this.nodes.set(path, { isFolder, size, modified, blob });
  }

  seed() {
    const R = this.root;
    const f = (p, size, mod) => this.put(join(R, p), false, size, mod);
    this.put(R, true);
    // Cartelle paziente (nomi inventati)
    f("Rossi Mario/2026_03_30 visita.pdf", 210000, "2026-03-30T09:00:00Z");
    f("Rossi Mario/2026_04_10 holter ecg.jpeg", 1800000, "2026-04-10T09:00:00Z");
    f("Rossi Mario/2026_05_06 ecoTT.jpeg", 950000, "2026-05-06T09:00:00Z");
    f("Rossi Mario/Esami eseguiti/2025_12_03 EE.jpeg", 700000, "2025-12-03T09:00:00Z");
    f("Bianchi Anna/Cartella clinica Anna Bianchi.docx", 52000, "2026-03-26T09:00:00Z");
    f("Bianchi Anna/Cartella clinica Anna Bianchi.pdf", 180000, "2026-03-26T09:00:00Z");
    this.put(join(R, "Gialli Franco"), true);
    f(".Pazienti/Neri Giulia/2026_02_12_VP Neri.pdf", 130000, "2026-02-12T09:00:00Z");
    f(".Pazienti/Neri Giulia/2026_01_27_RMN bacino.pdf", 2500000, "2026-01-27T09:00:00Z");
    f(".Pazienti/Verdi Paola.docx", 41000, "2026-04-02T09:00:00Z");
    f("Referti assistenze/Template referti.docx", 30000, "2026-01-10T09:00:00Z");
    if (this.bmc) {
      this.put(join(this.bmc, "Esposito Carla/2026_06_01_BMC Medical report_Esposito Carla.docx"), false, 60000, "2026-06-01T09:00:00Z");
      this.put(join(this.bmc, "Ferrari Luigi"), true);
      this.put(join(this.bmc, "Contratto cliente.pdf"), false, 400000, "2023-07-21T09:00:00Z");
    }
    // File sciolti
    f("Verdi Paola.pdf", 99000, "2026-04-02T09:00:00Z");
    f("VERDI PAOLA esami.zip", 5400000, "2026-04-03T09:00:00Z");
    f("Gialli Franco.docx", 38000, "2026-05-01T09:00:00Z");
    f("Gialli Franco.pdf", 120000, "2026-05-01T09:00:00Z");
    f("Anamnesi Luca Conti.docx", 36000, "2026-06-01T09:00:00Z");
    f("VP sara blu.docx", 37000, "2026-06-05T09:00:00Z");
    f("VP sara blu.pdf", 115000, "2026-06-05T09:00:00Z");
    f("Certificato marco.pdf", 80000, "2026-06-09T09:00:00Z");
    f("BMC medical report.pdf", 300000, "2026-07-01T09:00:00Z");
    f("Carta intestata.dotx", 25000, "2026-01-01T09:00:00Z");
    f("Scheda paziente v1.docx", 33000, "2026-01-02T09:00:00Z");
  }

  async list(path) {
    const out = [];
    const prefix = path + "/";
    for (const [p, n] of this.nodes) {
      if (!p.startsWith(prefix)) continue;
      const rest = p.slice(prefix.length);
      if (rest.includes("/")) continue;
      out.push({ name: rest, isFolder: n.isFolder, size: n.size, modified: n.modified, webUrl: "", path: p });
    }
    return out;
  }
  async stat(path) {
    const n = this.nodes.get(path);
    if (!n) return null;
    return { name: path.split("/").pop(), isFolder: n.isFolder, size: n.size, modified: n.modified, webUrl: "", path };
  }
  async mkdir(parent, name) { this.put(join(parent, name), true); return this.stat(join(parent, name)); }
  async ensureFolder(path) { if (!this.nodes.has(path)) this.put(path, true); }
  unique(parent, name) {
    let n = name, i = 1;
    const dot = name.lastIndexOf(".");
    const base = dot > 0 ? name.slice(0, dot) : name, ext = dot > 0 ? name.slice(dot) : "";
    while (this.nodes.has(join(parent, n))) n = `${base} ${++i}${ext}`;
    return n;
  }
  async upload(parent, name, blob, conflict = "rename", onProgress) {
    const n = conflict === "replace" ? name : this.unique(parent, name);
    this.put(join(parent, n), false, blob.size, new Date().toISOString(), blob);
    onProgress && onProgress(1);
    return this.stat(join(parent, n));
  }
  async move(path, destFolder, newName) {
    const n = this.nodes.get(path);
    const name = this.unique(destFolder, newName || path.split("/").pop());
    this.nodes.delete(path);
    this.put(join(destFolder, name), n.isFolder, n.size, n.modified, n.blob);
    return this.stat(join(destFolder, name));
  }
  async readJson(path) {
    const n = this.nodes.get(path);
    return n && n.blob ? JSON.parse(await n.blob.text()) : null;
  }
  async writeJson(path, obj) {
    const parts = path.split("/");
    const name = parts.pop();
    return this.upload(parts.join("/"), name, new Blob([JSON.stringify(obj)], { type: "application/json" }), "replace");
  }
  async blob(path) {
    const n = this.nodes.get(path);
    return n.blob || new Blob([`Documento di prova (demo)
${path}`], { type: "text/plain" });
  }
  async downloadUrl(path) {
    const n = this.nodes.get(path);
    const blob = n.blob || new Blob([`Documento di prova (demo)\n${path}`], { type: "text/plain" });
    return URL.createObjectURL(blob);
  }
}
