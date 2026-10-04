// Misurazione dei tempi di avvio: serve a capire dove si perde tempo (Impostazioni → Diagnostica velocità).
export const PERF = [];
const T0 = performance.now();
export function mark(label, ms) {
  PERF.push({ at: Math.round(performance.now() - T0), label, ms: ms == null ? null : Math.round(ms) });
}
export async function timed(label, fn) {
  const s = performance.now();
  try { return await fn(); } finally { mark(label, performance.now() - s); }
}
export function report() {
  return PERF.map((p) => `${String(p.at).padStart(6)} ms  ${p.label}${p.ms == null ? "" : `  (${p.ms} ms)`}`).join("\n");
}
