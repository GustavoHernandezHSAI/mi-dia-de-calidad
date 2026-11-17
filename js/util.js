// Utilidades de fechas, texto y DOM
export const pad = (n) => String(n).padStart(2, '0');
export const ymd = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
export const parseYmd = (s) => { const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d); };
export const addDays = (s, n) => { const d = parseYmd(s); d.setDate(d.getDate() + n); return ymd(d); };
export const today = () => ymd(new Date());
export const dow = (s) => parseYmd(s).getDay(); // 0 = domingo
export const diffDays = (a, b) => Math.round((parseYmd(b) - parseYmd(a)) / 86400000);
export const toMin = (hhmm) => { if (!hhmm) return null; const [h, m] = hhmm.split(':').map(Number); return h * 60 + m; };
export const toHHMM = (min) => `${pad(Math.floor(min / 60))}:${pad(min % 60)}`;
export const nowMin = () => { const d = new Date(); return d.getHours() * 60 + d.getMinutes(); };
export const roundUp = (min, step = 15) => Math.ceil(min / step) * step;
export const weekStart = (s) => { const d = dow(s); return addDays(s, d === 0 ? -6 : 1 - d); };

export const uid = () => (crypto.randomUUID ? crypto.randomUUID()
  : 'id-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 10));

export const DIAS = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
export const DIAS_C = ['Dom', 'Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb'];
export const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
export const MESES_C = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];

export const fmtShort = (s) => { if (!s) return ''; const d = parseYmd(s); return `${DIAS_C[d.getDay()]} ${d.getDate()} ${MESES_C[d.getMonth()]}`; };
export const fmtLong = (s) => { const d = parseYmd(s); return `${DIAS[d.getDay()]} ${d.getDate()} de ${MESES[d.getMonth()]}`; };
export const fmtDur = (m) => { if (!m) return '0 min'; const h = Math.floor(m / 60), r = m % 60; return h ? (r ? `${h} h ${r} min` : `${h} h`) : `${r} min`; };

export const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
export const norm = (s) => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
export const cap = (s) => (s ? s.charAt(0).toUpperCase() + s.slice(1) : s);

export const $ = (sel, root = document) => root.querySelector(sel);
export const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];

export function relDue(due, ref = today()) {
  if (!due) return '';
  const d = diffDays(ref, due);
  if (d < 0) return `vencida hace ${-d} d`;
  if (d === 0) return 'vence hoy';
  if (d === 1) return 'vence mañana';
  return `vence ${fmtShort(due)}`;
}

export async function copyText(text) {
  try { await navigator.clipboard.writeText(text); return true; }
  catch {
    const ta = document.createElement('textarea'); ta.value = text; document.body.appendChild(ta);
    ta.select(); let ok = false; try { ok = document.execCommand('copy'); } catch { /* */ }
    ta.remove(); return ok;
  }
}
