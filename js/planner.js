// Motor de planificación: jornada, bloques fijos, capacidad, conflictos, sugerencias y reorganización.
import { addDays, diffDays, dow, toMin, toHHMM, nowMin, roundUp, today, uid, fmtShort } from './util.js';

export const DEFAULT_SETTINGS = {
  id: 'settings',
  weekday_start: '09:00',
  weekday_end: '18:00',
  saturday_start: '09:00',
  saturday_end: '13:00',
  saturday_ref: '2026-11-21', // sábado laboral de referencia; se alterna cada semana
  lunch_start: '14:00',
  lunch_min: 60,
  buffer_pct: 20,       // % de la jornada libre que se reserva para imprevistos
  block_min: 60,        // bloque preferido
  block_max: 90,        // bloque máximo cuando hay espacio
  focus_min: 60,        // bloque de concentración protegido
  indicadores_dias: 5,  // primeros días hábiles del mes para revisión de indicadores
  followup_days: 7,     // días esperando a otros antes de marcar un documento para seguimiento
  directora_label: 'Directora de Calidad',
  hospitals: [],
  notifications: false,
  outlook_hide_dupes: true,
  recurring: [
    { id: 'dg', title: 'Reunión con Directora General', days: [1, 2, 3, 4, 5], start: '09:00', end: '10:00', mode: 'siempre', kind: 'dg', note: 'Puede extenderse' },
    { id: 'intramed', title: 'Reunión con Intramed', days: [2, 5], start: '11:00', end: '12:00', mode: 'siempre', kind: 'intramed', note: 'No movible' },
    { id: 'gerentes', title: 'Reunión con gerentes de calidad', days: [5], start: '15:00', end: '18:00', mode: 'siempre', kind: 'gerentes', note: 'Normalmente viernes' },
    { id: 'seguridad', title: 'Reunión con coordinadores de seguridad hospitalaria', days: [2], start: '15:00', end: '18:00', mode: 'manual', kind: 'seguridad', note: 'Solo algunos martes: actívala en el día' },
  ],
};

export const CLASIFICACION = {
  inmediata: 'Acción inmediata',
  programable: 'Programable',
  espera: 'Requiere info / decisión de otro',
  delegable: 'Puede delegarse (con autorización)',
  seguimiento: 'Seguimiento / monitoreo',
  referencia: 'Solo referencia',
};
export const ESTADOS = {
  pendiente: 'Pendiente', programada: 'Programada', en_curso: 'En curso', bloqueada: 'Bloqueada',
  completada: 'Completada', pospuesta: 'Pospuesta', cancelada: 'Cancelada',
};
export const CATEGORIAS = ['Documentos', 'Formatos', 'Normatividad', 'Certificación', 'Indicadores', 'Intramed', 'Coordinación', 'General'];
export const TIPOS_DIA = {
  normal: 'Normal', sabado_descanso: 'Sábado de descanso', sabado_laboral: 'Sábado laboral',
  jornada_especial: 'Jornada especial', visita: 'Visita a hospital (día completo)', inhabil: 'Día inhábil',
};

export const isOpen = (t) => !['completada', 'cancelada'].includes(t.status);
const overlap = (a, b) => a.s < b.e && b.s < a.e;

export function settingsOf(store) {
  const s = store.get('settings', 'settings');
  return { ...DEFAULT_SETTINGS, ...(s || {}), recurring: (s && s.recurring) || DEFAULT_SETTINGS.recurring };
}
export const dayRec = (store, date) => store.get('days', 'day-' + date) || null;

export function isWorkingSaturday(date, S) {
  const weeks = Math.round(diffDays(S.saturday_ref, date) / 7);
  return Math.abs(weeks) % 2 === 0;
}

// Ventana laboral del día (o null si no se trabaja)
export function workWindow(date, S, D) {
  const wd = dow(date);
  const tipo = D?.tipo || 'normal';
  if (tipo === 'inhabil' || tipo === 'sabado_descanso') return null;
  if (tipo === 'jornada_especial' && D.start && D.end) return { s: toMin(D.start), e: toMin(D.end) };
  if (wd === 0) return null;
  if (wd === 6) {
    if (tipo === 'sabado_laboral' || (tipo !== 'sabado_descanso' && isWorkingSaturday(date, S)))
      return { s: toMin(S.saturday_start), e: toMin(S.saturday_end) };
    return null;
  }
  return { s: toMin(S.weekday_start), e: toMin(S.weekday_end) };
}

export function nextWorkingDay(date, S, store, inclusive = false) {
  let d = inclusive ? date : addDays(date, 1);
  for (let i = 0; i < 30; i++) {
    const D = dayRec(store, d);
    if (workWindow(d, S, D) && D?.tipo !== 'visita') return d;
    d = addDays(d, 1);
  }
  return d;
}

// Primeros N días hábiles (lun–vie, sin días marcados como inhábiles) del mes
export function isIndicadoresWindow(date, S, store) {
  const first = date.slice(0, 8) + '01';
  let count = 0;
  for (let d = first; d.slice(0, 7) === date.slice(0, 7); d = addDays(d, 1)) {
    const wd = dow(d);
    const D = dayRec(store, d);
    if (wd >= 1 && wd <= 5 && D?.tipo !== 'inhabil') count++;
    if (d === date) return wd >= 1 && wd <= 5 && D?.tipo !== 'inhabil' && count <= S.indicadores_dias;
    if (count > S.indicadores_dias) return false;
  }
  return false;
}

// Bloques fijos: reuniones recurrentes, comida, eventos registrados y Outlook
export function fixedBlocks(date, ctx) {
  const { S, store, outlook } = ctx;
  const D = dayRec(store, date);
  const win = workWindow(date, S, D);
  const out = [];
  if (!win) return out;
  const wd = dow(date);
  if (D?.tipo === 'visita') {
    out.push({ id: 'visita-' + date, s: win.s, e: win.e, title: 'Visita a hospital' + (D.note ? ': ' + D.note : ''), kind: 'visita', fixed: true });
    return out;
  }
  for (const r of S.recurring) {
    const override = D?.rec?.[r.id];
    let on = r.mode === 'siempre' ? r.days.includes(wd) : false;
    if (override === true && r.days.includes(wd)) on = true;
    if (override === true && r.mode === 'manual') on = true;
    if (override === false) on = false;
    if (!on) continue;
    let e = toMin(r.end);
    if (r.id === 'dg' && D?.dg_end) e = Math.max(e, toMin(D.dg_end));
    out.push({ id: 'rec-' + r.id, s: toMin(r.start), e, title: r.title, kind: r.kind || 'reunion', fixed: true, rec: r.id });
  }
  const ls = toMin(S.lunch_start);
  const lmin = D?.lunch_min || S.lunch_min;
  if (ls >= win.s && ls < win.e) out.push({ id: 'comida', s: ls, e: Math.min(ls + lmin, win.e), title: lmin < 60 ? 'Comida (reducida)' : 'Comida', kind: 'comida', fixed: true });
  for (const ev of store.all('events').filter((x) => x.date === date)) {
    out.push({ id: ev.id, s: toMin(ev.start), e: toMin(ev.end), title: ev.title, kind: ev.kind || 'reunion', fixed: true, event: ev });
  }
  const oevs = (outlook && outlook[date]) || [];
  for (const o of oevs) {
    if (o.showAs === 'free') continue;
    const b = { id: 'ol-' + o.id, s: o.s, e: o.e, title: o.title, kind: 'outlook', fixed: true, allDay: o.allDay };
    if (S.outlook_hide_dupes) {
      // Si un evento de Outlook coincide con una reunión fija configurada, se muestra solo el de Outlook
      for (let i = out.length - 1; i >= 0; i--) {
        const f = out[i];
        if (f.rec && overlap(f, b) && Math.min(f.e, b.e) - Math.max(f.s, b.s) >= (f.e - f.s) / 2) out.splice(i, 1);
      }
    }
    out.push(b);
  }
  return out.sort((a, b) => a.s - b.s);
}

export function taskBlocks(date, store) {
  return store.all('tasks')
    .filter((t) => t.scheduled_date === date && t.scheduled_start && t.status !== 'cancelada')
    .map((t) => ({ id: t.id, s: toMin(t.scheduled_start), e: toMin(t.scheduled_start) + (t.duration_min || 30), title: t.title, kind: 'task', task: t }))
    .sort((a, b) => a.s - b.s);
}

function subtract(win, busy) {
  let free = [{ s: win.s, e: win.e }];
  for (const b of busy) {
    const next = [];
    for (const f of free) {
      if (!overlap(f, b)) { next.push(f); continue; }
      if (b.s > f.s) next.push({ s: f.s, e: b.s });
      if (b.e < f.e) next.push({ s: b.e, e: f.e });
    }
    free = next;
  }
  return free.filter((f) => f.e - f.s >= 15);
}

// Información completa de un día
export function dayInfo(date, ctx) {
  const { S, store } = ctx;
  const D = dayRec(store, date);
  const win = workWindow(date, S, D);
  const tipo = D?.tipo || (dow(date) === 6 ? (win ? 'sabado_laboral' : 'sabado_descanso') : (win ? 'normal' : 'inhabil'));
  const fixed = fixedBlocks(date, ctx);
  const tasks = taskBlocks(date, store);
  const info = { date, D, tipo, win, working: !!win, fixed, tasks, conflicts: {}, free: [], focus: null, capacity: 0, plannable: 0, used: 0 };
  if (!win) {
    tasks.forEach((t) => { if (isOpen(t.task)) info.conflicts[t.id] = ['Día no laboral']; });
    return info;
  }
  const fixedIn = fixed.map((f) => ({ ...f, s: Math.max(f.s, win.s), e: Math.min(f.e, win.e) })).filter((f) => f.e > f.s);
  const fixedMin = mergedMinutes(fixedIn);
  info.capacity = Math.max(0, win.e - win.s - fixedMin);
  info.plannable = D?.tipo === 'visita' ? 0 : Math.floor(info.capacity * (1 - S.buffer_pct / 100));
  info.used = tasks.filter((t) => isOpen(t.task) || t.task.status === 'completada').reduce((a, t) => a + (t.e - t.s), 0);
  // Conflictos
  for (const t of tasks) {
    if (!isOpen(t.task)) continue;
    const c = [];
    if (t.s < win.s || t.e > win.e) c.push('Fuera de horario laboral');
    for (const f of fixed) if (overlap(t, f)) c.push('Choca con ' + f.title);
    for (const o of tasks) if (o !== t && isOpen(o.task) && overlap(t, o)) c.push('Empalme con ' + o.title);
    if (c.length) info.conflicts[t.id] = c;
  }
  const busy = [...fixed, ...tasks.filter((t) => t.task.status !== 'cancelada')];
  info.free = subtract(win, busy);
  // Bloque de concentración protegido: el mayor espacio libre ≥ focus_min
  const nowCut = date === today() ? roundUp(nowMin(), 15) : 0;
  const future = info.free.map((f) => ({ s: Math.max(f.s, nowCut), e: f.e })).filter((f) => f.e > f.s);
  const cand = future.filter((f) => f.e - f.s >= S.focus_min).sort((a, b) => (b.e - b.s) - (a.e - a.s))[0];
  if (cand) info.focus = { s: cand.s, e: Math.min(cand.e, cand.s + S.block_max) };
  info.hasFocus = !!cand || tasks.some((t) => (t.e - t.s) >= S.focus_min && isOpen(t.task));
  return info;
}

function mergedMinutes(blocks) {
  const arr = [...blocks].sort((a, b) => a.s - b.s);
  let total = 0, cs = null, ce = null;
  for (const b of arr) {
    if (cs === null) { cs = b.s; ce = b.e; continue; }
    if (b.s <= ce) ce = Math.max(ce, b.e); else { total += ce - cs; cs = b.s; ce = b.e; }
  }
  if (cs !== null) total += ce - cs;
  return total;
}

// ---------- Priorización ----------
export function score(t, ref = today()) {
  let s = 0;
  if (t.is_urgent) s += 2000;
  if (t.riesgo) s += 1000;
  s += ({ 1: 400, 2: 250, 3: 100, 4: 0 })[t.priority || 3];
  if (t.due_date) {
    const d = diffDays(ref, t.due_date);
    s += d < 0 ? 600 : d === 0 ? 500 : d === 1 ? 400 : d <= 3 ? 300 : d <= 7 ? 150 : 0;
  }
  s += ({ alto: 150, medio: 75 })[t.impacto] || 0;
  if (t.unblocks) s += 80; // desbloquea a otros
  if ((t.duration_min || 30) <= 20) s += 40; // rápida
  return s;
}
export const byScore = (a, b) => score(b) - score(a);

export function flags(t, ctx) {
  const ref = today();
  const f = { overdue: false, atRisk: false, blocked: t.status === 'bloqueada', decision: t.clasificacion === 'espera', delegable: t.clasificacion === 'delegable' && t.delegation_status !== 'autorizada', stale: false };
  if (!isOpen(t)) return f;
  if (t.due_date && t.due_date < ref) f.overdue = true;
  if (t.scheduled_date && t.scheduled_date < ref) f.stale = true;
  if (t.due_date && !f.overdue) {
    const d = diffDays(ref, t.due_date);
    if (d <= 2 && (!t.scheduled_date || t.scheduled_date > t.due_date || d <= 1)) f.atRisk = true;
    if (t.scheduled_date && t.scheduled_date > t.due_date) f.atRisk = true;
  }
  return f;
}

// ---------- Sugerencia automática de horarios ----------
// Devuelve propuestas sin modificar datos: [{task, date, start, dur, part, parts, late}]
export function autoPlan(tasksToPlace, ctx, { from = today(), days = 10, useBuffer = false } = {}) {
  const { S, store } = ctx;
  const sim = {}; // estado simulado por día
  const getDay = (d) => {
    if (sim[d]) return sim[d];
    const info = dayInfo(d, ctx);
    let free = info.free.map((f) => ({ ...f }));
    if (d === today()) { const n = roundUp(nowMin(), 15); free = free.map((f) => ({ s: Math.max(f.s, n), e: f.e })).filter((f) => f.e - f.s >= 15); }
    const remaining = useBuffer ? info.capacity - info.used : info.plannable - info.used;
    sim[d] = { info, free, remaining, focus: info.focus };
    return sim[d];
  };
  const out = [];
  const list = [...tasksToPlace].sort(byScore);
  for (const t of list) {
    let need = t.duration_min || 30;
    const parts = [];
    let d = from;
    let skipped = 0;
    for (let i = 0; i < days * 2 && need > 0; i++, d = addDays(d, 1)) {
      const day = getDay(d);
      if (!day.info.working || day.info.D?.tipo === 'visita' || day.remaining < Math.min(need, 30)) continue;
      const deep = need >= 45;
      // Preferir un solo bloque continuo; dividir solo si la tarea es larga, vence o ya se buscó 2 días
      if (need <= S.block_max && day.remaining >= need) {
        const breaksFocus = (f) => day.focus && f.s <= day.focus.s && f.e >= day.focus.e && !deep && (f.e - f.s - need) < S.focus_min;
        const whole = day.free.find((f) => f.e - f.s >= need && !breaksFocus(f))
          || day.free.find((f) => f.e - f.s >= need);
        if (whole) {
          parts.push({ date: d, start: toHHMM(whole.s), dur: need });
          if (day.focus && whole.s <= day.focus.s && whole.e >= day.focus.e && (deep || whole.e - whole.s - need < S.focus_min)) day.focus = null;
          whole.s += need; day.remaining -= need; need = 0;
          day.free = day.free.filter((f) => f.e - f.s >= 15);
          break;
        }
        const mustSplit = (t.due_date && d >= t.due_date) || skipped >= 2 || parts.length > 0;
        if (!mustSplit) { skipped++; continue; }
      }
      for (const slot of day.free) {
        if (need <= 0 || day.remaining < 15) break;
        const len = slot.e - slot.s;
        // No consumir el bloque de concentración con tareas cortas, salvo que no haya otra opción
        const isFocus = day.focus && slot.s <= day.focus.s && slot.e >= day.focus.e;
        if (isFocus && !deep && day.free.length > 1 && (len - Math.min(need, len)) < S.focus_min) continue;
        let chunk;
        if (need <= S.block_max && need <= len) chunk = need;
        else if (len >= S.block_max && need >= S.block_max) chunk = S.block_max;
        else chunk = Math.min(need, len, S.block_min);
        chunk = Math.min(chunk, day.remaining);
        if (chunk < Math.min(need, 30)) continue;
        parts.push({ date: d, start: toHHMM(slot.s), dur: chunk });
        slot.s += chunk; need -= chunk; day.remaining -= chunk;
        if (isFocus) day.focus = null;
      }
      day.free = day.free.filter((f) => f.e - f.s >= 15);
    }
    parts.forEach((p, i) => out.push({ task: t, ...p, part: i + 1, parts: parts.length, late: !!(t.due_date && p.date > t.due_date), unplaced: false }));
    if (need > 0) out.push({ task: t, unplaced: true, remaining: need });
  }
  return out;
}

// Aplica propuestas: la primera parte actualiza la tarea; las siguientes crean subtareas.
export function applyPlan(proposals, store) {
  const byTask = {};
  proposals.filter((p) => !p.unplaced).forEach((p) => { (byTask[p.task.id] = byTask[p.task.id] || []).push(p); });
  for (const id of Object.keys(byTask)) {
    const ps = byTask[id];
    const base = store.get('tasks', id);
    if (!base) continue;
    if (ps.length === 1) {
      store.put('tasks', { ...base, scheduled_date: ps[0].date, scheduled_start: ps[0].start, duration_min: ps[0].dur, status: base.status === 'pendiente' || base.status === 'pospuesta' ? 'programada' : base.status });
    } else {
      ps.forEach((p, i) => {
        const t = i === 0 ? { ...base } : { ...base, id: uid(), parent_id: base.id, created_at: undefined };
        t.title = `${base.title.replace(/ \(\d+\/\d+\)$/, '')} (${i + 1}/${ps.length})`;
        Object.assign(t, { scheduled_date: p.date, scheduled_start: p.start, duration_min: p.dur, status: 'programada' });
        store.put('tasks', t);
      });
    }
  }
}

// Divide una tarea en partes iguales (sin programar las nuevas)
export function splitTask(t, parts, store) {
  const total = t.duration_min || 60;
  const each = Math.max(15, Math.round(total / parts / 15) * 15);
  // Si ya es una parte (p. ej. "(1/2)"), se nombra como subparte para no repetir títulos
  const isPart = / \(\d+\/\d+\)$/.test(t.title);
  const name = (i) => (isPart ? `${t.title} – parte ${i} de ${parts}` : `${t.title} (${i}/${parts})`);
  store.put('tasks', { ...t, title: name(1), duration_min: each });
  for (let i = 2; i <= parts; i++) {
    store.put('tasks', { ...t, id: uid(), created_at: undefined, parent_id: t.id, title: name(i), duration_min: each, scheduled_date: null, scheduled_start: null, status: 'pendiente' });
  }
}

// ---------- Imprevisto: inserta la atención no programada y reorganiza sin eliminar nada ----------
export function bomberazo(input, ctx) {
  const { S, store } = ctx;
  const changes = [];
  let date = today();
  let info = dayInfo(date, ctx);
  let start = roundUp(nowMin(), 5);
  let outside = false;
  if (!info.working || info.D?.tipo === 'visita' || start >= info.win.e - 15) {
    outside = !info.working || start >= (info.win?.e ?? 0);
    date = nextWorkingDay(date, S, store);
    info = dayInfo(date, ctx);
    start = info.win.s;
  }
  if (start < info.win.s) start = info.win.s;
  // Si el inicio cae dentro de una reunión fija, la urgencia empieza al terminar esa reunión
  for (const f of info.fixed) if (start >= f.s && start < f.e) start = f.e;
  const dur = input.duration_min || 30;
  const urgent = {
    id: uid(), title: input.title, notes: input.notes || '', due_date: input.due_date || null,
    priority: 1, is_urgent: true, riesgo: !!input.riesgo, clasificacion: 'inmediata', status: 'programada',
    duration_min: dur, category: input.category || 'General', area: input.area || '',
    waiting_on: input.requires || '', next_action: input.next_action || '', source: 'bomberazo',
    scheduled_date: date, scheduled_start: toHHMM(start),
  };
  // Si la urgencia cruza una reunión fija, se deja hasta el inicio de la reunión (el resto se ve como conflicto)
  const uBlock = { s: start, e: start + dur };
  // Reacomodo de tareas del día a partir del inicio de la urgencia
  const affected = taskBlocks(date, store).filter((b) => isOpen(b.task) && b.task.status !== 'en_curso' && b.e > start).sort((a, b) => a.s - b.s);
  const fixedBusy = info.fixed.map((f) => ({ s: f.s, e: f.e }));
  const affIds = new Set(affected.map((b) => b.id));
  const kept = taskBlocks(date, store).filter((b) => !affIds.has(b.id) && b.task.status !== 'cancelada').map((b) => ({ s: b.s, e: b.e }));
  const busy = [...fixedBusy, ...kept, uBlock];
  const postponed = [];
  for (const b of affected) {
    const len = b.e - b.s;
    const conflictNow = busy.some((x) => overlap(x, b));
    if (!conflictNow) { busy.push({ s: b.s, e: b.e }); continue; }
    const free = subtract({ s: Math.max(start, info.win.s), e: info.win.e }, busy).filter((f) => f.e - f.s >= len);
    if (free.length) {
      const ns = free[0].s;
      busy.push({ s: ns, e: ns + len });
      changes.push({ task_id: b.task.id, title: b.task.title, before: snap(b.task), after: { scheduled_date: date, scheduled_start: toHHMM(ns), status: b.task.status }, reason: `Se movió de ${toHHMM(b.s)} a ${toHHMM(ns)} por el imprevisto` });
    } else postponed.push(b.task);
  }
  // Las que ya no caben hoy pasan a los siguientes días hábiles
  if (postponed.length) {
    store.put('tasks', urgent, { silent: true }); // ocupa su lugar en la simulación
    const tmp = postponed.map((t) => ({ ...t, scheduled_date: null, scheduled_start: null }));
    // Simulación sin las tareas pospuestas
    const ctx2 = { ...ctx, store: maskStore(store, tmp.map((t) => t.id), changes) };
    const plan = autoPlan(tmp, ctx2, { from: nextWorkingDay(date, S, store), days: 10 });
    for (const t of postponed) {
      const p = plan.find((x) => x.task.id === t.id && !x.unplaced && x.part === 1);
      if (p) changes.push({ task_id: t.id, title: t.title, before: snap(t), after: { scheduled_date: p.date, scheduled_start: p.start, status: 'pospuesta' }, reason: `Sin espacio hoy: se pospuso a ${fmtShort(p.date)} ${p.start}` + (t.due_date && p.date > t.due_date ? ' (queda después de su fecha límite)' : '') });
      else changes.push({ task_id: t.id, title: t.title, before: snap(t), after: { scheduled_date: null, scheduled_start: null, status: 'pospuesta' }, reason: 'Sin espacio en los próximos días: queda sin programar' });
    }
  }
  if (input.blocks_task_id) {
    const bt = store.get('tasks', input.blocks_task_id);
    if (bt) changes.push({ task_id: bt.id, title: bt.title, before: snap(bt), after: { status: 'bloqueada', blocked_reason: 'En espera por imprevisto: ' + input.title }, reason: 'Queda bloqueada hasta resolver el imprevisto' });
  }
  return { urgent, changes, date, start: toHHMM(start), outside };
}
const snap = (t) => ({ scheduled_date: t.scheduled_date || null, scheduled_start: t.scheduled_start || null, status: t.status, blocked_reason: t.blocked_reason || '' });

// Vista del almacén para simular: oculta tareas pospuestas y aplica movimientos ya decididos
function maskStore(store, hideIds, changes) {
  const moved = Object.fromEntries(changes.map((c) => [c.task_id, c.after]));
  return {
    ...store,
    get: (c, id) => store.get(c, id),
    all: (c) => {
      const arr = store.all(c);
      if (c !== 'tasks') return arr;
      return arr.filter((t) => !hideIds.includes(t.id)).map((t) => (moved[t.id] ? { ...t, ...moved[t.id] } : t));
    },
  };
}

export function applyChanges(changes, store, dir = 'after') {
  for (const c of changes) {
    const t = store.get('tasks', c.task_id);
    if (t) store.put('tasks', { ...t, ...c[dir], postponed_count: (t.postponed_count || 0) + (dir === 'after' && c.after.status === 'pospuesta' ? 1 : 0) });
  }
}
