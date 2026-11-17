// Generación de textos: cierre de jornada y resumen para la Directora de Calidad.
import { fmtLong, fmtShort, today, addDays, diffDays, toMin, toHHMM, relDue } from './util.js';
import { isOpen, flags, byScore, dayInfo, nextWorkingDay } from './planner.js';
import { stageOf, isDocDone, canchaOf, FLOWS } from './flows.js';

const withArea = (t) => `${t.title}${t.area && !t.title.includes(t.area) ? ` (${t.area})` : ''}`;
const line = (t) => `• ${withArea(t)}`;
const sec = (title, items) => (items.length ? `${title}\n${items.join('\n')}\n` : '');

export function riesgos(ctx, horizon = 3) {
  const ref = today();
  return ctx.store.all('tasks').filter(isOpen).filter((t) => {
    const f = flags(t, ctx);
    return f.overdue || f.atRisk || (t.due_date && diffDays(ref, t.due_date) <= horizon && t.priority <= 2);
  }).sort(byScore);
}

export function planText(date, ctx) {
  const info = dayInfo(date, ctx);
  const blocks = [...info.fixed.map((f) => ({ ...f, t: 'f' })), ...info.tasks.filter((t) => isOpen(t.task)).map((t) => ({ ...t, t: 't' }))].sort((a, b) => a.s - b.s);
  if (!info.working) return `${fmtLong(date)}: día no laboral.`;
  const rows = blocks.map((b) => `${toHHMM(b.s)}–${toHHMM(b.e)}  ${b.t === 'f' ? '[' + b.title + ']' : b.title}`);
  return `${fmtLong(date)}\n${rows.join('\n') || 'Sin bloques programados.'}`;
}

// a = respuestas del asistente de cierre
export function closureOutputs(a, ctx) {
  const { store, S } = ctx;
  const date = a.date;
  const done = a.done.map((id) => store.get('tasks', id)).filter(Boolean);
  const pend = a.pending.map((id) => store.get('tasks', id)).filter(Boolean);
  const urg = store.all('bomberazos').filter((b) => b.date === date).map((b) => store.get('tasks', b.task_id)).filter(Boolean);
  const deleg = a.delegate.map((id) => store.get('tasks', id)).filter(Boolean);
  const espera = store.all('tasks').filter((t) => isOpen(t) && t.clasificacion === 'espera');
  const next = nextWorkingDay(date, S, store);
  const risk = riesgos(ctx).slice(0, 8);

  const dl = docLines(ctx, date, date);
  const resumen = [
    `Cierre de jornada — ${fmtLong(date)}`,
    '',
    sec('Completado:', done.map(line)) || 'Completado: sin tareas marcadas.\n',
    sec('Pendiente:', pend.map((t) => `${line(t)}${a.resched[t.id] ? ` → ${a.resched[t.id] === 'none' ? 'sin fecha' : fmtShort(a.resched[t.id])}` : ''}`)),
    a.blockers ? `Bloqueos:\n${a.blockers}\n` : '',
    sec('Imprevistos atendidos:', urg.map(line)) + (a.urgent ? `${a.urgent}\n` : ''),
    dl.length ? dl.join('\n') + '\n' : '',
    a.missing || espera.length ? `Información o decisiones pendientes:\n${[...espera.map((t) => `${line(t)}${t.waiting_on ? ' — depende de: ' + t.waiting_on : ''}`), a.missing].filter(Boolean).join('\n')}\n` : '',
  ].filter(Boolean).join('\n').replace(/\n{3,}/g, '\n\n').trim();

  const reprogramados = pend.filter((t) => a.resched[t.id]).map((t) => `${line(t)} → ${a.resched[t.id] === 'none' ? 'sin fecha asignada' : fmtShort(a.resched[t.id])}`);
  const riesgosTxt = risk.map((t) => `${line(t)} — ${t.due_date ? relDue(t.due_date) : 'prioridad alta'}${t.status === 'bloqueada' ? ' (bloqueada)' : ''}`);
  const plan = planText(next, ctx);
  const mensaje = bossMessage({ docs: dl, done, pend, urg, risk, blocked: a.blockers, decisions: [...espera.map((t) => t.title + (t.waiting_on ? ` (de: ${t.waiting_on})` : '')), a.missing].filter(Boolean), deleg, label: S.directora_label, period: 'hoy' });
  return { resumen, reprogramados, riesgos: riesgosTxt, plan, mensaje };
}

// Líneas de avance documental y seguimientos para un periodo [from, to]
export function docLines(ctx, from, to) {
  const { store, S } = ctx;
  const docs = store.all('docs');
  const inR = (iso) => iso && iso.slice(0, 10) >= from && iso.slice(0, 10) <= to;
  const name = (d) => `${d.code ? d.code + ' ' : ''}${d.name}`;
  const done = docs.filter((d) => isDocDone(d) && inR(d.completed_at));
  const moved = docs.filter((d) => !isDocDone(d) && (d.log || []).some((l) => inR(l.at) && /^Etapa:/.test(l.text)));
  const open = docs.filter((d) => !isDocDone(d));
  const mine = open.filter((d) => canchaOf(d) === 'calidad');
  const days = (d) => (d.stage_since ? Math.round((Date.parse(to) - Date.parse(d.stage_since.slice(0, 10))) / 864e5) : 0);
  const waitingLong = open.filter((d) => canchaOf(d) !== 'calidad' && days(d) > (S.followup_days || 7));
  const lines = [];
  if (docs.length) {
    lines.push(`Avance documental: ${open.length} en proceso (${mine.length} en revisión o trabajo de Calidad, ${open.length - mine.length} con dueños de proceso u otras áreas)${done.length ? `; ${done.length} terminado(s) en el periodo` : ''}.`);
    if (done.length) lines.push(...done.slice(0, 6).map((d) => `- Terminado: ${name(d)}`));
    if (moved.length) lines.push(...moved.slice(0, 6).map((d) => `- Avanzó a “${stageOf(d).label}”: ${name(d)}`));
    if (waitingLong.length) lines.push(`- En espera de respuesta más de ${S.followup_days || 7} días: ${waitingLong.slice(0, 5).map(name).join('; ')}${waitingLong.length > 5 ? '…' : ''}`);
    const toIntr = open.filter((d) => d.stage === 'intramed').length;
    if (toIntr) lines.push(`- Pendientes de cargar en Intramed: ${toIntr}`);
  }
  const projs = store.all('projects').filter((p) => p.status === 'activo');
  if (projs.length) {
    lines.push('Seguimientos activos:');
    for (const p of projs.slice(0, 6)) {
      const pd = docs.filter((d) => d.project_id === p.id);
      const et = p.etapas || [];
      const prog = pd.length ? `${pd.filter(isDocDone).length}/${pd.length} terminados` : et.length ? `${et.filter((e) => e.done).length}/${et.length} etapas` : 'en curso';
      lines.push(`- ${p.name}: ${prog}${p.next_action ? '. Siguiente: ' + p.next_action : ''}`);
    }
  }
  return lines;
}
void FLOWS;

export function bossMessage({ done, pend, urg = [], risk = [], blocked = '', decisions = [], deleg = [], label = 'Directora de Calidad', period = 'hoy', docs = [] }) {
  const L = (arr) => arr.map((t) => `- ${typeof t === 'string' ? t : withArea(t)}`).join('\n');
  const crit = pend.filter((t) => t.priority <= 2 || t.riesgo);
  const parts = [`${label}, le comparto el resumen de ${period === 'hoy' ? 'hoy' : 'la semana'}:`];
  if (done.length) parts.push(`Avances:\n${L(done.slice(0, 8))}`);
  if (urg.length) parts.push(`Atenciones no programadas:\n${L(urg.slice(0, 5))}`);
  if (docs.length) parts.push(docs.join('\n'));
  if (crit.length) parts.push(`Pendientes críticos:\n${L(crit.slice(0, 6))}`);
  if (risk.length) parts.push(`Riesgos próximos:\n${risk.slice(0, 5).map((t) => `- ${t.title}${t.due_date ? ` (${relDue(t.due_date)})` : ''}`).join('\n')}`);
  if (blocked && blocked.trim()) parts.push(`Bloqueos:\n${blocked.trim().split('\n').map((x) => '- ' + x.replace(/^[-•]\s*/, '')).join('\n')}`);
  if (decisions.length) parts.push(`Decisiones o información requerida:\n${L(decisions.slice(0, 6))}`);
  if (deleg.length) parts.push(`Propuesta de delegación (sujeta a su autorización):\n${deleg.map((t) => `- ${t.title}${t.delegate_to ? ` → ${t.delegate_to}` : ''}`).join('\n')}`);
  if (parts.length === 1) parts.push('Sin novedades relevantes.');
  return parts.join('\n\n');
}

// Resumen de un periodo (hoy o semana) para la Directora
export function periodSummary(period, ctx) {
  const { store, S } = ctx;
  const ref = today();
  const from = period === 'hoy' ? ref : addDays(ref, -6);
  const inRange = (iso) => iso && iso.slice(0, 10) >= from && iso.slice(0, 10) <= ref;
  const tasks = store.all('tasks');
  const done = tasks.filter((t) => t.status === 'completada' && inRange(t.completed_at)).sort(byScore);
  const pend = tasks.filter((t) => isOpen(t) && t.source !== 'bloque_diario').sort(byScore);
  const urg = tasks.filter((t) => t.is_urgent && inRange(t.created_at));
  const risk = riesgos(ctx);
  const blocked = tasks.filter((t) => isOpen(t) && t.status === 'bloqueada').map((t) => `${t.title}${t.blocked_reason ? ': ' + t.blocked_reason : ''}`).join('\n');
  const decisions = tasks.filter((t) => isOpen(t) && t.clasificacion === 'espera').map((t) => t.title + (t.waiting_on ? ` (de: ${t.waiting_on})` : ''));
  const deleg = tasks.filter((t) => isOpen(t) && t.clasificacion === 'delegable' && t.delegation_status !== 'autorizada');
  const intr = store.all('intramed').filter((i) => i.status !== 'cerrado');
  let msg = bossMessage({ docs: docLines(ctx, from, ref), done, pend, urg, risk, blocked, decisions, deleg, label: S.directora_label, period });
  if (intr.length) {
    const rep = intr.filter((i) => i.tipo === 'reporte').length, nec = intr.filter((i) => i.tipo === 'necesidad').length, sop = intr.filter((i) => i.tipo === 'soporte').length;
    const late = intr.filter((i) => i.due_date && i.due_date < ref).length;
    msg += `\n\nIntramed: ${intr.length} asunto(s) abiertos (${[rep && `${rep} reporte(s)`, nec && `${nec} necesidad(es) nueva(s)`, sop && `${sop} de soporte`].filter(Boolean).join(', ')})${late ? `, ${late} con compromiso vencido` : ''}.`;
  }
  return msg;
}

export { toMin };
