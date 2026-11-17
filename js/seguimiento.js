// Seguimientos (macro actividades con etapas, actividades y tiempo diario reservado)
// y Avance documental (documentos y formatos con flujo de revisiones).
import { today, addDays, fmtShort, fmtLong, fmtDur, toMin, toHHMM, esc, $, $$, diffDays, dow, uid, DIAS_C } from './util.js';
import { dayInfo, autoPlan, applyPlan, isOpen, byScore, CATEGORIAS, nextWorkingDay, settingsOf } from './planner.js';
import { FLOWS, CANCHA, CLASES, flowOf, stageOf, stageIndex, isDocDone, canchaOf, progress, nextStage, backStage, guessStage } from './flows.js';
import { readXlsx, readCsv, toCsv, toDate } from './xlsx.js';

let U = null; // utilidades de la interfaz (las pasa app.js)
const st = { projTab: 'activo', docFilter: 'proceso', docFlow: '', docStage: '', docQ: '', imp: null };
const PRIO = { 1: 'Crítica', 2: 'Alta', 3: 'Media', 4: 'Baja' };
const EST_P = { activo: 'Activo', pausa: 'En pausa', terminado: 'Terminado' };
const TPL = {
  blanco: { label: 'En blanco' },
  ece: {
    label: 'Formatos de expediente clínico electrónico',
    name: 'Formatos de expediente clínico electrónico', category: 'Formatos', priority: 2, doc_flow: 'formato_ece',
    description: 'Elaboro el formato → el consultor revisa contra la norma de expediente clínico y deja observaciones en la base compartida → corrijo → al terminar se envía al proveedor de ECE y a Compras (formatos físicos) → carga en Intramed.',
    daily: { min: 60, days: [1, 2, 3, 4, 5], start: '', until: '' },
  },
  docs: { label: 'Paquete de documentos', category: 'Documentos', priority: 2, doc_flow: 'documento_nuevo', daily: { min: 0, days: [1, 2, 3, 4, 5], start: '', until: '' } },
};

const ctx = () => U.ctx();
const store = () => U.store;
const nowISO = () => new Date().toISOString();
const daysIn = (doc) => (doc.stage_since ? Math.max(0, diffDays(doc.stage_since.slice(0, 10), today())) : 0);
const docLabel = (d) => `${d.code ? d.code + ' · ' : ''}${d.name}`;
const followDays = () => settingsOf(store()).followup_days || 7;

export function initSeguimiento(ui) {
  U = ui;
  return { views: { seguimientos: viewSeguimientos, documentos: viewDocumentos }, actions, forms, onChange, onTaskDone, ensureDailyBlocks, migrate, hoyCard, projectOptions, docFromTask };
}

// ---------------- Migración de datos existentes ----------------
export function migrate() {
  const S = store();
  for (const i of S.all('intramed')) {
    if (i.tipo === 'incidencia' || i.tipo === 'requerimiento' || !i.origen) {
      S.put('intramed', { ...i, tipo: i.tipo === 'requerimiento' ? 'necesidad' : i.tipo === 'incidencia' ? 'reporte' : (i.tipo || 'reporte'), origen: i.origen || (i.hospital ? 'hospital' : 'calidad') }, { silent: true });
    }
  }
}

// ---------------- Tiempo diario reservado ----------------
// Crea (una sola vez por día) la tarea "tiempo reservado" de cada seguimiento activo, para hoy y los próximos 6 días.
export function ensureDailyBlocks() {
  const S = store(); const c = ctx(); const t0 = today();
  let changed = 0;
  // Bloques de días pasados que no se usaron: se cierran sin arrastrarlos
  for (const t of S.all('tasks')) {
    if (t.source === 'bloque_diario' && t.scheduled_date && t.scheduled_date < t0 && isOpen(t) && t.status !== 'en_curso') {
      S.put('tasks', { ...t, status: 'cancelada', notes: (t.notes || '') + '\nBloque no utilizado.' }, { silent: true }); changed++;
    }
  }
  for (const p of S.all('projects')) {
    const d = p.daily || {};
    const active = p.status === 'activo' && +d.min > 0;
    const v = p.daily_v || 0;
    // Bloques futuros de una configuración anterior
    for (const t of S.all('tasks').filter((x) => x.project_id === p.id && x.source === 'bloque_diario' && (x.block_v || 0) !== v && isOpen(x) && x.status !== 'en_curso' && (x.scheduled_date || '') >= t0)) { S.remove('tasks', t.id); changed++; }
    for (let i = 0; i < 7; i++) {
      const day = addDays(t0, i);
      const id = `blk-${p.id}-${day}${v ? '-v' + v : ''}`;
      const existing = S.raw('tasks', id);
      const wants = active && (d.days || [1, 2, 3, 4, 5]).includes(dow(day)) && (!d.until || day <= d.until) && dayInfo(day, c).working && dayInfo(day, c).D?.tipo !== 'visita';
      if (!wants) {
        // Si se pausó o terminó el seguimiento, se quitan los bloques futuros no iniciados
        if (existing && !existing.deleted && isOpen(existing) && existing.status !== 'en_curso' && (!active || day > (d.until || '9999'))) { S.remove('tasks', id); changed++; }
        continue;
      }
      if (existing) continue; // ya existe (o el usuario lo eliminó: se respeta)
      const task = {
        id, title: `${p.name} · tiempo reservado`, project_id: p.id, source: 'bloque_diario', block_v: v, duration_min: +d.min,
        priority: p.priority || 2, clasificacion: 'programable', status: 'pendiente', category: p.category || 'General',
        next_action: p.next_action || '', notes: 'Bloque diario reservado hasta terminar el seguimiento.',
      };
      if (d.start) { task.scheduled_date = day; task.scheduled_start = d.start; task.status = 'programada'; S.put('tasks', task, { silent: true }); }
      else {
        const saved = S.put('tasks', task, { silent: true });
        const plan = autoPlan([saved], ctx(), { from: day, days: 1, useBuffer: false }).filter((x) => !x.unplaced && x.date === day);
        if (plan.length === 1) applyPlan(plan, S);
        else S.put('tasks', { ...saved, scheduled_date: day, scheduled_start: null }, { silent: true });
      }
      changed++;
    }
  }
  return changed;
}

// ---------------- Avance de documentos ----------------
function moveDoc(doc, stageId, note = '') {
  const S = store();
  const prev = stageOf(doc);
  const to = flowOf(doc).stages.find((s) => s.id === stageId);
  if (!to || to.id === prev.id) return doc;
  let round = doc.round || 0;
  if (to.round) round = prev.next === to.id || prev.id === to.back ? round + 1 : Math.max(1, round || 1);
  const log = [...(doc.log || []), { at: nowISO(), text: `Etapa: ${prev.label} → ${to.label}${to.round ? ` (ronda ${round})` : ''}${note ? '. ' + note : ''}` }];
  const upd = { ...doc, stage: to.id, stage_since: nowISO(), round, log, completed_at: to.done ? nowISO() : null };
  S.put('docs', upd);
  // La tarea de la etapa anterior (si era mía) se da por terminada al avanzar
  for (const t of S.all('tasks').filter((x) => x.doc_id === doc.id && x.doc_stage === prev.id && isOpen(x))) {
    S.put('tasks', { ...t, status: 'completada', completed_at: nowISO(), notes: (t.notes || '') + `\nCerrada al pasar el documento a “${to.label}”.` });
  }
  maybeCreateDocTask(upd);
  return upd;
}

function maybeCreateDocTask(doc) {
  const S = store(); const stg = stageOf(doc);
  if (stg.done || canchaOf(doc) !== 'calidad' || !stg.task) return null;
  const proj = doc.project_id ? S.get('projects', doc.project_id) : null;
  if (proj && proj.status === 'activo' && +proj.daily?.min > 0) return null; // se trabaja en el tiempo reservado del seguimiento
  if (S.all('tasks').some((t) => t.doc_id === doc.id && t.doc_stage === stg.id && isOpen(t))) return null;
  return S.put('tasks', {
    title: `${stg.task}${stg.round && doc.round > 1 ? ` (ronda ${doc.round})` : ''}: ${docLabel(doc)}`,
    doc_id: doc.id, doc_stage: stg.id, project_id: doc.project_id || null, duration_min: stg.min || 60,
    category: doc.clase === 'Formato' || doc.flow === 'formato_ece' ? 'Formatos' : 'Documentos',
    priority: doc.priority || 3, clasificacion: 'programable', status: 'pendiente', due_date: doc.due_date || null,
    area: doc.area || '', next_action: stg.task, source: 'documento',
  });
}

// Llamada por app.js cuando se completa una tarea
export function onTaskDone(t) {
  const S = store();
  if (t.doc_id) {
    const d = S.get('docs', t.doc_id);
    if (d && d.stage === t.doc_stage) {
      const nx = nextStage(d);
      if (nx) return { label: `Pasar a “${nx.label}”`, fn: () => { moveDoc(d, nx.id); U.toast('Documento: ' + nx.label); } };
    }
  }
  if (t.project_id) return { label: 'Anotar avance', fn: () => openNote(t.project_id) };
  return null;
}

// Crea un documento a partir de una tarea capturada en texto libre y los vincula
function docFromTask(task, text) {
  const t = String(text || '').toLowerCase();
  const isFormato = task.category === 'Formatos' || /formato/.test(t);
  const flow = /actualiz/.test(t) ? 'documento_actualizacion' : (isFormato && /expediente|ece|electronico|electrónico/.test(t) ? 'formato_ece' : 'documento_nuevo');
  const stage = guessStage(flow, text) || FLOWS[flow].stages[0].id;
  const name = task.title.replace(/^(revisar|actualizar|validar|emitir|elaborar|corregir|cargar|subir)\s+/i, '').replace(/ · .*$/, '');
  const doc = store().put('docs', { name: name.charAt(0).toUpperCase() + name.slice(1), flow, stage, stage_since: nowISO(), round: FLOWS[flow].stages.find((x) => x.id === stage)?.round ? 1 : 0, clase: isFormato ? 'Formato' : '', area: task.area || '', due_date: task.due_date || null, priority: task.priority || 3, log: [{ at: nowISO(), text: 'Registrado desde captura rápida' }] });
  store().put('tasks', { ...store().get('tasks', task.id), doc_id: doc.id, doc_stage: stage });
  return doc;
}

// ---------------- Vista: Seguimientos ----------------
function projStats(p) {
  const S = store();
  const docs = S.all('docs').filter((d) => d.project_id === p.id);
  const tasks = S.all('tasks').filter((t) => t.project_id === p.id && t.source !== 'bloque_diario');
  const et = p.etapas || [];
  const docsDone = docs.filter(isDocDone).length;
  const total = docs.length + et.length + tasks.length;
  const done = docsDone + et.filter((e) => e.done).length + tasks.filter((t) => t.status === 'completada').length;
  return { docs, tasks, et, docsDone, pct: total ? Math.round((done / total) * 100) : 0, total, done,
    mine: docs.filter((d) => !isDocDone(d) && canchaOf(d) === 'calidad').length,
    waiting: docs.filter((d) => !isDocDone(d) && canchaOf(d) !== 'calidad').length,
    openTasks: tasks.filter(isOpen).length };
}
const dailyText = (p) => (+p.daily?.min > 0 ? `${fmtDur(+p.daily.min)} diarios ${(p.daily.days || []).map((d) => DIAS_C[d]).join(' ')}${p.daily.start ? ' a las ' + p.daily.start : ''}${p.daily.until ? ' hasta ' + fmtShort(p.daily.until) : ' hasta terminar'}` : '');

function viewSeguimientos() {
  const all = store().all('projects');
  const list = all.filter((p) => (p.status || 'activo') === st.projTab).sort((a, b) => (a.priority || 3) - (b.priority || 3) || (a.due_date || '9').localeCompare(b.due_date || '9'));
  const n = (k) => all.filter((p) => (p.status || 'activo') === k).length;
  return `
  <section class="day-head"><div class="datenav-t"><h1>Seguimientos</h1><div class="muted">Actividades grandes con etapas, actividades diarias y tiempo reservado. Las tareas sueltas siguen en Tareas.</div></div>
  <div class="btnrow"><button class="btn primary small" data-act="proj-new">+ Nuevo seguimiento</button></div></section>
  <div class="tabs">${Object.entries(EST_P).map(([k, l]) => `<button class="${st.projTab === k ? 'on' : ''}" data-act="proj-tab" data-t="${k}">${l} (${n(k)})</button>`).join('')}</div>
  ${list.length ? `<div class="pgrid">${list.map((p) => {
    const s = projStats(p);
    return `<div class="card pcard" data-act="proj" data-id="${p.id}">
      <div class="trow-title"><span class="pdot p${p.priority || 3}"></span>${esc(p.name)}</div>
      <div class="chips">${dailyText(p) ? U.chip('⏱ ' + esc(dailyText(p)), 'info') : ''}${p.due_date ? U.chip('Objetivo ' + esc(fmtShort(p.due_date)), p.due_date < today() ? 'danger' : 'soft') : ''}${p.category ? U.chip(esc(p.category), 'soft') : ''}</div>
      <div class="pbar"><i style="width:${s.pct}%"></i></div>
      <div class="small muted">${s.pct}% · ${s.done} de ${s.total} elementos${s.docs.length ? ` · ${s.docsDone}/${s.docs.length} documentos terminados` : ''}</div>
      ${s.docs.length ? `<div class="small">En mi cancha: <b>${s.mine}</b> · Esperando a otros: <b>${s.waiting}</b></div>` : ''}
      ${s.openTasks ? `<div class="small">Actividades abiertas: <b>${s.openTasks}</b></div>` : ''}
      ${p.next_action ? `<div class="small">Siguiente: ${esc(p.next_action)}</div>` : ''}
    </div>`;
  }).join('')}</div>` : `<div class="card empty"><p class="muted">${st.projTab === 'activo' ? 'Sin seguimientos activos. Crea uno para trabajos con varias etapas (p. ej., formatos de expediente clínico) y, si quieres, resérvale tiempo diario.' : 'Nada aquí.'}</p></div>`}`;
}

function openProject(id, tplKey = 'blanco', keep = false) {
  const S = store();
  const base = TPL[tplKey] || TPL.blanco;
  const p = id ? S.get('projects', id) : { status: 'activo', priority: base.priority || 3, category: base.category || 'General', name: base.name || '', description: base.description || '', doc_flow: base.doc_flow || '', daily: base.daily || { min: 0, days: [1, 2, 3, 4, 5], start: '', until: '' }, etapas: [], log: [] };
  if (!p) return;
  const d = p.daily || { min: 0, days: [1, 2, 3, 4, 5] };
  const s = id ? projStats(p) : null;
  const projTasks = id ? s.tasks.filter(isOpen).sort(byScore) : [];
  const doneTasks = id ? s.tasks.filter((t) => t.status === 'completada').length : 0;
  U.openSheet(id ? 'Seguimiento' : 'Nuevo seguimiento', `
  ${!id ? `<div class="btnrow wrap"><span class="small muted">Plantilla:</span>${Object.entries(TPL).map(([k, t]) => `<button class="btn small ${k === tplKey ? 'on' : ''}" type="button" data-act="proj-tpl" data-t="${k}">${esc(t.label)}</button>`).join('')}</div>` : ''}
  <form data-form="project" data-id="${id || ''}" class="stack">
    <input name="name" class="title-input" required value="${esc(p.name || '')}" placeholder="Nombre del seguimiento" ${id ? '' : 'autofocus'}>
    <label>Descripción / flujo de trabajo<textarea name="description" rows="3">${esc(p.description || '')}</textarea></label>
    <div class="grid2">
      <label>Estado<select name="status">${Object.entries(EST_P).map(([k, l]) => `<option value="${k}" ${p.status === k ? 'selected' : ''}>${l}</option>`).join('')}</select></label>
      <label>Prioridad<select name="priority">${[1, 2, 3, 4].map((x) => `<option value="${x}" ${+p.priority === x ? 'selected' : ''}>P${x} · ${PRIO[x]}</option>`).join('')}</select></label>
      <label>Categoría<select name="category">${CATEGORIAS.map((x) => `<option ${p.category === x ? 'selected' : ''}>${x}</option>`).join('')}</select></label>
      <label>Fecha objetivo<input type="date" name="due_date" value="${p.due_date || ''}"></label>
    </div>
    <label>Siguiente acción<input name="next_action" value="${esc(p.next_action || '')}"></label>
    <div class="subcard">
      <div class="lbl">Tiempo diario reservado (hasta terminar)</div>
      <div class="seg small">${[0, 30, 45, 60, 90, 120].map((m) => `<label><input type="radio" name="daily_min" value="${m}" ${+d.min === m ? 'checked' : ''}><span>${m ? fmtDur(m) : 'No reservar'}</span></label>`).join('')}</div>
      <div class="days">${[1, 2, 3, 4, 5, 6].map((x) => `<label class="dchk"><input type="checkbox" name="daily_days" value="${x}" ${(d.days || []).includes(x) ? 'checked' : ''}>${DIAS_C[x]}</label>`).join('')}</div>
      <div class="grid2"><label>Hora preferida (opcional)<input type="time" step="900" name="daily_start" value="${d.start || ''}"></label><label>Hasta (vacío = hasta terminar)<input type="date" name="daily_until" value="${d.until || ''}"></label></div>
      <p class="small muted">Se agenda solo cada día como “${esc(p.name || 'Seguimiento')} · tiempo reservado”. Si no lo usas, se cierra al día siguiente sin arrastrarse.</p>
    </div>
    <label>Flujo para documentos/formatos de este seguimiento<select name="doc_flow"><option value="">— Sin documentos —</option>${Object.entries(FLOWS).map(([k, f]) => `<option value="${k}" ${p.doc_flow === k ? 'selected' : ''}>${esc(f.label)}</option>`).join('')}</select></label>
    <div class="btnrow savebar"><button class="btn primary grow">${id ? 'Guardar' : 'Crear seguimiento'}</button>${id ? `<button class="btn ghost danger-t" type="button" data-act="proj-del" data-id="${id}">Eliminar</button>` : ''}</div>
  </form>
  ${id ? `
  <h3>Etapas (${(p.etapas || []).filter((e) => e.done).length}/${(p.etapas || []).length})</h3>
  ${(p.etapas || []).map((e) => `<div class="etapa ${e.done ? 'done' : ''}"><label class="cb"><input type="checkbox" data-act="proj-etapa" data-id="${id}" data-e="${e.id}" ${e.done ? 'checked' : ''}> <span>${esc(e.title)}</span></label><button class="icon-btn sm" type="button" data-act="proj-etapa-del" data-id="${id}" data-e="${e.id}" aria-label="Quitar">✕</button></div>`).join('')}
  <form data-form="proj-etapa" data-id="${id}" class="inline-add"><input name="title" placeholder="Agregar etapa (p. ej., Validación final con la Directora)"><button class="btn small">Agregar</button></form>

  <h3>Actividades (${projTasks.length} abiertas · ${doneTasks} completadas)</h3>
  ${projTasks.map((t) => U.taskRow(t)).join('') || '<p class="small muted">Sin actividades abiertas.</p>'}
  <form data-form="proj-act" data-id="${id}" class="inline-add"><input name="title" placeholder="Nueva actividad (micro actividad)"><select name="dur">${[15, 30, 45, 60, 90].map((m) => `<option value="${m}" ${m === 30 ? 'selected' : ''}>${fmtDur(m)}</option>`).join('')}</select><button class="btn small">Agregar</button></form>

  ${p.doc_flow || s.docs.length ? `<h3>Documentos / formatos (${s.docsDone}/${s.docs.length} terminados)</h3>
  ${s.docs.sort((a, b) => stageIndex(a) - stageIndex(b)).map((dd) => docRowMini(dd)).join('') || '<p class="small muted">Aún no hay documentos.</p>'}
  <div class="btnrow"><button class="btn small" type="button" data-act="doc-new" data-p="${id}">+ Agregar documento o formato</button></div>` : ''}

  <h3>Bitácora</h3>
  <form data-form="proj-note" data-id="${id}" class="inline-add"><input name="text" placeholder="¿Qué avanzaste o qué pasó?"><button class="btn small">Anotar</button></form>
  <div class="log">${[...(p.log || [])].reverse().map((l) => `<div><span class="small muted">${esc(new Date(l.at).toLocaleString('es-MX', { dateStyle: 'short', timeStyle: 'short' }))}</span> ${esc(l.text)}</div>`).join('') || '<p class="small muted">Sin notas.</p>'}</div>` : ''}
  `, { wide: true, keep });
}
function docRowMini(d) {
  const c = canchaOf(d);
  return `<div class="irow" data-act="doc" data-id="${d.id}"><div class="irow-main"><div class="trow-title">${esc(docLabel(d))}</div>
  <div class="chips">${U.chip(esc(stageOf(d).label), isDocDone(d) ? 'st-cerrado' : c === 'calidad' ? 'warn' : 'soft')}${d.round > 1 ? U.chip('Ronda ' + d.round, 'soft') : ''}${!isDocDone(d) ? U.chip(`${esc(CANCHA[c])} · ${daysIn(d)} d`, c !== 'calidad' && daysIn(d) > followDays() ? 'danger' : 'soft') : ''}</div></div></div>`;
}
function openNote(projId) {
  const p = store().get('projects', projId); if (!p) return;
  U.openSheet('Anotar avance', `<form data-form="proj-note" data-id="${projId}" data-close="1" class="stack"><p class="small muted">${esc(p.name)}</p><textarea name="text" rows="3" autofocus placeholder="Ej.: Corregí 3 formatos con observaciones del consultor"></textarea><button class="btn primary block">Guardar en bitácora</button></form>`);
}
export function projectOptions(sel) {
  return `<option value="">— Ninguno (tarea aislada) —</option>` + store().all('projects').filter((p) => p.status !== 'terminado' || p.id === sel).map((p) => `<option value="${p.id}" ${p.id === sel ? 'selected' : ''}>${esc(p.name)}</option>`).join('');
}

// ---------------- Vista: Avance documental ----------------
function docFilters() {
  const ref = today();
  return {
    proceso: ['En proceso', (d) => !isDocDone(d)],
    mia: ['En mi cancha', (d) => !isDocDone(d) && canchaOf(d) === 'calidad'],
    otros: ['Esperando a otros', (d) => !isDocDone(d) && canchaOf(d) !== 'calidad'],
    atrasados: [`Esperando > ${followDays()} días`, (d) => !isDocDone(d) && canchaOf(d) !== 'calidad' && daysIn(d) > followDays()],
    vencidos: ['Compromiso vencido', (d) => !isDocDone(d) && d.due_date && d.due_date < ref],
    intramed: ['Por cargar en Intramed', (d) => d.stage === 'intramed'],
    terminados: ['Terminados', (d) => isDocDone(d)],
    todos: ['Todos', () => true],
  };
}
function viewDocumentos() {
  const S = store(); const all = S.all('docs');
  const F = docFilters();
  const fl = F[st.docFilter] || F.proceso;
  let list = all.filter(fl[1]);
  if (st.docFlow) list = list.filter((d) => d.flow === st.docFlow);
  if (st.docStage) list = list.filter((d) => d.stage === st.docStage);
  const q = st.docQ.trim().toLowerCase();
  if (q) list = list.filter((d) => [d.code, d.name, d.owner, d.area, d.clase].join(' ').toLowerCase().includes(q));
  list.sort((a, b) => (canchaOf(a) === 'calidad' ? 0 : 1) - (canchaOf(b) === 'calidad' ? 0 : 1) || (a.due_date || '9').localeCompare(b.due_date || '9') || daysIn(b) - daysIn(a));
  const month = today().slice(0, 7);
  const tiles = [
    ['proceso', 'En proceso', all.filter(F.proceso[1]).length, ''],
    ['mia', 'En mi cancha', all.filter(F.mia[1]).length, 'warn'],
    ['otros', 'Esperando a otros', all.filter(F.otros[1]).length, ''],
    ['atrasados', `Esperando > ${followDays()} d`, all.filter(F.atrasados[1]).length, 'danger'],
    ['intramed', 'Por cargar en Intramed', all.filter(F.intramed[1]).length, 'info'],
    ['terminados', 'Terminados este mes', all.filter((d) => isDocDone(d) && (d.completed_at || '').slice(0, 7) === month).length, ''],
  ];
  const flowsPresent = Object.keys(FLOWS).filter((k) => all.some((d) => d.flow === k && !isDocDone(d)));
  const pipeline = flowsPresent.map((k) => {
    const docs = all.filter((d) => d.flow === k);
    const stages = FLOWS[k].stages.filter((s2) => !s2.done);
    const max = Math.max(1, ...stages.map((s2) => docs.filter((d) => d.stage === s2.id).length));
    return `<div class="pipe"><div class="pipe-h">${esc(FLOWS[k].label)} <span class="muted small">· ${docs.filter((d) => !isDocDone(d)).length} en proceso, ${docs.filter(isDocDone).length} terminados</span></div>
      ${stages.map((s2) => { const n = docs.filter((d) => d.stage === s2.id).length; return `<button class="pipe-row ${st.docStage === s2.id && st.docFlow === k ? 'on' : ''}" data-act="doc-stage" data-f="${k}" data-s="${s2.id}"><span class="pipe-l">${esc(s2.label)}</span><span class="pipe-bar"><i class="${s2.cancha === 'calidad' ? 'mine' : ''}" style="width:${(n / max) * 100}%"></i></span><b>${n}</b></button>`; }).join('')}
    </div>`;
  }).join('');
  return `
  <section class="day-head"><div class="datenav-t"><h1>Avance documental</h1><div class="muted">Documentos nuevos, actualizaciones y formatos: etapa, ronda de revisión y en qué cancha está</div></div>
  <div class="btnrow"><button class="btn primary small" data-act="doc-new">+ Documento o formato</button><button class="btn small" data-act="doc-import">Importar Excel / CSV</button><button class="btn small" data-act="doc-export">Exportar CSV</button></div></section>
  <div class="tiles">${tiles.map(([k, l, n, cls]) => `<button class="alert-btn ${n ? cls : 'zero'} ${st.docFilter === k ? 'sel' : ''}" data-act="doc-filter" data-f="${k}"><b>${n}</b><span>${l}</span></button>`).join('')}</div>
  ${all.length ? `<div class="doc-grid">
    <div class="card">${pipeline || '<p class="small muted">Sin documentos en proceso.</p>'}<p class="small muted">Barras verdes: etapas en tu cancha. Toca una etapa para filtrar.</p></div>
    <div>
      <input class="search" type="search" placeholder="Buscar código, nombre, responsable…" value="${esc(st.docQ)}" data-input="docq">
      <div class="filters"><button class="fchip ${!st.docFlow ? 'on' : ''}" data-act="doc-flow" data-f="">Todos los flujos</button>${Object.entries(FLOWS).map(([k, f]) => `<button class="fchip ${st.docFlow === k ? 'on' : ''}" data-act="doc-flow" data-f="${k}">${esc(f.label)}</button>`).join('')}${st.docStage ? `<button class="fchip on" data-act="doc-stage" data-f="" data-s="">Etapa: ${esc((FLOWS[st.docFlow]?.stages.find((x) => x.id === st.docStage) || {}).label || '')} ✕</button>` : ''}</div>
      <div class="card list"><div class="small muted">${esc(fl[0])}: ${list.length}</div>${list.map(docRow).join('') || '<p class="muted">Sin documentos en este filtro.</p>'}</div>
    </div>
  </div>` : `<div class="card empty"><h3>Aún no hay documentos</h3><p class="muted">Agrega uno o importa tu base de Excel; la app reconoce columnas como código, nombre, tipo, responsable, estatus y fechas.</p><div class="btnrow" style="justify-content:center"><button class="btn primary" data-act="doc-import">Importar Excel / CSV</button><button class="btn" data-act="doc-new">+ Documento o formato</button></div></div>`}`;
}
function docRow(d) {
  const c = canchaOf(d); const days = daysIn(d); const late = d.due_date && d.due_date < today() && !isDocDone(d);
  const proj = d.project_id ? store().get('projects', d.project_id) : null;
  return `<div class="drow" data-act="doc" data-id="${d.id}">
    <div class="drow-main">
      <div class="trow-title">${d.code ? `<span class="muted">${esc(d.code)}</span> ` : ''}${esc(d.name)}</div>
      <div class="chips">${U.chip(esc(FLOWS[d.flow]?.label || ''), 'soft')}${d.clase ? U.chip(esc(d.clase), 'soft') : ''}${d.owner ? U.chip('Resp.: ' + esc(d.owner), 'soft') : ''}${proj ? U.chip(esc(proj.name), 'info') : ''}${d.due_date ? U.chip('Compromiso ' + esc(fmtShort(d.due_date)), late ? 'danger' : 'soft') : ''}${d.excel_status && d.stage_unsure ? U.chip('Etapa por confirmar: ' + esc(d.excel_status), 'warn') : ''}</div>
    </div>
    <div class="drow-stage">
      <div class="small"><b>${esc(stageOf(d).label)}</b>${d.round > 1 ? ` · ronda ${d.round}` : ''}</div>
      <div class="pbar sm"><i style="width:${progress(d)}%"></i></div>
      ${isDocDone(d) ? `<div class="small muted">Terminado ${d.completed_at ? esc(fmtShort(d.completed_at.slice(0, 10))) : ''}</div>` : `<div class="small ${c === 'calidad' ? 'warn-t' : days > followDays() ? 'danger-t' : 'muted'}">${c === 'calidad' ? 'En mi cancha' : esc(CANCHA[c])} · ${days} d</div>`}
    </div>
  </div>`;
}

function openDoc(id, draft = {}) {
  const S = store();
  const d = id ? S.get('docs', id) : { flow: draft.flow || 'documento_nuevo', priority: 3, round: 0, log: [], start_date: today(), ...draft };
  if (!d) return;
  if (!d.stage) d.stage = FLOWS[d.flow].stages[0].id;
  const fl = flowOf(d); const cur = stageOf(d); const nx = nextStage(d); const bk = backStage(d);
  const tasks = id ? S.all('tasks').filter((t) => t.doc_id === id) : [];
  U.openSheet(id ? 'Documento' : 'Nuevo documento o formato', `
  ${id ? `<div class="stepper">${fl.stages.map((s2, i) => `<button type="button" class="step ${i < stageIndex(d) ? 'past' : ''} ${s2.id === cur.id ? 'cur' : ''} ${s2.cancha === 'calidad' ? 'mine' : ''}" data-act="doc-goto" data-id="${id}" data-s="${s2.id}" title="Mover a esta etapa"><i></i><span>${esc(s2.label)}</span><small>${esc(CANCHA[s2.cancha] || '')}</small></button>`).join('')}</div>
  <div class="alert ${canchaOf(d) === 'calidad' ? 'warn' : ''} small">Etapa actual: <b>${esc(cur.label)}</b>${d.round ? ` · ronda ${d.round}` : ''} · ${isDocDone(d) ? 'terminado' : `en cancha de <b>${esc(CANCHA[canchaOf(d)])}</b> desde hace ${daysIn(d)} día(s)`}</div>
  <div class="btnrow wrap">
    ${nx ? `<button class="btn primary" type="button" data-act="doc-next" data-id="${id}">Avanzar → ${esc(nx.label)}</button>` : ''}
    ${bk ? `<button class="btn" type="button" data-act="doc-back" data-id="${id}">Regresar con observaciones → ${esc(bk.label)}</button>` : ''}
    ${!isDocDone(d) ? `<button class="btn small" type="button" data-act="doc-task" data-id="${id}">Crear tarea</button>` : ''}
    ${!isDocDone(d) && canchaOf(d) !== 'calidad' ? `<button class="btn small" type="button" data-act="doc-remind" data-id="${id}">Recordatorio de seguimiento</button>` : ''}
  </div>
  <label class="small">Nota para el cambio de etapa (opcional)<input id="doc-move-note" placeholder="Ej.: 5 observaciones enviadas por correo"></label>` : ''}
  <form data-form="doc" data-id="${id || ''}" class="stack">
    <div class="grid2">
      <label>Código<input name="code" value="${esc(d.code || '')}" placeholder="CHSAI-…"></label>
      <label>Clase<select name="clase"><option value="">—</option>${CLASES.map((x) => `<option ${d.clase === x ? 'selected' : ''}>${x}</option>`).join('')}</select></label>
    </div>
    <label>Nombre<input name="name" required value="${esc(d.name || '')}" ${id ? '' : 'autofocus'}></label>
    <div class="grid2">
      <label>Flujo<select name="flow">${Object.entries(FLOWS).map(([k, f]) => `<option value="${k}" ${d.flow === k ? 'selected' : ''}>${esc(f.label)}</option>`).join('')}</select></label>
      ${!id ? `<label>Etapa inicial<select name="stage">${fl.stages.map((s2) => `<option value="${s2.id}" ${d.stage === s2.id ? 'selected' : ''}>${esc(s2.label)}</option>`).join('')}</select></label>` : `<label>Ronda de revisión<input type="number" min="0" name="round" value="${d.round || 0}"></label>`}
      <label>Dueño de proceso / responsable<input name="owner" value="${esc(d.owner || '')}"></label>
      <label>Área / hospital<input name="area" list="dl-hosp" value="${esc(d.area || '')}"></label>
      <label>Fecha de inicio<input type="date" name="start_date" value="${d.start_date || ''}"></label>
      <label>Fecha compromiso<input type="date" name="due_date" value="${d.due_date || ''}"></label>
      <label>Versión<input name="version" value="${esc(d.version || '')}"></label>
      <label>Prioridad<select name="priority">${[1, 2, 3, 4].map((x) => `<option value="${x}" ${+d.priority === x ? 'selected' : ''}>P${x} · ${PRIO[x]}</option>`).join('')}</select></label>
    </div>
    <label>Seguimiento al que pertenece<select name="project_id">${projectOptions(d.project_id)}</select></label>
    <label>Enlace al archivo (Teams / SharePoint)<input name="link" type="url" value="${esc(d.link || '')}" placeholder="https://…"></label>
    ${d.link ? `<a class="small" href="${esc(d.link)}" target="_blank" rel="noopener">Abrir archivo</a>` : ''}
    <label>Notas<textarea name="notes" rows="2">${esc(d.notes || '')}</textarea></label>
    ${id ? '<label>Agregar a la bitácora<input name="newlog" placeholder="Qué pasó con el documento"></label>' : ''}
    <div class="btnrow savebar"><button class="btn primary grow">${id ? 'Guardar' : 'Crear'}</button>${id ? `<button class="btn ghost danger-t" type="button" data-act="doc-del" data-id="${id}">Eliminar</button>` : ''}</div>
  </form>
  ${tasks.length ? `<h3>Tareas vinculadas</h3>${tasks.sort((a, b) => (isOpen(b) ? 1 : 0) - (isOpen(a) ? 1 : 0)).slice(0, 10).map((t) => U.taskRow(t)).join('')}` : ''}
  ${(d.log || []).length ? `<h3>Historial</h3><div class="log">${[...d.log].reverse().map((l) => `<div><span class="small muted">${esc(new Date(l.at).toLocaleString('es-MX', { dateStyle: 'short', timeStyle: 'short' }))}</span> ${esc(l.text)}</div>`).join('')}</div>` : ''}
  ${U.hospitalsDatalist()}`, { wide: true });
}

// ---------------- Importar Excel / CSV ----------------
const FIELDS = [
  ['code', 'Código', /codigo|clave|folio|^id$/],
  ['name', 'Nombre del documento', /nombre|titulo|documento|formato|descripcion/],
  ['clase', 'Clase (procedimiento, formato…)', /clase|tipo de documento|categoria|^tipo doc/],
  ['tipo', 'Nuevo / actualización', /nuevo|actualiz|movimiento|tipo$|^tipo/],
  ['owner', 'Dueño de proceso / responsable', /dueno|responsable|elabor|propietario|owner/],
  ['area', 'Área / hospital', /area|hospital|unidad|servicio|departamento/],
  ['status', 'Estatus / etapa', /estatus|status|estado|etapa|avance|situacion/],
  ['start_date', 'Fecha de inicio', /inicio|recepcion|recibido|solicitud|fecha de ingreso/],
  ['due_date', 'Fecha compromiso', /compromiso|limite|entrega|vencim|programad/],
  ['version', 'Versión', /version|rev/],
  ['link', 'Enlace', /liga|link|enlace|url|ruta/],
  ['notes', 'Observaciones / notas', /observ|nota|coment/],
];
function openImport() {
  st.imp = null;
  U.openSheet('Importar avance documental', `<div class="stack">
    <p class="small muted">Sube tu Excel (.xlsx) o un CSV. Después eliges qué columna corresponde a cada dato; la app intenta reconocerlas sola. Si el código ya existe en la app, se actualiza ese documento en lugar de duplicarlo.</p>
    <label class="btn primary">Elegir archivo<input type="file" accept=".xlsx,.csv,text/csv" data-docimport hidden></label>
    <div id="imp-body"></div></div>`, { wide: true });
}
async function loadImportFile(file) {
  const box = $('#imp-body'); box.innerHTML = '<p class="muted">Leyendo…</p>';
  try {
    const sheets = /\.csv$/i.test(file.name) ? await readCsv(file) : await readXlsx(file);
    st.imp = { file: file.name, sheets, sheet: 0 };
    pickSheet(0);
  } catch (e) { console.error(e); box.innerHTML = `<p class="alert danger">No se pudo leer el archivo: ${esc(e.message)}. Prueba guardarlo como .xlsx o CSV.</p>`; }
}
function pickSheet(i) {
  const imp = st.imp; imp.sheet = i;
  const rows = imp.sheets[i].rows.filter((r) => r && r.some((x) => String(x).trim() !== ''));
  // Encabezado: la primera fila con al menos 2 textos
  let h = rows.findIndex((r) => r.filter((x) => typeof x === 'string' && x.trim()).length >= 2);
  if (h < 0) h = 0;
  imp.headers = (rows[h] || []).map((x, j) => String(x || `Columna ${j + 1}`).trim());
  imp.data = rows.slice(h + 1);
  const n = (s) => s.toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
  imp.map = {};
  const used = new Set();
  for (const [k, , re] of FIELDS) {
    const j = imp.headers.findIndex((hd, idx) => !used.has(idx) && re.test(n(hd)));
    if (j >= 0) { imp.map[k] = j; used.add(j); }
  }
  imp.formatFlow = 'formato_ece';
  renderImport();
}
function importRows() {
  const imp = st.imp; const m = imp.map;
  const get = (r, k) => (m[k] === undefined || m[k] === '' ? '' : r[m[k]]);
  return imp.data.map((r) => {
    const name = String(get(r, 'name') || '').trim(); if (!name) return null;
    const clase = String(get(r, 'clase') || '').trim();
    const tipo = String(get(r, 'tipo') || '').toLowerCase();
    const isFormato = /formato/i.test(clase) || /formato/i.test(name);
    let flow = /actualiz|modific|revision/.test(tipo) ? 'documento_actualizacion' : 'documento_nuevo';
    if (isFormato && imp.formatFlow === 'formato_ece') flow = 'formato_ece';
    const status = String(get(r, 'status') || '').trim();
    const stage = guessStage(flow, status);
    const claseNorm = CLASES.find((c) => c.toLowerCase() === clase.toLowerCase()) || (isFormato ? 'Formato' : clase ? 'Otro' : '');
    return {
      code: String(get(r, 'code') || '').trim(), name, clase: claseNorm, flow,
      owner: String(get(r, 'owner') || '').trim(), area: String(get(r, 'area') || '').trim(),
      stage: stage || FLOWS[flow].stages[0].id, stage_unsure: !stage && !!status, excel_status: status,
      start_date: toDate(get(r, 'start_date')), due_date: toDate(get(r, 'due_date')),
      version: String(get(r, 'version') || '').trim(), link: String(get(r, 'link') || '').trim(), notes: String(get(r, 'notes') || '').trim(),
    };
  }).filter(Boolean);
}
function renderImport() {
  const imp = st.imp; const box = $('#imp-body'); if (!box || !imp) return;
  const rows = importRows();
  const existing = new Map(store().all('docs').filter((d) => d.code).map((d) => [d.code.toLowerCase(), d]));
  const upd = rows.filter((r) => r.code && existing.has(r.code.toLowerCase())).length;
  const unsure = rows.filter((r) => r.stage_unsure).length;
  box.innerHTML = `
    ${imp.sheets.length > 1 ? `<label>Hoja<select data-impsheet>${imp.sheets.map((s, i) => `<option value="${i}" ${i === imp.sheet ? 'selected' : ''}>${esc(s.name)}</option>`).join('')}</select></label>` : ''}
    <div class="grid2">${FIELDS.map(([k, l]) => `<label>${esc(l)}<select data-impmap="${k}"><option value="">— No importar —</option>${imp.headers.map((hd, j) => `<option value="${j}" ${imp.map[k] === j ? 'selected' : ''}>${esc(hd)}</option>`).join('')}</select></label>`).join('')}</div>
    <label>Flujo para los formatos<select data-impformat><option value="formato_ece" ${imp.formatFlow === 'formato_ece' ? 'selected' : ''}>Formato de expediente clínico electrónico (consultor → proveedor → Compras)</option><option value="documento" ${imp.formatFlow !== 'formato_ece' ? 'selected' : ''}>Igual que documentos (revisión con dueño de proceso)</option></select></label>
    <div class="alert small">${rows.length} documento(s) a importar · ${upd} se actualizarán por código · ${rows.length - upd} nuevos${unsure ? ` · <b>${unsure}</b> con estatus no reconocido (quedan en la primera etapa con aviso “Etapa por confirmar”)` : ''}. Los días en etapa se cuentan desde la fecha de recepción hasta que muevas cada documento.</div>
    <div class="imp-prev">${rows.slice(0, 6).map((r) => `<div class="small">• ${esc(r.code ? r.code + ' · ' : '')}${esc(r.name)} → <b>${esc(FLOWS[r.flow].stages.find((s) => s.id === r.stage).label)}</b>${r.excel_status ? ` <span class="muted">(Excel: ${esc(r.excel_status)})</span>` : ''}</div>`).join('')}${rows.length > 6 ? `<div class="small muted">… y ${rows.length - 6} más</div>` : ''}</div>
    <button class="btn primary block" type="button" data-act="doc-import-run" ${rows.length && imp.map.name !== undefined ? '' : 'disabled'}>Importar ${rows.length} documento(s)</button>
    ${imp.map.name === undefined ? '<p class="small danger-t">Elige al menos la columna del nombre del documento.</p>' : ''}`;
}
function runImport() {
  const S = store(); const rows = importRows();
  const existing = new Map(S.all('docs').filter((d) => d.code).map((d) => [d.code.toLowerCase(), d]));
  let created = 0, updated = 0;
  for (const r of rows) {
    const ex = r.code ? existing.get(r.code.toLowerCase()) : null;
    const clean = Object.fromEntries(Object.entries(r).filter(([, v]) => v !== '' && v !== null));
    if (ex) {
      const stageChanged = r.excel_status && !r.stage_unsure && ex.stage !== r.stage && ex.flow === r.flow;
      S.put('docs', { ...ex, ...clean, stage: stageChanged ? r.stage : ex.stage, stage_since: stageChanged ? nowISO() : ex.stage_since, log: [...(ex.log || []), { at: nowISO(), text: `Actualizado desde ${st.imp.file}${stageChanged ? ': etapa ' + FLOWS[r.flow].stages.find((s) => s.id === r.stage).label : ''}` }] }, { silent: true });
      updated++;
    } else {
      S.put('docs', { ...clean, priority: 3, round: ['revision', 'consultor', 'correcciones', 'correccion', 'vobo', 'diagrama', 'autorizacion', 'intramed', 'terminado'].includes(r.stage) ? 1 : 0, stage_since: r.start_date ? r.start_date + 'T12:00:00.000Z' : nowISO(), completed_at: null, log: [{ at: nowISO(), text: `Importado desde ${st.imp.file}${r.excel_status ? ' (estatus: ' + r.excel_status + ')' : ''}` }] }, { silent: true });
      created++;
    }
  }
  S.put('settings', { ...settingsOf(S), id: 'settings', last_doc_import: { at: nowISO(), file: st.imp.file } });
  U.closeSheet(); U.toast(`${created} documento(s) nuevos, ${updated} actualizados`);
  st.docFilter = 'proceso'; U.render();
}
function exportCsv() {
  const S = store();
  const rows = [['Código', 'Documento', 'Clase', 'Flujo', 'Etapa', 'En cancha de', 'Ronda', 'Días en etapa', 'Responsable', 'Área', 'Seguimiento', 'Inicio', 'Compromiso', 'Terminado', 'Versión', 'Enlace', 'Notas']];
  for (const d of S.all('docs').sort((a, b) => (a.code || a.name).localeCompare(b.code || b.name))) {
    const p = d.project_id ? S.get('projects', d.project_id) : null;
    rows.push([d.code, d.name, d.clase, FLOWS[d.flow]?.label, stageOf(d).label, isDocDone(d) ? '' : CANCHA[canchaOf(d)], d.round || 0, isDocDone(d) ? '' : daysIn(d), d.owner, d.area, p?.name || '', d.start_date || '', d.due_date || '', d.completed_at ? d.completed_at.slice(0, 10) : '', d.version, d.link, d.notes]);
  }
  const blob = new Blob([toCsv(rows)], { type: 'text/csv;charset=utf-8' });
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `avance-documental-${today()}.csv`; a.click();
}

// ---------------- Tarjeta para la vista Hoy ----------------
export function hoyCard() {
  const docs = store().all('docs').filter((d) => !isDocDone(d) && canchaOf(d) === 'calidad');
  const late = store().all('docs').filter((d) => !isDocDone(d) && canchaOf(d) !== 'calidad' && daysIn(d) > followDays());
  if (!docs.length && !late.length) return '';
  return `<div class="card"><div class="card-h"><h3>Documentos en mi cancha (${docs.length})</h3><a class="small" href="#documentos" data-act="doc-filter" data-f="mia">Ver</a></div>
  ${docs.slice(0, 5).map((d) => `<div class="irow" data-act="doc" data-id="${d.id}"><div class="small"><b>${esc(stageOf(d).label)}</b> · ${esc(docLabel(d))}</div><span class="small muted">${daysIn(d)} d</span></div>`).join('')}
  ${late.length ? `<div class="small danger-t" data-act="doc-filter" data-f="atrasados" style="cursor:pointer">${late.length} documento(s) esperando a otros más de ${followDays()} días</div>` : ''}</div>`;
}

// ---------------- Acciones y formularios ----------------
const actions = {
  'proj-new': () => openProject(null),
  'proj-tpl': (el) => openProject(null, el.dataset.t),
  proj: (el) => openProject(el.dataset.id),
  'proj-tab': (el) => { st.projTab = el.dataset.t; U.render(); },
  'proj-etapa': (el) => {
    const p = store().get('projects', el.dataset.id);
    const etapas = (p.etapas || []).map((e) => (e.id === el.dataset.e ? { ...e, done: !e.done, done_at: !e.done ? nowISO() : null } : e));
    const e = etapas.find((x) => x.id === el.dataset.e);
    store().put('projects', { ...p, etapas, log: [...(p.log || []), { at: nowISO(), text: `${e.done ? 'Etapa completada' : 'Etapa reabierta'}: ${e.title}` }] });
    openProject(p.id, 'blanco', true);
  },
  'proj-etapa-del': (el) => { const p = store().get('projects', el.dataset.id); store().put('projects', { ...p, etapas: (p.etapas || []).filter((e) => e.id !== el.dataset.e) }); openProject(p.id, 'blanco', true); },
  'proj-del': (el) => {
    if (!U.confirmInline(el, '¿Eliminar seguimiento?')) return;
    const id = el.dataset.id; store().remove('projects', id);
    for (const t of store().all('tasks').filter((x) => x.project_id === id && x.source === 'bloque_diario' && isOpen(x))) store().remove('tasks', t.id);
    U.closeSheet(); U.toast('Seguimiento eliminado (sus actividades y documentos se conservan)');
  },
  'doc-new': (el) => {
    const pid = el.dataset.p || null; const p = pid ? store().get('projects', pid) : null;
    openDoc(null, { project_id: pid, flow: p?.doc_flow || 'documento_nuevo', clase: p?.doc_flow === 'formato_ece' ? 'Formato' : '' });
  },
  doc: (el) => openDoc(el.dataset.id),
  'doc-next': (el) => { const d = store().get('docs', el.dataset.id); const nx = nextStage(d); if (nx) { moveDoc(d, nx.id, $('#doc-move-note')?.value.trim()); openDoc(d.id); U.toast('Etapa: ' + nx.label); } },
  'doc-back': (el) => { const d = store().get('docs', el.dataset.id); const bk = backStage(d); if (bk) { moveDoc(d, bk.id, $('#doc-move-note')?.value.trim()); openDoc(d.id); U.toast('Regresó con observaciones'); } },
  'doc-goto': (el) => { const d = store().get('docs', el.dataset.id); if (d.stage !== el.dataset.s) { moveDoc(d, el.dataset.s, $('#doc-move-note')?.value.trim()); openDoc(d.id); } },
  'doc-task': (el) => {
    const d = store().get('docs', el.dataset.id); const stg = stageOf(d);
    const t = store().put('tasks', { title: `${stg.task || 'Dar seguimiento'}: ${docLabel(d)}`, doc_id: d.id, doc_stage: stg.id, project_id: d.project_id || null, duration_min: stg.min || 30, category: d.flow === 'formato_ece' ? 'Formatos' : 'Documentos', priority: d.priority || 3, clasificacion: 'programable', status: 'pendiente', due_date: d.due_date || null, area: d.area || '', source: 'documento' });
    U.closeSheet(); U.openTask(t.id);
  },
  'doc-remind': (el) => {
    const d = store().get('docs', el.dataset.id);
    const when = nextWorkingDay(today(), settingsOf(store()), store());
    const t = store().put('tasks', { title: `Dar seguimiento con ${d.owner || CANCHA[canchaOf(d)]}: ${docLabel(d)}`, doc_id: d.id, doc_stage: d.stage, project_id: d.project_id || null, duration_min: 15, category: 'Documentos', priority: 3, clasificacion: 'seguimiento', status: 'pendiente', due_date: when, waiting_on: d.owner || CANCHA[canchaOf(d)], source: 'documento' });
    U.closeSheet(); U.toast('Recordatorio creado para ' + fmtShort(when)); void t;
  },
  'doc-del': (el) => { if (U.confirmInline(el, '¿Eliminar?')) { store().remove('docs', el.dataset.id); U.closeSheet(); } },
  'doc-filter': (el) => { st.docFilter = el.dataset.f; st.docStage = ''; if (location.hash.replace(/^#\/?/, '') !== 'documentos') location.hash = '#documentos'; else U.render(); },
  'doc-flow': (el) => { st.docFlow = el.dataset.f; st.docStage = ''; U.render(); },
  'doc-stage': (el) => { const same = st.docStage === el.dataset.s && st.docFlow === el.dataset.f; st.docFlow = same ? '' : el.dataset.f; st.docStage = same ? '' : el.dataset.s; if (st.docStage && st.docFilter === 'terminados') st.docFilter = 'todos'; U.render(); },
  'doc-import': openImport,
  'doc-import-run': runImport,
  'doc-export': exportCsv,
};

const forms = {
  project: (f, fd) => {
    const id = f.dataset.id; const S = store();
    const base = id ? S.get('projects', id) : { etapas: [], log: [{ at: nowISO(), text: 'Seguimiento creado' }] };
    const p = {
      ...base, name: fd.get('name').trim(), description: fd.get('description').trim(), status: fd.get('status'),
      priority: +fd.get('priority'), category: fd.get('category'), due_date: fd.get('due_date') || null, next_action: fd.get('next_action').trim(),
      doc_flow: fd.get('doc_flow') || '',
      daily: { min: +fd.get('daily_min') || 0, days: fd.getAll('daily_days').map(Number), start: fd.get('daily_start') || '', until: fd.get('daily_until') || '' },
    };
    if (id && base.status !== p.status) p.log = [...(p.log || []), { at: nowISO(), text: `Estado: ${EST_P[base.status] || base.status} → ${EST_P[p.status]}` }];
    const saved = S.put('projects', p);
    // Si cambió el bloque diario, se generan bloques nuevos (versión) y se quitan los futuros anteriores
    if (id && JSON.stringify(base.daily) !== JSON.stringify(p.daily)) S.put('projects', { ...saved, daily_v: (base.daily_v || 0) + 1 });
    const n = ensureDailyBlocks();
    U.closeSheet(); U.toast(id ? 'Guardado' : `Seguimiento creado${n ? ' · tiempo reservado en tu agenda' : ''}`);
    if (!id) openProject(saved.id);
  },
  'proj-etapa': (f, fd) => {
    const t = (fd.get('title') || '').trim(); if (!t) return;
    const p = store().get('projects', f.dataset.id);
    store().put('projects', { ...p, etapas: [...(p.etapas || []), { id: uid(), title: t, done: false }] });
    openProject(p.id, 'blanco', true); setTimeout(() => $('form[data-form="proj-etapa"] input')?.focus(), 50);
  },
  'proj-act': (f, fd) => {
    const t = (fd.get('title') || '').trim(); if (!t) return;
    const p = store().get('projects', f.dataset.id);
    store().put('tasks', { title: t, project_id: p.id, duration_min: +fd.get('dur') || 30, priority: p.priority || 3, category: p.category || 'General', clasificacion: 'programable', status: 'pendiente', source: 'seguimiento' });
    openProject(p.id, 'blanco', true); U.toast('Actividad agregada a la bandeja sin programar'); setTimeout(() => $('form[data-form="proj-act"] input')?.focus(), 50);
  },
  'proj-note': (f, fd) => {
    const t = (fd.get('text') || '').trim(); if (!t) return;
    const p = store().get('projects', f.dataset.id);
    store().put('projects', { ...p, log: [...(p.log || []), { at: nowISO(), text: t }] });
    if (f.dataset.close) { U.closeSheet(); U.toast('Anotado en la bitácora'); } else openProject(p.id, 'blanco', true);
  },
  doc: (f, fd) => {
    const id = f.dataset.id; const S = store();
    const base = id ? S.get('docs', id) : { log: [{ at: nowISO(), text: 'Registrado' }], stage_since: nowISO() };
    const d = { ...base };
    ['code', 'name', 'clase', 'flow', 'owner', 'area', 'version', 'link', 'notes'].forEach((k) => { d[k] = (fd.get(k) || '').toString().trim(); });
    ['start_date', 'due_date'].forEach((k) => { d[k] = fd.get(k) || null; });
    d.priority = +fd.get('priority') || 3; d.project_id = fd.get('project_id') || null;
    if (!id) { d.stage = fd.get('stage'); d.round = FLOWS[d.flow].stages.find((x) => x.id === d.stage)?.round ? 1 : 0; }
    else {
      d.round = +fd.get('round') || 0;
      if (base.flow !== d.flow && !FLOWS[d.flow].stages.some((x) => x.id === d.stage)) { d.stage = FLOWS[d.flow].stages[0].id; d.stage_since = nowISO(); }
      d.stage_unsure = false;
      const nl = (fd.get('newlog') || '').trim(); if (nl) d.log = [...(d.log || []), { at: nowISO(), text: nl }];
    }
    if (!FLOWS[d.flow].stages.some((x) => x.id === d.stage)) d.stage = FLOWS[d.flow].stages[0].id;
    const saved = S.put('docs', d);
    if (!id) maybeCreateDocTask(saved);
    U.closeSheet(); U.toast(id ? 'Guardado' : 'Documento registrado');
  },
};

function onChange(el) {
  if (el.matches('[data-docimport]') && el.files[0]) { loadImportFile(el.files[0]); return true; }
  if (el.matches('[data-impmap]') && st.imp) { st.imp.map[el.dataset.impmap] = el.value === '' ? undefined : +el.value; renderImport(); return true; }
  if (el.matches('[data-impsheet]') && st.imp) { pickSheet(+el.value); return true; }
  if (el.matches('[data-impformat]') && st.imp) { st.imp.formatFlow = el.value; renderImport(); return true; }
  return false;
}
export const segState = st;
void fmtLong; void toMin; void toHHMM; void $$;
