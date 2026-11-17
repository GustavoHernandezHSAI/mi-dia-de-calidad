// Mi Día de Calidad — interfaz principal
import { CONFIG } from './config.js';
import { store } from './store.js';
import { outlook } from './outlook.js';
import { parseTask } from './parser.js';
import {
  today, addDays, weekStart, fmtShort, fmtLong, fmtDur, toMin, toHHMM, nowMin, roundUp, esc, $, $$, relDue, copyText, uid, DIAS_C, diffDays, dow,
} from './util.js';
import {
  settingsOf, dayInfo, dayRec, autoPlan, applyPlan, splitTask, bomberazo, applyChanges, score, byScore, flags, isOpen,
  nextWorkingDay, isIndicadoresWindow, CLASIFICACION, ESTADOS, CATEGORIAS, TIPOS_DIA, DEFAULT_SETTINGS,
} from './planner.js';
import { closureOutputs, periodSummary } from './summary.js';
import { initSeguimiento, segState } from './seguimiento.js';

// ---------------- Estado de la interfaz ----------------
const state = {
  view: 'hoy', date: today(), week: weekStart(today()), filter: 'abiertas', q: '',
  intrTab: 'abiertos', intrStatus: '', intrTipo: '', closure: null, closureResult: null, period: 'hoy',
};
const ctx = () => ({ S: settingsOf(store), store, outlook: outlook.events });
const PRIO = { 1: 'Crítica', 2: 'Alta', 3: 'Media', 4: 'Baja' };
// Pendiente sin horario (los bloques diarios de otros días no cuentan)
const floating = (t) => isOpen(t) && !t.scheduled_start && t.clasificacion !== 'referencia' && !(t.source === 'bloque_diario' && t.scheduled_date !== today());
let SEG = null;
const INTR_EST = { recibido: 'Recibido', en_analisis: 'En análisis', enviado: 'Enviado a Intramed', en_desarrollo: 'En desarrollo', en_validacion: 'En validación', implementado: 'Implementado', cerrado: 'Cerrado' };
const INTR_TIPO = { reporte: 'Reporte (incidencia / falla)', necesidad: 'Necesidad nueva', soporte: 'Soporte técnico' };
const INTR_TIPO_C = { reporte: 'Reporte', necesidad: 'Necesidad nueva', soporte: 'Soporte' };
const INTR_ORIGEN = { hospital: 'Hospital / unidad', calidad: 'Calidad corporativa', otra: 'Otra área' };
const PX = 1.3; // píxeles por minuto en la línea de tiempo

// ---------------- Arranque ----------------
async function boot() {
  applyTheme();
  await outlook.handleRedirect();
  const res = await store.init();
  if (!res.ready) { renderLogin(res.error); return; }
  start();
}

function start() {
  document.body.classList.remove('login-mode');
  SEG = initSeguimiento({ store, ctx, openSheet, closeSheet, toast, chip, taskRow, render, hospitalsDatalist, confirmInline, openTask });
  try { SEG.migrate(); SEG.ensureDailyBlocks(); } catch (e) { console.error(e); }
  $('#app').innerHTML = shellHTML();
  window.addEventListener('hashchange', route);
  store.on((info) => {
    if (info.signedOut) { location.reload(); return; }
    if (info.recovery) openPasswordSheet();
    scheduleRender();
  });
  if (store._recovery) openPasswordSheet();
  outlook.loadCache();
  refreshOutlook();
  setInterval(refreshOutlook, 10 * 60 * 1000);
  let lastDay = today();
  setInterval(() => { if (today() !== lastDay) { lastDay = today(); state.date = today(); SEG.ensureDailyBlocks(); } if (state.view === 'hoy') scheduleRender(); tickNotifications(); }, 60 * 1000);
  // Tras sincronizar con la nube, asegurar bloques del día (sin duplicar: el id es fijo por día)
  setTimeout(() => SEG.ensureDailyBlocks(), 4000);
  bindEvents();
  route();
  if ('serviceWorker' in navigator && location.protocol !== 'file:') navigator.serviceWorker.register('sw.js').catch(() => {});
}

function route() {
  const v = (location.hash.replace(/^#\/?/, '') || 'hoy').split('?')[0];
  state.view = ['hoy', 'semana', 'tareas', 'seguimientos', 'documentos', 'intramed', 'cierre', 'resumen', 'ajustes', 'mas'].includes(v) ? v : 'hoy';
  render();
  window.scrollTo(0, 0);
}

let rq = null;
function scheduleRender() { if (rq) return; rq = requestAnimationFrame(() => { rq = null; render(); }); }

function render() {
  const main = $('#view'); if (!main) return;
  document.body.dataset.view = state.view;
  const views = { hoy: viewHoy, semana: viewSemana, tareas: viewTareas, intramed: viewIntramed, cierre: viewCierre, resumen: viewResumen, ajustes: viewAjustes, mas: viewMas, ...SEG.views };
  try { main.innerHTML = views[state.view](); }
  catch (e) { console.error(e); main.innerHTML = `<div class="card"><h3>Ocurrió un error al mostrar esta vista</h3><pre class="small">${esc(e.stack || e)}</pre></div>`; }
  $$('.nav a').forEach((a) => a.classList.toggle('active', a.dataset.v === state.view || (state.view === 'mas' && a.dataset.v === 'mas') || (['semana', 'tareas', 'cierre', 'resumen', 'ajustes'].includes(state.view) && a.dataset.v === 'mas' && a.closest('.bottomnav'))));
  renderSync();
  if (state.view === 'hoy') { const n = $('.now-line'); if (n && !render._scrolled) { render._scrolled = true; } }
}

function renderSync() {
  const el = $('#sync'); if (!el) return;
  const s = store.status;
  const txt = { local: 'Modo local', online: 'Sincronizado', syncing: 'Sincronizando…', offline: `Sin conexión${store.pending ? ` · ${store.pending} por enviar` : ''}`, error: 'Error de sincronización' }[s];
  el.className = 'sync ' + s; el.title = txt; el.innerHTML = `<i></i><span>${txt}</span>`;
}

// ---------------- Tema ----------------
function applyTheme() {
  const t = localStorage.getItem('mdc:theme') || 'sistema';
  const dark = t === 'oscuro' || (t === 'sistema' && matchMedia('(prefers-color-scheme: dark)').matches);
  document.documentElement.dataset.theme = dark ? 'dark' : 'light';
}

// ---------------- Login ----------------
function renderLogin(err) {
  document.body.classList.add('login-mode');
  $('#app').innerHTML = `
  <div class="login">
    <form class="card login-card" data-form="login">
      <div class="brand big"><span class="logo">✓</span> Mi Día de Calidad</div>
      <p class="muted">Inicia sesión para sincronizar tu información entre computadora y celular.</p>
      ${err ? `<p class="alert danger">${esc(err)}</p>` : ''}
      <label>Correo<input name="email" type="email" autocomplete="username" required></label>
      <label>Contraseña<input name="password" type="password" autocomplete="current-password" required></label>
      <button class="btn primary block" type="submit">Entrar</button>
      <button class="btn ghost block" type="button" id="forgot">Olvidé mi contraseña</button>
      <p class="msg small"></p>
    </form>
  </div>`;
  const f = $('form[data-form="login"]');
  f.addEventListener('submit', async (e) => {
    e.preventDefault();
    const fd = new FormData(f); const msg = $('.msg', f);
    msg.textContent = 'Entrando…';
    try { await store.signIn(fd.get('email').trim(), fd.get('password')); start(); }
    catch (er) { msg.textContent = er.message === 'Invalid login credentials' ? 'Correo o contraseña incorrectos.' : er.message; }
  });
  $('#forgot').addEventListener('click', async () => {
    const email = f.email.value.trim(); const msg = $('.msg', f);
    if (!email) { msg.textContent = 'Escribe tu correo primero.'; return; }
    try { await store.resetPassword(email); msg.textContent = 'Te enviamos un correo para restablecer la contraseña.'; }
    catch (er) { msg.textContent = er.message; }
  });
}

function openPasswordSheet() {
  store._recovery = false;
  openSheet('Nueva contraseña', `
    <form data-form="newpass" class="stack">
      <label>Nueva contraseña<input name="p1" type="password" minlength="8" required autocomplete="new-password"></label>
      <label>Repetir<input name="p2" type="password" minlength="8" required autocomplete="new-password"></label>
      <button class="btn primary block">Guardar contraseña</button>
    </form>`);
}

// ---------------- Estructura ----------------
function shellHTML() {
  const nav = [['hoy', 'Hoy', 'sun'], ['semana', 'Semana', 'cal'], ['tareas', 'Tareas', 'list'], ['seguimientos', 'Seguimientos', 'layers'], ['documentos', 'Avance documental', 'doc'], ['intramed', 'Intramed', 'tool'], ['cierre', 'Cierre del día', 'moon'], ['resumen', 'Resumen', 'send'], ['ajustes', 'Ajustes', 'gear']];
  const bottom = [['hoy', 'Hoy', 'sun'], ['seguimientos', 'Seguim.', 'layers'], ['documentos', 'Docs', 'doc'], ['intramed', 'Intramed', 'tool'], ['mas', 'Más', 'dots']];
  return `
  <header class="topbar">
    <a class="brand" href="#hoy"><span class="logo">✓</span><span class="brand-t">Mi Día de Calidad</span></a>
    <span id="sync" class="sync"></span>
    <div class="top-actions">
      <button class="btn fire" data-act="bomberazo" title="Imprevisto (I)">${ico('fire')}<span>Imprevisto</span></button>
      <button class="btn primary" data-act="new" title="Nueva tarea (N)">${ico('plus')}<span>Nueva tarea</span></button>
    </div>
  </header>
  <nav class="nav sidenav">${nav.map(([v, l, i]) => `<a href="#${v}" data-v="${v}">${ico(i)}<span>${l}</span></a>`).join('')}
    <div class="side-foot small muted">${store.mode === 'local' ? 'Modo local: datos solo en este navegador' : esc(store.user?.email || '')}</div>
  </nav>
  <main id="view" class="view"></main>
  <div class="fabs">
    <button class="fab fire" data-act="bomberazo" aria-label="Imprevisto">${ico('fire')}</button>
    <button class="fab primary" data-act="new" aria-label="Nueva tarea">${ico('plus')}</button>
  </div>
  <nav class="nav bottomnav">${bottom.map(([v, l, i]) => `<a href="#${v}" data-v="${v}">${ico(i)}<span>${l}</span></a>`).join('')}</nav>
  <div id="sheet" class="sheet-wrap" hidden><div class="sheet-bg" data-act="close"></div><div class="sheet" role="dialog" aria-modal="true"><div class="sheet-head"><h2 id="sheet-title"></h2><button class="icon-btn" data-act="close" aria-label="Cerrar">✕</button></div><div id="sheet-body" class="sheet-body"></div></div></div>
  <div id="toast" class="toast" hidden></div>`;
}

function ico(n) {
  const p = {
    sun: '<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/>',
    cal: '<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4"/>',
    list: '<path d="M9 6h12M9 12h12M9 18h12"/><circle cx="4" cy="6" r="1"/><circle cx="4" cy="12" r="1"/><circle cx="4" cy="18" r="1"/>',
    tool: '<path d="M14.7 6.3a4 4 0 0 0-5.4 5.4L3 18l3 3 6.3-6.3a4 4 0 0 0 5.4-5.4l-2.5 2.5-2.5-.5-.5-2.5z"/>',
    moon: '<path d="M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8z"/>',
    send: '<path d="M22 2 11 13M22 2l-7 20-4-9-9-4z"/>',
    gear: '<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"/>',
    layers: '<path d="M12 2 2 7l10 5 10-5-10-5z"/><path d="m2 17 10 5 10-5"/><path d="m2 12 10 5 10-5"/>',
    doc: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6M8 13h8M8 17h5"/>',
    dots: '<circle cx="5" cy="12" r="1.5"/><circle cx="12" cy="12" r="1.5"/><circle cx="19" cy="12" r="1.5"/>',
    fire: '<path d="M12 22c4 0 7-3 7-7 0-4-3-6-4-9-1 2-2 3-3 3 0-2-1-4-3-6 0 4-4 6-4 11 0 5 3 8 7 8z"/>',
    plus: '<path d="M12 5v14M5 12h14"/>',
    check: '<path d="M20 6 9 17l-5-5"/>',
  }[n] || '';
  return `<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${p}</svg>`;
}

// ---------------- Sheets y toasts ----------------
function openSheet(title, html, { wide = false, keep = false } = {}) {
  const body = $('#sheet-body'); const y = body.scrollTop;
  $('#sheet-title').textContent = title;
  body.innerHTML = html;
  body.scrollTop = keep && !$('#sheet').hidden ? y : 0;
  const w = $('#sheet'); w.hidden = false; $('.sheet', w).classList.toggle('wide', wide);
  document.body.classList.add('noscroll');
  setTimeout(() => { const a = $('#sheet [autofocus]'); if (a) a.focus(); }, 30);
}
function closeSheet() { $('#sheet').hidden = true; document.body.classList.remove('noscroll'); }
let toastT = null;
function toast(msg, action) {
  const t = $('#toast'); t.hidden = false;
  t.innerHTML = `<span>${esc(msg)}</span>${action ? `<button class="btn small ghost" id="toast-act">${esc(action.label)}</button>` : ''}`;
  if (action) $('#toast-act').onclick = () => { action.fn(); t.hidden = true; };
  clearTimeout(toastT); toastT = setTimeout(() => { t.hidden = true; }, action ? 7000 : 3500);
}

// ---------------- Componentes ----------------
const chip = (t, cls = '') => `<span class="chip ${cls}">${t}</span>`;
function taskChips(t, { showSched = true } = {}) {
  const c = ctx(); const f = flags(t, c); const out = [];
  if (t.is_urgent) out.push(chip('Imprevisto', 'fire'));
  if (t.riesgo) out.push(chip('Riesgo / cumplimiento', 'danger'));
  if (f.overdue) out.push(chip(esc(relDue(t.due_date)), 'danger'));
  else if (t.due_date) out.push(chip(esc(relDue(t.due_date)), f.atRisk ? 'warn' : ''));
  if (f.blocked) out.push(chip('Bloqueada', 'danger'));
  if (f.decision) out.push(chip('Requiere info/decisión', 'warn'));
  if (t.clasificacion === 'delegable') out.push(chip(t.delegation_status === 'autorizada' ? 'Delegación autorizada' : 'Delegable (por autorizar)', 'info'));
  if (t.clasificacion === 'seguimiento') out.push(chip('Seguimiento'));
  if (t.clasificacion === 'referencia') out.push(chip('Referencia'));
  if (t.project_id) { const p = store.get('projects', t.project_id); if (p) out.push(chip(esc(p.name), 'info')); }
  if (t.category && t.category !== 'General') out.push(chip(esc(t.category), 'soft'));
  if (t.area) out.push(chip(esc(t.area), 'soft'));
  if (showSched && t.scheduled_date && isOpen(t)) out.push(chip(`${fmtShort(t.scheduled_date)}${t.scheduled_start ? ' ' + t.scheduled_start : ''}`, f.stale ? 'warn' : 'soft'));
  out.push(chip(fmtDur(t.duration_min || 30), 'soft'));
  return out.join('');
}
function taskRow(t, opts = {}) {
  const done = t.status === 'completada';
  return `<div class="trow ${done ? 'done' : ''} p${t.priority || 3}" ${opts.drag ? `draggable="true" data-drag-task="${t.id}"` : ''}>
    <button class="check ${done ? 'on' : ''}" data-act="toggle-done" data-id="${t.id}" aria-label="Completar">${ico('check')}</button>
    <div class="trow-main" data-act="task" data-id="${t.id}">
      <div class="trow-title"><span class="pdot p${t.priority || 3}" title="Prioridad ${PRIO[t.priority || 3]}"></span>${esc(t.title)}</div>
      <div class="chips">${taskChips(t, opts)}</div>
      ${opts.next && t.next_action ? `<div class="small muted">Siguiente: ${esc(t.next_action)}</div>` : ''}
    </div>
  </div>`;
}

function alertCounts() {
  const c = ctx(); const open = store.all('tasks').filter(isOpen);
  const n = { overdue: 0, atRisk: 0, blocked: 0, decision: 0, delegable: 0, stale: 0 };
  open.forEach((t) => { const f = flags(t, c); Object.keys(n).forEach((k) => { if (f[k]) n[k]++; }); });
  return n;
}

// ================= VISTA HOY =================
function viewHoy() {
  const c = ctx(); const S = c.S; const d = state.date; const isToday = d === today();
  const info = dayInfo(d, c);
  const n = alertCounts();
  const unsched = store.all('tasks').filter(floating).sort(byScore);
  const stale = store.all('tasks').filter((t) => isOpen(t) && t.scheduled_date && t.scheduled_date < today()).sort(byScore);
  const bz = store.all('bomberazos').filter((b) => b.date === d && b.state === 'aplicado');
  const banners = [];
  if (store.mode === 'local') banners.push(`<div class="alert info small">Modo local: los datos se guardan solo en este navegador. Configura Supabase para sincronizar (ver instrucciones).</div>`);
  if (isIndicadoresWindow(d, S, store)) banners.push(`<div class="alert info">Periodo de revisión mensual de indicadores de los gerentes (primeros ${S.indicadores_dias} días hábiles). <button class="btn small" data-act="new-prefill" data-text="Revisar indicadores de calidad de los gerentes del mes">Crear tarea</button></div>`);
  bz.forEach((b) => banners.push(`<div class="alert fire"><b>Reorganización por imprevisto:</b> ${esc(b.title)} — ${b.changes.length} cambio(s). <span class="btnrow"><button class="btn small" data-act="bz-view" data-id="${b.id}">Ver cambios</button><button class="btn small primary" data-act="bz-accept" data-id="${b.id}">Aceptar</button><button class="btn small ghost" data-act="bz-revert" data-id="${b.id}">Revertir</button></span></div>`));
  if (outlook.error) banners.push(`<div class="alert warn small">Outlook: ${esc(outlook.error)} <a href="#ajustes">Ajustes</a></div>`);
  if (info.working && !info.hasFocus && info.D?.tipo !== 'visita') banners.push(`<div class="alert warn small">Hoy no queda un bloque libre de ${S.focus_min} min para trabajo concentrado. Considera mover algo.</div>`);
  if ((dow(d) === 2 || dow(d) === 5) && info.fixed.some((f) => f.kind === 'intramed' || /intramed/i.test(f.title))) {
    const openI = store.all('intramed').filter((i) => i.status !== 'cerrado').length;
    if (openI) banners.push(`<div class="alert small">Reunión con Intramed hoy · ${openI} asunto(s) abiertos. <a href="#intramed">Preparar</a></div>`);
  }
  const conflictsN = Object.keys(info.conflicts).length;
  if (conflictsN) banners.push(`<div class="alert danger small">${conflictsN} tarea(s) con conflicto de horario (marcadas en rojo). <button class="btn small" data-act="fix-conflicts">Reacomodar automáticamente</button></div>`);

  const pct = info.plannable ? Math.round((info.used / info.plannable) * 100) : 0;
  const nowBlock = isToday && info.working ? currentNext(info) : null;
  return `
  <section class="day-head">
    <div class="datenav">
      <button class="icon-btn" data-act="day-prev" aria-label="Día anterior">‹</button>
      <div class="datenav-t"><h1>${isToday ? 'Hoy' : esc(fmtShort(d))}</h1><div class="muted">${esc(fmtLong(d))}</div></div>
      <button class="icon-btn" data-act="day-next" aria-label="Día siguiente">›</button>
      ${!isToday ? '<button class="btn small" data-act="day-today">Ir a hoy</button>' : ''}
    </div>
    <div class="btnrow">
      ${chip(esc(TIPOS_DIA[info.tipo] || 'Normal'), info.working ? 'soft' : 'warn')}
      <button class="btn small" data-act="day-settings">Ajustar día</button>
      <button class="btn small" data-act="add-event">+ Reunión imprevista</button>
    </div>
  </section>
  <div class="alerts-row">
    ${alertBtn('overdue', 'Retrasadas', n.overdue + n.stale, 'danger')}
    ${alertBtn('atRisk', 'En riesgo', n.atRisk, 'warn')}
    ${alertBtn('blocked', 'Bloqueadas', n.blocked, 'danger')}
    ${alertBtn('decision', 'Requieren info/decisión', n.decision, 'warn')}
    ${alertBtn('delegable', 'Para delegar', n.delegable, 'info')}
  </div>
  ${banners.join('')}
  ${nowBlock ? `<div class="card now-card">${nowBlock}</div>` : ''}
  <div class="hoy-grid">
    <section class="card tl-card">
      ${info.working ? `
      <div class="cap">
        <div class="cap-bar"><i style="width:${Math.min(100, pct)}%" class="${pct > 100 ? 'over' : pct > 85 ? 'high' : ''}"></i></div>
        <div class="small muted">Planeado ${fmtDur(info.used)} de ${fmtDur(info.plannable)} planeables · margen ${S.buffer_pct}% (${fmtDur(info.capacity - info.plannable)}) para imprevistos</div>
      </div>
      ${timeline(info, isToday)}` : `<div class="empty"><h3>Día no laboral</h3><p class="muted">No se programa trabajo este día.${info.tasks.length ? ' Hay tareas asignadas aquí: muévelas.' : ''}</p>${info.tasks.map((b) => taskRow(b.task)).join('')}</div>`}
    </section>
    <aside class="side">
      ${SEG.hoyCard()}
      ${stale.length ? `<div class="card"><div class="card-h"><h3>De días anteriores (${stale.length})</h3><button class="btn small" data-act="replan-stale">Reprogramar</button></div>${stale.slice(0, 6).map((t) => taskRow(t)).join('')}</div>` : ''}
      <div class="card drop-unsched" data-drop-unsched="1">
        <div class="card-h"><h3>Sin programar (${unsched.length})</h3>${unsched.length ? '<button class="btn small primary" data-act="suggest">Sugerir horarios</button>' : ''}</div>
        ${unsched.length ? unsched.slice(0, 12).map((t) => taskRow(t, { drag: true, next: true })).join('') : '<p class="muted small">Todo lo pendiente tiene lugar en la agenda.</p>'}
        ${unsched.length > 12 ? `<a class="small" href="#tareas" data-act="filter" data-f="sinprog">Ver todas</a>` : ''}
      </div>
    </aside>
  </div>`;
}
function alertBtn(f, label, n, cls) {
  return `<button class="alert-btn ${n ? cls : 'zero'}" data-act="filter" data-f="${f}"><b>${n}</b><span>${label}</span></button>`;
}
function currentNext(info) {
  const n = nowMin();
  const all = [...info.fixed.map((f) => ({ ...f, isT: false })), ...info.tasks.filter((t) => isOpen(t.task)).map((t) => ({ ...t, isT: true }))].sort((a, b) => a.s - b.s);
  const cur = all.find((b) => b.s <= n && b.e > n);
  const nxt = all.find((b) => b.s > n);
  const fmt = (b) => `<b>${toHHMM(b.s)}–${toHHMM(b.e)}</b> ${esc(b.title)}`;
  if (!cur && !nxt) return n >= info.win.e - 30 ? `<div>Fin de la jornada. <a href="#cierre">Hacer cierre del día</a></div>` : '<div class="muted">Sin bloques próximos.</div>';
  return `${cur ? `<div><span class="small muted">Ahora</span><div>${fmt(cur)}</div></div>` : ''}${nxt ? `<div><span class="small muted">Siguiente</span><div>${fmt(nxt)}</div></div>` : ''}`;
}

function timeline(info, isToday) {
  const S = ctx().S;
  const s0 = Math.floor(info.win.s / 60) * 60, e0 = Math.ceil(info.win.e / 60) * 60;
  const H = (e0 - s0) * PX;
  const y = (m) => (m - s0) * PX;
  let hours = '';
  for (let m = s0; m <= e0; m += 60) hours += `<div class="hr" style="top:${y(m)}px"><span>${toHHMM(m)}</span></div>`;
  const free = info.free.map((f) => {
    const isFocus = info.focus && f.s <= info.focus.s && f.e >= info.focus.e;
    const past = isToday && f.e <= nowMin();
    if (isFocus && f.s < info.focus.s) {
      // Parte del espacio antes del bloque de foco (p. ej., tiempo ya transcurrido)
      return `<div class="slot ${isToday && info.focus.s <= nowMin() + 15 ? 'past' : ''}" style="top:${y(f.s)}px;height:${(info.focus.s - f.s) * PX}px" data-act="slot" data-s="${f.s}" data-e="${info.focus.s}"><span>Libre</span></div><div class="slot focus" style="top:${y(info.focus.s)}px;height:${(f.e - info.focus.s) * PX}px" data-act="slot" data-s="${info.focus.s}" data-e="${f.e}"><span>Bloque de foco protegido · ${fmtDur(info.focus.e - info.focus.s)}</span></div>`;
    }
    return `<div class="slot ${isFocus ? 'focus' : ''} ${past ? 'past' : ''}" style="top:${y(f.s)}px;height:${(f.e - f.s) * PX}px" data-act="slot" data-s="${f.s}" data-e="${f.e}">${isFocus ? `<span>Bloque de foco protegido · ${fmtDur(info.focus.e - info.focus.s)}</span>` : `<span>Libre ${toHHMM(f.s)}–${toHHMM(f.e)}</span>`}</div>`;
  }).join('');
  const fixed = info.fixed.map((f) => {
    const s = Math.max(f.s, s0), e = Math.min(f.e, e0);
    if (e <= s) return '';
    return `<div class="blk fixed k-${f.kind}" style="top:${y(s)}px;height:${Math.max(18, (e - s) * PX)}px" data-act="fixed" data-id="${esc(f.id)}"><b>${toHHMM(f.s)}–${toHHMM(f.e)}</b> ${esc(f.title)}${f.kind === 'outlook' ? ' <small>· Outlook</small>' : ''}</div>`;
  }).join('');
  // Tareas: columnas para empalmes
  const lanes = [];
  const tasks = info.tasks.filter((b) => b.task.status !== 'cancelada').map((b) => {
    let lane = 0; while (lanes[lane] && lanes[lane] > b.s) lane++;
    lanes[lane] = b.e; return { ...b, lane };
  });
  const nl = Math.max(1, lanes.length);
  const tb = tasks.map((b) => {
    const t = b.task; const conf = info.conflicts[b.id];
    const h = Math.max(22, (b.e - b.s) * PX);
    const onFixed = info.fixed.some((f) => f.s < b.e && b.s < f.e);
    const L = onFixed ? 42 : 0; // si choca con una reunión, se corre a la derecha para que ambas se vean
    return `<div class="blk task p${t.priority || 3} ${t.status === 'completada' ? 'done' : ''} ${conf ? 'conflict' : ''} ${t.is_urgent ? 'urgent' : ''} ${t.status === 'bloqueada' ? 'blocked' : ''}" draggable="true" data-drag-task="${t.id}" style="top:${y(b.s)}px;height:${h}px;left:calc(${L + (b.lane / nl) * (100 - L)}% + 2px);width:calc(${(100 - L) / nl}% - 4px)" data-act="task" data-id="${t.id}" title="${esc(conf ? conf.join(' · ') : '')}">
      <button class="check mini ${t.status === 'completada' ? 'on' : ''}" data-act="toggle-done" data-id="${t.id}" aria-label="Completar">${ico('check')}</button>
      <div class="blk-t"><b>${toHHMM(b.s)}</b> ${esc(t.title)}</div>
      ${h > 40 ? `<div class="blk-s">${t.area ? esc(t.area) + ' · ' : ''}${fmtDur(b.e - b.s)}${conf ? ' · ⚠ ' + esc(conf[0]) : ''}</div>` : ''}
    </div>`;
  }).join('');
  const now = isToday && nowMin() >= s0 && nowMin() <= e0 ? `<div class="now-line" style="top:${y(nowMin())}px"></div>` : '';
  return `<div class="timeline" style="height:${H}px" data-drop-timeline="1" data-s0="${s0}">${hours}<div class="lanes">${free}${fixed}${tb}</div>${now}<div class="drop-ghost" hidden></div></div>
  <p class="small muted tl-help">Toca un espacio libre para colocar una tarea. En computadora puedes arrastrar tareas.</p>`;
}

// ================= VISTA SEMANA =================
function viewSemana() {
  const c = ctx(); const ws = state.week;
  const days = [0, 1, 2, 3, 4, 5].map((i) => addDays(ws, i));
  const sun = addDays(ws, 6);
  if (store.all('tasks').some((t) => t.scheduled_date === sun && isOpen(t))) days.push(sun);
  const unsched = store.all('tasks').filter(floating).sort(byScore);
  const cols = days.map((d) => {
    const info = dayInfo(d, c);
    const pct = info.plannable ? Math.round((info.used / info.plannable) * 100) : (info.used ? 999 : 0);
    const items = [...info.fixed.map((f) => ({ ...f, isT: false })), ...info.tasks.map((t) => ({ ...t, isT: true }))].sort((a, b) => a.s - b.s);
    return `<div class="wcol ${d === today() ? 'today' : ''} ${info.working ? '' : 'off'}" data-drop-day="${d}">
      <div class="wcol-h" data-act="goto-day" data-d="${d}">
        <div><b>${DIAS_C[dow(d)]} ${+d.slice(8)}</b> ${info.working ? '' : `<span class="small muted">${esc(TIPOS_DIA[info.tipo])}</span>`}</div>
        ${info.working ? `<div class="load"><i style="width:${Math.min(100, pct)}%" class="${pct > 100 ? 'over' : pct > 85 ? 'high' : ''}"></i></div><div class="small muted">${fmtDur(info.used)} / ${fmtDur(info.plannable)}${Object.keys(info.conflicts).length ? ' · <span class="danger-t">⚠ conflictos</span>' : ''}</div>` : ''}
      </div>
      <div class="wcol-b">
        ${items.map((b) => b.isT
          ? `<div class="wtask p${b.task.priority || 3} ${b.task.status === 'completada' ? 'done' : ''} ${info.conflicts[b.id] ? 'conflict' : ''}" draggable="true" data-drag-task="${b.task.id}" data-act="task" data-id="${b.task.id}"><b>${toHHMM(b.s)}</b> ${esc(b.task.title)}</div>`
          : `<div class="wfixed k-${b.kind}" data-act="fixed" data-id="${esc(b.id)}" data-d="${d}">${toHHMM(b.s)} ${esc(b.title)}</div>`).join('')}
      </div>
    </div>`;
  }).join('');
  return `
  <section class="day-head">
    <div class="datenav">
      <button class="icon-btn" data-act="week-prev" aria-label="Semana anterior">‹</button>
      <div class="datenav-t"><h1>Semana</h1><div class="muted">${esc(fmtShort(ws))} – ${esc(fmtShort(addDays(ws, 5)))}</div></div>
      <button class="icon-btn" data-act="week-next" aria-label="Semana siguiente">›</button>
      ${ws !== weekStart(today()) ? '<button class="btn small" data-act="week-today">Esta semana</button>' : ''}
    </div>
    <div class="btnrow">${unsched.length ? `<button class="btn primary small" data-act="suggest-week">Sugerir horarios para ${unsched.length} pendiente(s)</button>` : ''}</div>
  </section>
  <div class="week">${cols}</div>
  <div class="card drop-unsched" data-drop-unsched="1">
    <div class="card-h"><h3>Sin programar (${unsched.length})</h3></div>
    <div class="wunsched">${unsched.map((t) => `<div class="wtask p${t.priority || 3}" draggable="true" data-drag-task="${t.id}" data-act="task" data-id="${t.id}">${esc(t.title)} <small>${fmtDur(t.duration_min || 30)}</small></div>`).join('') || '<p class="small muted">Nada pendiente sin horario.</p>'}</div>
    <p class="small muted">Arrastra una tarea a un día (computadora) o tócala y elige “Mover a” (celular).</p>
  </div>`;
}

// ================= VISTA TAREAS =================
const FILTERS = {
  abiertas: ['Abiertas', (t) => isOpen(t)],
  sinprog: ['Sin programar', (t) => floating(t)],
  overdue: ['Retrasadas', (t, f) => f.overdue || f.stale],
  atRisk: ['En riesgo', (t, f) => f.atRisk],
  blocked: ['Bloqueadas', (t, f) => f.blocked],
  decision: ['Requieren info/decisión', (t, f) => f.decision && isOpen(t)],
  delegable: ['Para delegar', (t, f) => f.delegable && isOpen(t)],
  seguimiento: ['Seguimiento', (t) => isOpen(t) && t.clasificacion === 'seguimiento'],
  referencia: ['Referencia', (t) => isOpen(t) && t.clasificacion === 'referencia'],
  completadas: ['Completadas', (t) => t.status === 'completada'],
};
function viewTareas() {
  const c = ctx(); const all = store.all('tasks');
  const q = state.q.trim().toLowerCase();
  const fl = FILTERS[state.filter] || FILTERS.abiertas;
  const hideBlk = (t) => !(t.source === 'bloque_diario' && t.scheduled_date !== today() && isOpen(t));
  let list = all.filter((t) => hideBlk(t) && fl[1](t, flags(t, c)));
  if (q) list = list.filter((t) => [t.title, t.notes, t.area, t.category].join(' ').toLowerCase().includes(q));
  list.sort(state.filter === 'completadas' ? (a, b) => (b.completed_at || '').localeCompare(a.completed_at || '') : byScore);
  const counts = Object.fromEntries(Object.entries(FILTERS).map(([k, [, fn]]) => [k, all.filter((t) => hideBlk(t) && fn(t, flags(t, c))).length]));
  return `
  <section class="day-head"><div class="datenav-t"><h1>Tareas</h1><div class="muted">Ordenadas por prioridad: riesgo y cumplimiento → fecha límite → impacto → dependencias → rapidez</div></div></section>
  <input class="search" type="search" placeholder="Buscar…" value="${esc(state.q)}" data-input="q">
  <div class="filters">${Object.entries(FILTERS).map(([k, [l]]) => `<button class="fchip ${state.filter === k ? 'on' : ''}" data-act="filter" data-f="${k}">${l} <b>${counts[k]}</b></button>`).join('')}</div>
  <div class="card list">${list.length ? list.slice(0, 200).map((t) => taskRow(t, { next: true })).join('') : '<p class="muted">Sin tareas en este filtro.</p>'}</div>`;
}

// ================= VISTA INTRAMED =================
function viewIntramed() {
  const items = store.all('intramed');
  const ref = today();
  const open = items.filter((i) => i.status !== 'cerrado');
  if (state.intrTab === 'reuniones') {
    const ms = store.all('intramed_meetings').sort((a, b) => b.date.localeCompare(a.date));
    return intrHead(open.length, items.length - open.length) + `<div class="card list">${ms.map((m) => `<div class="irow" data-act="meeting" data-id="${m.id}"><div><b>${esc(fmtLong(m.date))}</b><div class="small muted">${esc((m.notes || '').slice(0, 140))}</div></div><span class="chip soft">${(m.items || []).length} asunto(s)</span></div>`).join('') || '<p class="muted">Sin notas de reunión.</p>'}</div>`;
  }
  let list = state.intrTab === 'abiertos' ? open : items.filter((i) => i.status === 'cerrado');
  if (state.intrStatus && state.intrTab === 'abiertos') list = list.filter((i) => i.status === state.intrStatus);
  if (state.intrTipo) list = list.filter((i) => i.tipo === state.intrTipo);
  list.sort((a, b) => (a.due_date || '9999').localeCompare(b.due_date || '9999') || (a.priority || 3) - (b.priority || 3));
  const counts = Object.fromEntries(Object.keys(INTR_EST).map((k) => [k, open.filter((i) => i.status === k).length]));
  return intrHead(open.length, items.length - open.length) + `
  ${state.intrTab !== 'reuniones' ? `<div class="filters"><button class="fchip ${!state.intrTipo ? 'on' : ''}" data-act="intr-tipo" data-s="">Todos los tipos</button>${Object.entries(INTR_TIPO_C).map(([k, l]) => `<button class="fchip ${state.intrTipo === k ? 'on' : ''}" data-act="intr-tipo" data-s="${k}">${l} <b>${(state.intrTab === 'abiertos' ? open : items.filter((i) => i.status === 'cerrado')).filter((i) => i.tipo === k).length}</b></button>`).join('')}</div>` : ''}
  ${state.intrTab === 'abiertos' ? `<div class="filters"><button class="fchip ${!state.intrStatus ? 'on' : ''}" data-act="intr-status" data-s="">Todos <b>${open.length}</b></button>${Object.entries(INTR_EST).filter(([k]) => k !== 'cerrado').map(([k, l]) => `<button class="fchip ${state.intrStatus === k ? 'on' : ''}" data-act="intr-status" data-s="${k}">${l} <b>${counts[k]}</b></button>`).join('')}</div>` : ''}
  <div class="card list">${list.map((i) => {
    const days = diffDays(i.requested_at || ref, ref);
    const late = i.due_date && i.due_date < ref && i.status !== 'cerrado';
    return `<div class="irow" data-act="intr" data-id="${i.id}">
      <div class="irow-main">
        <div class="trow-title"><span class="pdot p${i.priority || 3}"></span>${i.folio ? `<span class="muted">#${esc(i.folio)}</span> ` : ''}${esc(i.title)}</div>
        <div class="chips">${chip(INTR_TIPO_C[i.tipo] || 'Reporte', i.tipo === 'reporte' ? 'warn' : i.tipo === 'necesidad' ? 'info' : '')}${chip(esc(INTR_EST[i.status]), 'st-' + i.status)}${chip(esc(i.hospital || INTR_ORIGEN[i.origen] || ''), 'soft')}${i.due_date ? chip('Compromiso ' + esc(fmtShort(i.due_date)), late ? 'danger' : 'soft') : ''}${i.status !== 'cerrado' ? chip(`${days} d abierto`, 'soft') : ''}${i.responsable ? chip('Resp.: ' + esc(i.responsable), 'soft') : ''}</div>
        ${i.next_action ? `<div class="small">Siguiente: ${esc(i.next_action)}</div>` : ''}
      </div>
    </div>`;
  }).join('') || '<p class="muted">Sin asuntos.</p>'}</div>`;
}
function intrHead(nOpen, nClosed) {
  return `<section class="day-head"><div class="datenav-t"><h1>Intramed</h1><div class="muted">Requerimientos, incidencias y soporte de los hospitales</div></div>
  <div class="btnrow"><button class="btn primary small" data-act="intr-new">+ Nuevo asunto</button><button class="btn small" data-act="meeting-new">+ Nota de reunión</button></div></section>
  <div class="tabs"><button class="${state.intrTab === 'abiertos' ? 'on' : ''}" data-act="intr-tab" data-t="abiertos">Abiertos (${nOpen})</button><button class="${state.intrTab === 'cerrados' ? 'on' : ''}" data-act="intr-tab" data-t="cerrados">Cerrados (${nClosed})</button><button class="${state.intrTab === 'reuniones' ? 'on' : ''}" data-act="intr-tab" data-t="reuniones">Reuniones</button></div>`;
}

// ================= CIERRE DE JORNADA =================
const CIERRE_Q = ['¿Qué completaste?', '¿Qué quedó pendiente?', '¿Qué bloqueó tu avance?', '¿Qué surgió como urgente?', '¿Qué necesitas reprogramar?', '¿Qué información o decisión falta?', '¿Alguna tarea puede delegarse? (con autorización de la Directora de Calidad)'];
function viewCierre() {
  const c = ctx(); const d = today();
  const existing = store.get('closures', 'closure-' + d);
  if (state.closureResult || (existing && !state.closure)) return closureResultView(state.closureResult || existing.outputs, existing);
  if (!state.closure) {
    const todays = store.all('tasks').filter((t) => t.scheduled_date === d || (t.status === 'completada' && (t.completed_at || '').slice(0, 10) === d));
    state.closure = {
      step: 0, date: d,
      done: todays.filter((t) => t.status === 'completada').map((t) => t.id),
      pending: todays.filter((t) => isOpen(t)).map((t) => t.id),
      blockers: '', urgent: '', missing: '', resched: {}, delegate: [], delegate_to: {},
    };
  }
  const C = state.closure; const step = C.step;
  const todays = store.all('tasks').filter((t) => t.scheduled_date === d || (t.status === 'completada' && (t.completed_at || '').slice(0, 10) === d) || C.done.includes(t.id));
  const openAll = store.all('tasks').filter(isOpen).sort(byScore);
  let body = '';
  if (step === 0) body = checkList(todays.length ? todays : openAll.slice(0, 15), C.done, 'done') + '<p class="small muted">Marca lo que terminaste hoy (también puedes marcar tareas no programadas).</p>';
  if (step === 1) { const pend = todays.filter((t) => !C.done.includes(t.id) && isOpen(t)); C.pending = pend.map((t) => t.id); body = pend.length ? pend.map((t) => taskRow(t)).join('') : '<p class="muted">Nada pendiente de lo programado hoy.</p>'; }
  if (step === 2) body = `${store.all('tasks').filter((t) => isOpen(t) && t.status === 'bloqueada').map((t) => `<div class="small">• ${esc(t.title)}${t.blocked_reason ? ': ' + esc(t.blocked_reason) : ''}</div>`).join('')}<textarea data-cl="blockers" rows="4" placeholder="Ej.: Falta respuesta de Intramed sobre…">${esc(C.blockers)}</textarea>`;
  if (step === 3) body = `${store.all('bomberazos').filter((b) => b.date === d).map((b) => `<div class="small">• ${esc(b.title)}</div>`).join('') || '<p class="small muted">Sin imprevistos registrados hoy.</p>'}<textarea data-cl="urgent" rows="3" placeholder="Algo más que haya surgido…">${esc(C.urgent)}</textarea>`;
  if (step === 4) {
    const next = nextWorkingDay(d, c.S, store);
    const pend = C.pending.map((id) => store.get('tasks', id)).filter(Boolean);
    pend.forEach((t) => { if (!(t.id in C.resched)) C.resched[t.id] = next; });
    body = pend.length ? pend.map((t) => `<div class="resched"><div>${esc(t.title)}</div><select data-resched="${t.id}"><option value="${next}" ${C.resched[t.id] === next ? 'selected' : ''}>Siguiente día hábil (${esc(fmtShort(next))})</option>${[1, 2, 3, 4, 5, 6, 7].map((i) => addDays(d, i)).filter((x) => x !== next && dayInfo(x, c).working).map((x) => `<option value="${x}" ${C.resched[t.id] === x ? 'selected' : ''}>${esc(fmtShort(x))}</option>`).join('')}<option value="none" ${C.resched[t.id] === 'none' ? 'selected' : ''}>Sin fecha (bandeja)</option><option value="keep" ${C.resched[t.id] === 'keep' ? 'selected' : ''}>Dejar como está</option></select></div>`).join('') : '<p class="muted">Nada que reprogramar.</p>';
  }
  if (step === 5) body = `${store.all('tasks').filter((t) => isOpen(t) && t.clasificacion === 'espera').map((t) => `<div class="small">• ${esc(t.title)}${t.waiting_on ? ' — de: ' + esc(t.waiting_on) : ''}</div>`).join('')}<textarea data-cl="missing" rows="3" placeholder="Ej.: Validación del formato por parte de…">${esc(C.missing)}</textarea>`;
  if (step === 6) body = `<p class="small muted">La delegación solo se propone; requiere autorización de la ${esc(c.S.directora_label)}.</p>${openAll.slice(0, 20).map((t) => `<div class="deleg"><label class="cb"><input type="checkbox" data-deleg="${t.id}" ${C.delegate.includes(t.id) ? 'checked' : ''}> ${esc(t.title)}</label>${C.delegate.includes(t.id) ? `<input data-deleg-to="${t.id}" placeholder="¿A quién? (p. ej., gerente de calidad de…)" value="${esc(C.delegate_to[t.id] || t.delegate_to || '')}">` : ''}</div>`).join('')}`;
  return `
  <section class="day-head"><div class="datenav-t"><h1>Cierre del día</h1><div class="muted">${esc(fmtLong(d))} · Paso ${step + 1} de 7</div></div></section>
  <div class="steps">${CIERRE_Q.map((_, i) => `<i class="${i < step ? 'ok' : i === step ? 'cur' : ''}"></i>`).join('')}</div>
  <div class="card"><h2>${esc(CIERRE_Q[step])}</h2>${body}</div>
  <div class="btnrow wizard-nav">
    ${step > 0 ? '<button class="btn" data-act="cl-prev">Anterior</button>' : '<span></span>'}
    ${step < 6 ? '<button class="btn primary" data-act="cl-next">Siguiente</button>' : '<button class="btn primary" data-act="cl-finish">Generar cierre</button>'}
  </div>`;
}
function checkList(tasks, selected, key) {
  return tasks.map((t) => `<label class="cb row"><input type="checkbox" data-clcheck="${key}" value="${t.id}" ${selected.includes(t.id) ? 'checked' : ''}> <span>${esc(t.title)} ${t.scheduled_start ? `<small class="muted">${t.scheduled_start}</small>` : ''}</span></label>`).join('');
}
function closureResultView(o, existing) {
  const block = (title, text, id) => `<div class="card"><div class="card-h"><h3>${title}</h3><button class="btn small" data-act="copy" data-target="${id}">Copiar</button></div><pre id="${id}" class="out">${esc(text)}</pre></div>`;
  return `
  <section class="day-head"><div class="datenav-t"><h1>Cierre del día</h1><div class="muted">${esc(fmtLong(today()))}${existing ? ' · guardado' : ''}</div></div>
  <div class="btnrow"><button class="btn small" data-act="cl-redo">Rehacer cierre</button></div></section>
  ${block('Mensaje para la ' + esc(ctx().S.directora_label), o.mensaje, 'o-msg')}
  ${block('Resumen de lo realizado', o.resumen, 'o-res')}
  ${block('Pendientes reprogramados', o.reprogramados.join('\n') || 'Ninguno.', 'o-rep')}
  ${block('Riesgos para los siguientes días', o.riesgos.join('\n') || 'Sin riesgos detectados.', 'o-rie')}
  ${block('Plan sugerido para el siguiente día hábil', o.plan, 'o-plan')}`;
}

// ================= RESUMEN =================
function viewResumen() {
  const text = periodSummary(state.period, ctx());
  return `
  <section class="day-head"><div class="datenav-t"><h1>Resumen para la ${esc(ctx().S.directora_label)}</h1><div class="muted">Breve, objetivo y listo para copiar. Puedes editarlo antes.</div></div></section>
  <div class="tabs"><button class="${state.period === 'hoy' ? 'on' : ''}" data-act="period" data-p="hoy">Hoy</button><button class="${state.period === 'semana' ? 'on' : ''}" data-act="period" data-p="semana">Últimos 7 días</button></div>
  <div class="card"><textarea id="sum-text" rows="18" class="mono">${esc(text)}</textarea>
  <div class="btnrow"><button class="btn primary" data-act="copy" data-target="sum-text">Copiar</button><button class="btn" data-act="mail-summary">Abrir en correo</button></div></div>`;
}

// ================= MÁS (móvil) =================
function viewMas() {
  return `<section class="day-head"><h1>Más</h1></section>
  <div class="card list menu">
    <a href="#semana" class="mrow">${ico('cal')} Semana</a>
    <a href="#tareas" class="mrow">${ico('list')} Tareas</a>
    <a href="#cierre" class="mrow">${ico('moon')} Cierre del día</a>
    <a href="#resumen" class="mrow">${ico('send')} Resumen para la Directora</a>
    <a href="#ajustes" class="mrow">${ico('gear')} Ajustes, horario y Outlook</a>
  </div>`;
}

// ================= AJUSTES =================
function viewAjustes() {
  const S = ctx().S;
  const theme = localStorage.getItem('mdc:theme') || 'sistema';
  const rec = S.recurring.map((r) => `
    <div class="rec">
      <div class="rec-h"><b>${esc(r.title)}</b> <span class="small muted">${esc(r.note || '')}</span></div>
      <div class="grid4">
        <label>Inicio<input type="time" name="rec.${r.id}.start" value="${r.start}"></label>
        <label>Fin<input type="time" name="rec.${r.id}.end" value="${r.end}"></label>
        <label>Modo<select name="rec.${r.id}.mode"><option value="siempre" ${r.mode === 'siempre' ? 'selected' : ''}>Siempre activa</option><option value="manual" ${r.mode === 'manual' ? 'selected' : ''}>Solo si la activo en el día</option><option value="off" ${r.mode === 'off' ? 'selected' : ''}>Desactivada</option></select></label>
        <div class="days">${[1, 2, 3, 4, 5, 6].map((d) => `<label class="dchk"><input type="checkbox" name="rec.${r.id}.days" value="${d}" ${r.days.includes(d) ? 'checked' : ''}>${DIAS_C[d]}</label>`).join('')}</div>
      </div>
    </div>`).join('');
  const olState = !outlook.configured
    ? `<p>Pendiente de configuración. Requiere que TI registre la aplicación en Microsoft Entra ID y te entregue el <b>ID de aplicación</b> y el <b>ID de directorio</b>; después se capturan en <code>js/config.js</code>. La app funciona completa sin Outlook.</p>`
    : outlook.connected
      ? `<p>Conectado (solo lectura). Última lectura: ${outlook.lastSync ? esc(new Date(outlook.lastSync).toLocaleString('es-MX')) : '—'}</p>${outlook.error ? `<p class="alert warn small">${esc(outlook.error)}</p>` : ''}
         <label class="cb"><input type="checkbox" name="outlook_hide_dupes" ${S.outlook_hide_dupes ? 'checked' : ''}> Si una reunión fija también está en Outlook, mostrar solo la de Outlook</label>
         <div class="btnrow"><button class="btn" type="button" data-act="ol-refresh">Actualizar ahora</button><button class="btn ghost" type="button" data-act="ol-disconnect">Desconectar</button></div>`
      : `<p>Lee tus eventos para detectar horarios ocupados y conflictos. No crea, cambia ni borra eventos. Los eventos no se guardan en la nube.</p><button class="btn primary" type="button" data-act="ol-connect">Conectar Outlook</button>${outlook.error ? `<p class="alert warn small">${esc(outlook.error)}</p>` : ''}`;
  return `
  <section class="day-head"><h1>Ajustes</h1></section>
  <form data-form="settings" class="stack">
    <div class="card"><h3>Jornada</h3>
      <div class="grid4">
        <label>Lun–Vie inicio<input type="time" name="weekday_start" value="${S.weekday_start}"></label>
        <label>Lun–Vie fin<input type="time" name="weekday_end" value="${S.weekday_end}"></label>
        <label>Sábado inicio<input type="time" name="saturday_start" value="${S.saturday_start}"></label>
        <label>Sábado fin<input type="time" name="saturday_end" value="${S.saturday_end}"></label>
        <label>Un sábado laboral de referencia<input type="date" name="saturday_ref" value="${S.saturday_ref}"></label>
        <label>Comida inicio<input type="time" name="lunch_start" value="${S.lunch_start}"></label>
        <label>Comida (min)<select name="lunch_min"><option value="60" ${+S.lunch_min === 60 ? 'selected' : ''}>60</option><option value="30" ${+S.lunch_min === 30 ? 'selected' : ''}>30</option></select></label>
        <label>Margen para imprevistos (%)<input type="number" min="0" max="60" name="buffer_pct" value="${S.buffer_pct}"></label>
        <label>Bloque preferido (min)<input type="number" min="15" step="15" name="block_min" value="${S.block_min}"></label>
        <label>Bloque máximo (min)<input type="number" min="30" step="15" name="block_max" value="${S.block_max}"></label>
        <label>Bloque de foco protegido (min)<input type="number" min="30" step="15" name="focus_min" value="${S.focus_min}"></label>
        <label>Días hábiles de revisión de indicadores<input type="number" min="1" max="10" name="indicadores_dias" value="${S.indicadores_dias}"></label>
        <label>Días de espera para avisar seguimiento de documentos<input type="number" min="1" max="60" name="followup_days" value="${S.followup_days}"></label>
      </div>
      <p class="small muted">Sábados alternos: se calculan a partir del sábado de referencia. Puedes cambiar cualquier día en “Ajustar día”.</p>
    </div>
    <div class="card"><h3>Reuniones y compromisos recurrentes</h3>${rec}</div>
    <div class="card"><h3>Hospitales / unidades</h3>
      <textarea name="hospitals" rows="5" placeholder="Uno por línea, como los quieras ver en la app">${esc((S.hospitals || []).join('\n'))}</textarea>
      <label>Nombre del cargo para los resúmenes<input name="directora_label" value="${esc(S.directora_label)}"></label>
    </div>
    <div class="card"><h3>Outlook (Microsoft 365)</h3>${olState}</div>
    <div class="card"><h3>Apariencia y avisos</h3>
      <div class="seg">${['sistema', 'claro', 'oscuro'].map((t) => `<label><input type="radio" name="theme" value="${t}" ${theme === t ? 'checked' : ''}><span>${t[0].toUpperCase() + t.slice(1)}</span></label>`).join('')}</div>
      <label class="cb"><input type="checkbox" name="notifications" ${S.notifications ? 'checked' : ''}> Avisos del navegador 5 min antes de cada bloque y al final de la jornada (con la app abierta)</label>
    </div>
    <div class="savebar"><button class="btn primary" type="submit">Guardar ajustes</button></div>
  </form>
  <div class="card"><h3>Datos y cuenta</h3>
    <p class="small muted">${store.mode === 'local' ? 'Modo local: los datos están solo en este navegador.' : `Sesión: ${esc(store.user?.email || '')} · Estado: ${esc(store.status)}`}</p>
    <div class="btnrow">
      <button class="btn" data-act="export">Descargar respaldo (JSON)</button>
      <label class="btn">Importar respaldo<input type="file" accept="application/json" data-import hidden></label>
      ${store.mode === 'remote' ? '<button class="btn" data-act="sync-now">Sincronizar ahora</button><button class="btn" data-act="change-pass">Cambiar contraseña</button><button class="btn ghost" data-act="logout">Cerrar sesión</button>' : ''}
    </div>
  </div>`;
}

// ================= HOJAS (formularios) =================
function hospitalsDatalist() {
  const S = ctx().S;
  const fromData = [...new Set([...store.all('tasks').map((t) => t.area), ...store.all('intramed').map((i) => i.hospital)].filter(Boolean))];
  return `<datalist id="dl-hosp">${[...new Set([...(S.hospitals || []), ...fromData])].map((h) => `<option value="${esc(h)}">`).join('')}</datalist>`;
}

// ---- Captura rápida ----
function openCapture(prefill = '', at = null) {
  openSheet('Nueva tarea', `
    <form data-form="capture" class="stack" ${at ? `data-at-date="${at.date}" data-at-start="${at.start}"` : ''}>
      <textarea name="text" rows="2" autofocus placeholder="Escribe como lo dirías: “Intramed reportó un problema en Hospital X, revisar incidencia y dar seguimiento”" data-parse>${esc(prefill)}</textarea>
      <div id="parse-preview" class="parse"></div>
      <div class="btnrow">
        <button class="btn primary grow" type="submit">Guardar</button>
        <button class="btn" type="button" data-act="capture-detail">Revisar detalles</button>
      </div>
      <div class="seg small">
        <label><input type="radio" name="plan" value="auto" ${at ? '' : 'checked'}><span>Programar automático</span></label>
        <label><input type="radio" name="plan" value="none"><span>Dejar sin programar</span></label>
        ${at ? `<label><input type="radio" name="plan" value="at" checked><span>${toHHMM(at.start)}</span></label>` : ''}
      </div>
      <p class="small muted">Enter para guardar · Shift+Enter para nueva línea</p>
    </form>${hospitalsDatalist()}`);
  updateParsePreview();
}
let lastParse = null;
function updateParsePreview() {
  const ta = $('[data-parse]'); const box = $('#parse-preview'); if (!ta || !box) return;
  const txt = ta.value.trim();
  if (!txt) { box.innerHTML = ''; lastParse = null; return; }
  const answers = {}; $$('[data-q]', box).forEach((i) => { answers[i.dataset.q] = i.value; });
  const r = parseTask(txt, { hospitals: ctx().S.hospitals });
  lastParse = r;
  box.innerHTML = `
    <div class="parse-title">${esc(r.title)}</div>
    <div class="chips">${chip('P' + r.priority + ' ' + PRIO[r.priority], 'p' + r.priority)}${chip(esc(r.category), 'soft')}${chip(esc(CLASIFICACION[r.clasificacion]), 'soft')}${chip(fmtDur(r.duration_min) + (r.duration_estimated ? ' (estimado)' : ''), 'soft')}${r.due_date ? chip('Límite ' + esc(fmtShort(r.due_date)), 'warn') : ''}${r.area ? chip(esc(r.area), 'soft') : ''}${r.riesgo ? chip('Riesgo / cumplimiento', 'danger') : ''}</div>
    ${r.deliverable ? `<div class="small"><span class="muted">Entregable sugerido:</span> ${esc(r.deliverable)}</div>` : ''}
    ${r.next_action ? `<div class="small"><span class="muted">Siguiente acción:</span> ${esc(r.next_action)}</div>` : ''}
    ${r.done_criteria ? `<div class="small"><span class="muted">Criterio de cierre sugerido:</span> ${esc(r.done_criteria)}</div>` : ''}
    ${r.questions.length ? `<div class="qs">${r.questions.map((q) => `<label class="small">${esc(q.q)}${q.field === 'due_date' ? `<input type="date" data-q="due_date" value="${esc(answers.due_date || '')}">` : `<input data-q="${q.field}" ${q.field === 'area' ? 'list="dl-hosp"' : ''} value="${esc(answers[q.field] || '')}" placeholder="Opcional">`}</label>`).join('')}</div>` : ''}
    ${r.intramed ? `<label class="cb small"><input type="checkbox" name="to_intramed" checked> Registrar también en seguimiento de Intramed (${esc(INTR_TIPO[r.intramed_tipo])})</label>` : ''}
    ${r.documental ? '<label class="cb small"><input type="checkbox" name="to_doc"> Dar seguimiento en Avance documental (etapas y rondas de revisión)</label>' : ''}`;
}
function taskFromParse(r, answers = {}) {
  const t = {
    title: r.title, notes: r.notes, due_date: answers.due_date || r.due_date || null, priority: r.priority,
    clasificacion: r.clasificacion, status: 'pendiente', duration_min: r.duration_min, area: answers.area || r.area || '',
    category: r.category, riesgo: r.riesgo, impacto: r.impacto, deliverable: r.deliverable, next_action: answers.next_action || r.next_action,
    done_criteria: r.done_criteria, waiting_on: answers.waiting_on || '', source: 'texto', programa: r.programa || '',
  };
  if (t.area && !t.title.includes(t.area)) t.title += ' · ' + t.area;
  return t;
}
function saveCapture(form, detail = false) {
  const txt = form.text.value.trim(); if (!txt) return;
  const r = parseTask(txt, { hospitals: ctx().S.hospitals });
  const answers = {}; $$('[data-q]', form).forEach((i) => { if (i.value) answers[i.dataset.q] = i.value; });
  const t = taskFromParse(r, answers);
  const toIntr = form.to_intramed ? form.to_intramed.checked : !!r.intramed;
  if (detail) { openTask(null, { ...t, _toIntr: toIntr }); return; }
  const saved = store.put('tasks', t);
  if (toIntr) linkIntramed(saved, r.intramed_tipo);
  if (form.to_doc && form.to_doc.checked) SEG.docFromTask(saved, txt);
  const plan = form.querySelector('[name=plan]:checked')?.value || 'auto';
  closeSheet();
  if (plan === 'at') { scheduleAt(saved.id, form.dataset.atDate, +form.dataset.atStart); toast('Tarea agregada en ' + toHHMM(+form.dataset.atStart)); return; }
  if (plan === 'auto' && !['espera', 'referencia'].includes(saved.clasificacion)) {
    const p = autoPlan([saved], ctx(), { days: 10 });
    if (p.length && !p[0].unplaced) {
      applyPlan(p, store);
      toast(`Programada: ${fmtShort(p[0].date)} ${p[0].start}${p.length > 1 ? ` (en ${p.filter((x) => !x.unplaced).length} bloques)` : ''}`, { label: 'Deshacer', fn: () => { const x = store.get('tasks', saved.id); store.put('tasks', { ...x, scheduled_date: null, scheduled_start: null, status: 'pendiente' }); } });
      return;
    }
    toast('Guardada sin horario: no hay espacio libre en los próximos días.');
    return;
  }
  toast('Tarea guardada en la bandeja');
}
function linkIntramed(task, tipo = 'reporte') {
  const i = store.put('intramed', {
    tipo, origen: task.area ? 'hospital' : 'calidad', title: task.title.replace(/^Intramed:\s*/, ''), description: task.notes || '', hospital: task.area || '',
    requested_at: today(), due_date: task.due_date || null, responsable: '', status: 'recibido', next_action: task.next_action || '',
    priority: task.priority, log: [{ at: new Date().toISOString(), text: 'Registrado desde captura rápida' }],
  });
  store.put('tasks', { ...store.get('tasks', task.id), intramed_id: i.id });
  return i;
}

// ---- Hoja de tarea ----
function openTask(id, draft = null) {
  const t = id ? store.get('tasks', id) : { status: 'pendiente', priority: 3, clasificacion: 'programable', duration_min: 30, category: 'General', ...draft };
  if (!t) return;
  const c = ctx();
  const opt = (obj, v) => Object.entries(obj).map(([k, l]) => `<option value="${k}" ${v === k ? 'selected' : ''}>${esc(l)}</option>`).join('');
  const nextWD = nextWorkingDay(today(), c.S, store);
  const ws = weekStart(today());
  const moveDays = [today(), nextWD, ...[0, 1, 2, 3, 4, 5].map((i) => addDays(ws, i)), ...[0, 1, 2, 3, 4].map((i) => addDays(ws, 7 + i))].filter((x, i, a) => a.indexOf(x) === i && x >= today() && dayInfo(x, c).working);
  const intr = t.intramed_id ? store.get('intramed', t.intramed_id) : null;
  openSheet(id ? 'Tarea' : 'Nueva tarea', `
  <form data-form="task" data-id="${id || ''}" class="stack">
    <input name="title" value="${esc(t.title || '')}" placeholder="Título" required class="title-input" ${id ? '' : 'autofocus'}>
    ${id ? `<div class="btnrow wrap">
      <button class="btn small ${t.status === 'completada' ? '' : 'primary'}" type="button" data-act="toggle-done" data-id="${t.id}">${t.status === 'completada' ? 'Reabrir' : '✓ Completar'}</button>
      ${t.status !== 'en_curso' && isOpen(t) ? `<button class="btn small" type="button" data-act="t-start" data-id="${t.id}">Iniciar</button>` : ''}
      <button class="btn small" type="button" data-act="t-postpone" data-id="${t.id}" data-to="${nextWD}">Posponer a ${esc(fmtShort(nextWD))}</button>
      <button class="btn small" type="button" data-act="t-postpone" data-id="${t.id}" data-to="${addDays(ws, 7)}">Próxima semana</button>
      <button class="btn small" type="button" data-act="t-suggest" data-id="${t.id}">Sugerir horario</button>
      <button class="btn small" type="button" data-act="t-split" data-id="${t.id}">Dividir</button>
      ${t.scheduled_start ? `<button class="btn small" type="button" data-act="t-unsched" data-id="${t.id}">Quitar de agenda</button>` : ''}
    </div>` : ''}
    ${id ? `<div class="movebox"><div class="small muted">Mover a</div><div class="btnrow wrap">${moveDays.map((d) => `<button class="btn small ${t.scheduled_date === d ? 'on' : ''}" type="button" data-act="t-move" data-id="${t.id}" data-d="${d}">${d === today() ? 'Hoy' : esc(fmtShort(d))}</button>`).join('')}</div></div>` : ''}
    <div class="grid2">
      <label>Fecha programada<input type="date" name="scheduled_date" value="${t.scheduled_date || ''}"></label>
      <label>Hora de inicio<input type="time" step="900" name="scheduled_start" value="${t.scheduled_start || ''}"></label>
      <label>Fecha límite<input type="date" name="due_date" value="${t.due_date || ''}"></label>
      <label>Duración (min)<input type="number" min="5" step="5" name="duration_min" value="${t.duration_min || 30}" list="dl-dur"></label>
      <label>Prioridad<select name="priority">${[1, 2, 3, 4].map((p) => `<option value="${p}" ${+t.priority === p ? 'selected' : ''}>P${p} · ${PRIO[p]}</option>`).join('')}</select></label>
      <label>Clasificación<select name="clasificacion">${opt(CLASIFICACION, t.clasificacion)}</select></label>
      <label>Estado<select name="status">${opt(ESTADOS, t.status)}</select></label>
      <label>Categoría<select name="category">${CATEGORIAS.map((x) => `<option ${t.category === x ? 'selected' : ''}>${x}</option>`).join('')}</select></label>
      <label>Área / hospital<input name="area" list="dl-hosp" value="${esc(t.area || '')}"></label>
      <label>Impacto<select name="impacto"><option value="alto" ${t.impacto === 'alto' ? 'selected' : ''}>Alto</option><option value="medio" ${t.impacto !== 'alto' && t.impacto !== 'bajo' ? 'selected' : ''}>Medio</option><option value="bajo" ${t.impacto === 'bajo' ? 'selected' : ''}>Bajo</option></select></label>
    </div>
    <label class="cb"><input type="checkbox" name="riesgo" ${t.riesgo ? 'checked' : ''}> Riesgo para el paciente, seguridad, cumplimiento normativo o certificación</label>
    <label class="cb"><input type="checkbox" name="unblocks" ${t.unblocks ? 'checked' : ''}> Desbloquea el avance de otros</label>
    <label>Siguiente acción<input name="next_action" value="${esc(t.next_action || '')}"></label>
    <label>Entregable<input name="deliverable" value="${esc(t.deliverable || '')}"></label>
    <label>Criterio de cierre<input name="done_criteria" value="${esc(t.done_criteria || '')}"></label>
    <div class="grid2">
      <label>Depende de (info / decisión)<input name="waiting_on" value="${esc(t.waiting_on || '')}"></label>
      <label>Motivo de bloqueo<input name="blocked_reason" value="${esc(t.blocked_reason || '')}"></label>
      <label>Delegación propuesta a<input name="delegate_to" value="${esc(t.delegate_to || '')}" placeholder="Gerente de calidad…"></label>
      <label>Estatus de delegación<select name="delegation_status"><option value="">—</option><option value="propuesta" ${t.delegation_status === 'propuesta' ? 'selected' : ''}>Propuesta (por autorizar)</option><option value="autorizada" ${t.delegation_status === 'autorizada' ? 'selected' : ''}>Autorizada por la Directora</option><option value="no_autorizada" ${t.delegation_status === 'no_autorizada' ? 'selected' : ''}>No autorizada</option></select></label>
    </div>
    <label>Seguimiento (opcional)<select name="project_id">${SEG.projectOptions(t.project_id)}</select></label>
    ${t.doc_id && store.get('docs', t.doc_id) ? `<p class="small">Documento: <a href="#" data-act="doc" data-id="${t.doc_id}">${esc(store.get('docs', t.doc_id).name)}</a></p>` : ''}
    <label>Notas<textarea name="notes" rows="3">${esc(t.notes || '')}</textarea></label>
    ${intr ? `<p class="small">Vinculada a Intramed: <a href="#" data-act="intr" data-id="${intr.id}">${esc(intr.title)}</a> (${esc(INTR_EST[intr.status])})</p>` : ''}
    ${!id && draft?._toIntr ? '<label class="cb"><input type="checkbox" name="to_intramed" checked> Registrar también en Intramed</label>' : ''}
    <div class="btnrow savebar">
      <button class="btn primary grow" type="submit">Guardar</button>
      ${id ? `<button class="btn ghost danger-t" type="button" data-act="t-delete" data-id="${t.id}">Eliminar</button>` : ''}
    </div>
  </form>${hospitalsDatalist()}<datalist id="dl-dur">${[15, 30, 45, 60, 90, 120].map((d) => `<option value="${d}">`).join('')}</datalist>`, { wide: true });
}
function readTaskForm(f, base) {
  const fd = new FormData(f);
  const t = { ...base };
  ['title', 'scheduled_date', 'scheduled_start', 'due_date', 'clasificacion', 'status', 'category', 'area', 'impacto', 'next_action', 'deliverable', 'done_criteria', 'waiting_on', 'blocked_reason', 'delegate_to', 'delegation_status', 'notes'].forEach((k) => { t[k] = (fd.get(k) || '').toString().trim() || (['scheduled_date', 'scheduled_start', 'due_date'].includes(k) ? null : ''); });
  t.priority = +fd.get('priority'); t.duration_min = +fd.get('duration_min') || 30;
  t.riesgo = !!fd.get('riesgo'); t.unblocks = !!fd.get('unblocks'); t.project_id = fd.get('project_id') || null;
  if (!t.scheduled_date) t.scheduled_start = null;
  if (t.scheduled_start && t.status === 'pendiente') t.status = 'programada';
  if (!t.scheduled_start && t.status === 'programada') t.status = 'pendiente';
  if (t.status === 'completada' && !t.completed_at) t.completed_at = new Date().toISOString();
  if (t.status !== 'completada') t.completed_at = null;
  if (t.delegate_to && !t.delegation_status) t.delegation_status = 'propuesta';
  delete t._toIntr;
  return t;
}

function scheduleAt(id, date, startMin) {
  const t = store.get('tasks', id); if (!t) return;
  store.put('tasks', { ...t, scheduled_date: date, scheduled_start: toHHMM(startMin), status: ['pendiente', 'pospuesta'].includes(t.status) ? 'programada' : t.status });
}
function moveToDay(id, date) {
  const t = store.get('tasks', id); if (!t) return;
  const c = ctx(); const info = dayInfo(date, c);
  if (!info.working) { toast('Ese día no es laboral.'); return; }
  const others = { ...c, store: { ...store, all: (k) => store.all(k).filter((x) => k !== 'tasks' || x.id !== id), get: store.get.bind(store) } };
  const p = autoPlan([{ ...t, scheduled_date: null, scheduled_start: null }], others, { from: date, days: 1, useBuffer: true }).filter((x) => !x.unplaced && x.date === date);
  if (p.length && p[0].dur >= (t.duration_min || 30)) { scheduleAt(id, date, toMin(p[0].start)); toast(`Movida a ${fmtShort(date)} ${p[0].start}`); }
  else {
    const free = info.free.filter((f) => date !== today() || f.e > nowMin());
    const s = free.length ? Math.max(free[0].s, date === today() ? roundUp(nowMin(), 15) : 0) : info.win.s;
    scheduleAt(id, date, s);
    toast(`Movida a ${fmtShort(date)} ${toHHMM(s)}: no cabe completa, revisa el conflicto.`);
  }
}

// ---- Imprevisto (antes “bomberazo”) ----
function openBomberazo() {
  const open = store.all('tasks').filter((t) => isOpen(t) && !t.is_urgent).sort(byScore).slice(0, 30);
  openSheet('Imprevisto', `
  <form data-form="bomberazo" class="stack">
    <label>¿Qué ocurrió?<textarea name="what" rows="2" required autofocus placeholder="Ej.: Hospital X no puede imprimir recetas en Intramed"></textarea></label>
    <div><div class="lbl">¿Cuánto tiempo tomará?</div><div class="seg">${[15, 30, 60, 90, 120].map((m) => `<label><input type="radio" name="dur" value="${m}" ${m === 30 ? 'checked' : ''}><span>${fmtDur(m)}</span></label>`).join('')}</div></div>
    <div><div class="lbl">¿Tiene fecha límite?</div><div class="seg"><label><input type="radio" name="due" value="hoy" checked><span>Hoy</span></label><label><input type="radio" name="due" value="manana"><span>Mañana</span></label><label><input type="radio" name="due" value="none"><span>Sin fecha</span></label></div></div>
    <label>¿Bloquea otra tarea?<select name="blocks"><option value="">No</option>${open.map((t) => `<option value="${t.id}">${esc(t.title)}</option>`).join('')}</select></label>
    <label>¿Requiere a alguien más?<input name="requires" placeholder="Ej.: Soporte Intramed, gerente de calidad…"></label>
    <label>Área / hospital<input name="area" list="dl-hosp"></label>
    <label class="cb"><input type="checkbox" name="intramed"> Es un tema de Intramed (registrar en seguimiento)</label>
    <button class="btn fire block" type="submit">Reorganizar mi día</button>
    <p class="small muted">No se elimina ninguna tarea. Podrás aceptar, ajustar o revertir.</p>
  </form>${hospitalsDatalist()}`);
}
function runBomberazo(f) {
  const fd = new FormData(f);
  const what = fd.get('what').trim();
  const r = parseTask(what, { hospitals: ctx().S.hospitals });
  const dueSel = fd.get('due');
  const input = {
    title: r.title || what.slice(0, 80), notes: what, duration_min: +fd.get('dur'),
    due_date: dueSel === 'hoy' ? today() : dueSel === 'manana' ? addDays(today(), 1) : null,
    blocks_task_id: fd.get('blocks') || null, requires: fd.get('requires').trim(), area: fd.get('area').trim() || r.area,
    category: fd.get('intramed') ? 'Intramed' : r.category, riesgo: r.riesgo, next_action: r.next_action,
  };
  if (input.area && !input.title.includes(input.area)) input.title += ' · ' + input.area;
  const res = bomberazo(input, ctx());
  const urgent = store.put('tasks', res.urgent);
  if (fd.get('intramed')) linkIntramed(urgent, r.intramed_tipo);
  applyChanges(res.changes, store, 'after');
  const rec = store.put('bomberazos', { date: res.date, at: new Date().toISOString(), title: urgent.title, task_id: urgent.id, duration_min: input.duration_min, requires: input.requires, blocks_task_id: input.blocks_task_id, changes: res.changes, state: 'aplicado' });
  state.date = res.date; location.hash = '#hoy';
  showBzResult(rec, res);
}
function showBzResult(rec, res) {
  const moved = rec.changes.filter((c) => c.after.status !== 'pospuesta' && c.after.status !== 'bloqueada');
  const post = rec.changes.filter((c) => c.after.status === 'pospuesta');
  const blk = rec.changes.filter((c) => c.after.status === 'bloqueada');
  const u = store.get('tasks', rec.task_id);
  openSheet('Reorganización propuesta', `
    <div class="stack">
      ${res?.outside ? '<p class="alert warn small">Fuera de horario laboral: el imprevisto se programó al inicio del siguiente día hábil.</p>' : ''}
      <div class="alert fire"><b>Imprevisto:</b> ${esc(rec.title)}<br>${u ? `${esc(fmtShort(u.scheduled_date))} ${u.scheduled_start}, ${fmtDur(u.duration_min)}` : ''}</div>
      <h3>Se movió (${moved.length})</h3>${moved.map((c) => `<div class="small">• <b>${esc(c.title)}</b> — ${esc(c.reason)}</div>`).join('') || '<p class="small muted">Nada.</p>'}
      <h3>Se pospuso (${post.length})</h3>${post.map((c) => `<div class="small">• <b>${esc(c.title)}</b> — ${esc(c.reason)}</div>`).join('') || '<p class="small muted">Nada.</p>'}
      ${blk.length ? `<h3>Queda bloqueada</h3>${blk.map((c) => `<div class="small">• <b>${esc(c.title)}</b> — ${esc(c.reason)}</div>`).join('')}` : ''}
      <div class="btnrow">
        <button class="btn primary grow" data-act="bz-accept" data-id="${rec.id}">Aceptar</button>
        <button class="btn" data-act="close">Ajustar manualmente</button>
        <button class="btn ghost" data-act="bz-revert" data-id="${rec.id}">Revertir</button>
      </div>
    </div>`);
}

// ---- Día: ajustes, eventos ----
function openDaySettings() {
  const c = ctx(); const d = state.date; const D = dayRec(store, d) || {};
  const wd = dow(d);
  const recs = c.S.recurring.filter((r) => r.mode !== 'off' && r.days.includes(wd));
  const info = dayInfo(d, c);
  const events = store.all('events').filter((e) => e.date === d);
  openSheet('Ajustar ' + fmtShort(d), `
  <form data-form="day" class="stack">
    <label>Tipo de día<select name="tipo">${Object.entries(TIPOS_DIA).map(([k, l]) => `<option value="${k}" ${(D.tipo || info.tipo) === k ? 'selected' : ''}>${l}</option>`).join('')}</select></label>
    <div class="grid2"><label>Inicio (jornada especial)<input type="time" name="start" value="${D.start || ''}"></label><label>Fin (jornada especial)<input type="time" name="end" value="${D.end || ''}"></label></div>
    <label>Nota (p. ej., hospital que visitas)<input name="note" value="${esc(D.note || '')}"></label>
    <label>Comida<select name="lunch_min"><option value="">Normal (${c.S.lunch_min} min)</option><option value="30" ${+D.lunch_min === 30 ? 'selected' : ''}>Reducida a 30 min</option><option value="60" ${+D.lunch_min === 60 ? 'selected' : ''}>60 min</option></select></label>
    ${recs.length ? `<div class="lbl">Reuniones de este día</div>${recs.map((r) => {
      const on = D.rec?.[r.id] ?? (r.mode === 'siempre' && r.days.includes(wd));
      return `<label class="cb"><input type="checkbox" name="rec_${r.id}" ${on ? 'checked' : ''}> ${esc(r.title)} <small class="muted">${r.start}–${r.end}</small></label>`;
    }).join('')}` : ''}
    <label>La reunión con la Directora General se extendió hasta<input type="time" name="dg_end" value="${D.dg_end || ''}"></label>
    <button class="btn primary block" type="submit">Guardar día</button>
  </form>
  ${events.length ? `<h3>Reuniones imprevistas</h3>${events.map((e) => `<div class="irow"><div>${e.start}–${e.end} ${esc(e.title)}</div><button class="btn small ghost" data-act="ev-del" data-id="${e.id}">Quitar</button></div>`).join('')}` : ''}`);
}
function openEvent(date, s = null, e = null) {
  const st = s ?? roundUp(nowMin(), 15);
  openSheet('Reunión imprevista / bloqueo', `
  <form data-form="event" class="stack" data-date="${date}">
    <label>Título<input name="title" required autofocus placeholder="Reunión con…"></label>
    <div class="grid2"><label>Inicio<input type="time" name="start" value="${toHHMM(st)}" required></label><label>Fin<input type="time" name="end" value="${toHHMM(e ?? st + 60)}" required></label></div>
    <label>Tipo<select name="kind"><option value="reunion">Reunión imprevista</option><option value="bloqueo">Bloqueo / ocupado</option><option value="visita">Visita a hospital</option></select></label>
    <label class="cb"><input type="checkbox" name="reflow" checked> Reacomodar las tareas que choquen</label>
    <button class="btn primary block">Agregar</button>
  </form>`);
}
function reflowAfterEvent(date) {
  // Mueve las tareas abiertas que chocan con bloques fijos a otro espacio
  const c = ctx(); const info = dayInfo(date, c);
  const clashing = info.tasks.filter((b) => isOpen(b.task) && info.conflicts[b.id] && b.task.status !== 'en_curso').map((b) => b.task);
  if (!clashing.length) return 0;
  const masked = { ...c, store: { ...store, all: (k) => store.all(k).filter((x) => k !== 'tasks' || !clashing.some((t) => t.id === x.id)), get: store.get.bind(store) } };
  const p = autoPlan(clashing.map((t) => ({ ...t, scheduled_date: null, scheduled_start: null })), masked, { from: date, days: 10, useBuffer: true });
  applyPlan(p, store);
  return clashing.length;
}
function openSlot(s, e) {
  const unsched = store.all('tasks').filter(floating).sort(byScore);
  const len = e - s;
  openSheet(`Espacio libre ${toHHMM(s)}–${toHHMM(e)}`, `
    <div class="stack">
      <div class="btnrow"><button class="btn primary" data-act="new-at" data-s="${s}">+ Nueva tarea aquí</button><button class="btn" data-act="event-at" data-s="${s}" data-e="${e}">+ Reunión aquí</button></div>
      <h3>Colocar una tarea sin programar</h3>
      ${unsched.length ? unsched.slice(0, 15).map((t) => `<div class="irow"><div>${esc(t.title)} <small class="muted">${fmtDur(t.duration_min || 30)}${(t.duration_min || 30) > len ? ' · no cabe completa' : ''}</small></div><button class="btn small" data-act="place" data-id="${t.id}" data-s="${s}">Colocar</button></div>`).join('') : '<p class="muted small">No hay tareas sin programar.</p>'}
    </div>`);
}
function openFixed(id, date = state.date) {
  const c = ctx(); const info = dayInfo(date, c);
  const f = info.fixed.find((x) => x.id === id); if (!f) return;
  openSheet(f.title, `<div class="stack">
    <p><b>${toHHMM(f.s)}–${toHHMM(f.e)}</b> · ${esc(fmtLong(date))}</p>
    ${f.kind === 'outlook' ? '<p class="small muted">Evento de Outlook (solo lectura). Se modifica en Outlook.</p>' : ''}
    ${f.rec ? `<p class="small muted">Reunión recurrente configurada.</p><div class="btnrow">${f.rec === 'dg' ? `<button class="btn" data-act="dg-extend" data-d="${date}" data-m="30">Se extendió 30 min</button><button class="btn" data-act="dg-extend" data-d="${date}" data-m="60">Se extendió 1 h</button>` : ''}<button class="btn ghost" data-act="rec-off" data-d="${date}" data-r="${f.rec}">No hay esta reunión este día</button></div>` : ''}
    ${f.event ? `<button class="btn ghost" data-act="ev-del" data-id="${f.event.id}">Quitar</button>` : ''}
    ${f.kind === 'comida' ? `<button class="btn" data-act="lunch-short" data-d="${date}">Reducir comida a 30 min este día</button>` : ''}
  </div>`);
}

// ---- Sugerencias ----
function openSuggest(from, onlyWeek = false) {
  const c = ctx();
  const list = store.all('tasks').filter((t) => floating(t) && t.source !== 'bloque_diario' && !['espera'].includes(t.clasificacion) && t.status !== 'bloqueada');
  const props = autoPlan(list, c, { from, days: onlyWeek ? 6 : 10 });
  window._props = props;
  const placed = props.filter((p) => !p.unplaced);
  const un = props.filter((p) => p.unplaced);
  openSheet('Horarios sugeridos', `<div class="stack">
    <p class="small muted">Orden: riesgo y cumplimiento → fecha límite → impacto → dependencias → tareas rápidas. Respeta reuniones fijas, comida, margen de ${c.S.buffer_pct}% y el bloque de foco.</p>
    ${placed.map((p, i) => `<label class="cb row"><input type="checkbox" data-prop="${i}" checked> <span><b>${esc(fmtShort(p.date))} ${p.start}</b> · ${esc(p.task.title)}${p.parts > 1 ? ` (${p.part}/${p.parts})` : ''} <small class="muted">${fmtDur(p.dur)}</small>${p.late ? ' <span class="chip danger">después de la fecha límite</span>' : ''}</span></label>`).join('') || '<p class="muted">No hay espacio disponible.</p>'}
    ${un.length ? `<p class="small alert warn">Sin espacio en los próximos días: ${un.map((p) => esc(p.task.title)).join(', ')}</p>` : ''}
    ${placed.length ? '<button class="btn primary block" data-act="apply-props">Aplicar seleccionadas</button>' : ''}
  </div>`, { wide: true });
}

// ---- Intramed ----
function openIntramed(id, draft = {}) {
  const i = id ? store.get('intramed', id) : { tipo: 'reporte', origen: 'hospital', status: 'recibido', requested_at: today(), priority: 3, log: [], ...draft };
  if (!i) return;
  const tasks = store.all('tasks').filter((t) => t.intramed_id === i.id);
  openSheet(id ? 'Asunto Intramed' : 'Nuevo asunto Intramed', `
  <form data-form="intramed" data-id="${id || ''}" class="stack">
    <input name="title" class="title-input" required value="${esc(i.title || '')}" placeholder="Título del requerimiento o incidencia" ${id ? '' : 'autofocus'}>
    <div class="grid2">
      <label>Tipo<select name="tipo">${Object.entries(INTR_TIPO).map(([k, l]) => `<option value="${k}" ${i.tipo === k ? 'selected' : ''}>${l}</option>`).join('')}</select></label>
      <label>Origen<select name="origen">${Object.entries(INTR_ORIGEN).map(([k, l]) => `<option value="${k}" ${(i.origen || 'hospital') === k ? 'selected' : ''}>${l}</option>`).join('')}</select></label>
      <label>Estatus<select name="status">${Object.entries(INTR_EST).map(([k, l]) => `<option value="${k}" ${i.status === k ? 'selected' : ''}>${l}</option>`).join('')}</select></label>
      <label>Hospital o área solicitante<input name="hospital" placeholder="Vacío si es una necesidad de Calidad" list="dl-hosp" value="${esc(i.hospital || '')}"></label>
      <label>Folio Intramed (opcional)<input name="folio" value="${esc(i.folio || '')}"></label>
      <label>Fecha de solicitud<input type="date" name="requested_at" value="${i.requested_at || ''}"></label>
      <label>Fecha límite / compromiso<input type="date" name="due_date" value="${i.due_date || ''}"></label>
      <label>Responsable<input name="responsable" value="${esc(i.responsable || '')}"></label>
      <label>Prioridad<select name="priority">${[1, 2, 3, 4].map((p) => `<option value="${p}" ${+i.priority === p ? 'selected' : ''}>P${p} · ${PRIO[p]}</option>`).join('')}</select></label>
    </div>
    <label>Siguiente acción<input name="next_action" value="${esc(i.next_action || '')}"></label>
    <label>Descripción<textarea name="description" rows="3">${esc(i.description || '')}</textarea></label>
    <label>Agregar seguimiento<textarea name="newlog" rows="2" placeholder="Qué se hizo, qué respondió Intramed, qué sigue…"></textarea></label>
    <div class="btnrow savebar"><button class="btn primary grow" type="submit">Guardar</button>${id ? `<button class="btn" type="button" data-act="intr-task" data-id="${i.id}">Crear tarea de seguimiento</button><button class="btn ghost danger-t" type="button" data-act="intr-del" data-id="${i.id}">Eliminar</button>` : ''}</div>
  </form>
  ${tasks.length ? `<h3>Tareas vinculadas</h3>${tasks.map((t) => taskRow(t)).join('')}` : ''}
  ${(i.log || []).length ? `<h3>Historial</h3><div class="log">${[...i.log].reverse().map((l) => `<div><span class="small muted">${esc(new Date(l.at).toLocaleString('es-MX', { dateStyle: 'short', timeStyle: 'short' }))}</span> ${esc(l.text)}</div>`).join('')}</div>` : ''}
  ${hospitalsDatalist()}`, { wide: true });
}
function openMeeting(id) {
  const m = id ? store.get('intramed_meetings', id) : { date: today(), notes: '', acuerdos: '', items: [] };
  const open = store.all('intramed').filter((i) => i.status !== 'cerrado' || (m.items || []).includes(i.id));
  openSheet(id ? 'Nota de reunión Intramed' : 'Nueva nota de reunión Intramed', `
  <form data-form="meeting" data-id="${id || ''}" class="stack">
    <label>Fecha<input type="date" name="date" value="${m.date}"></label>
    <label>Notas<textarea name="notes" rows="4">${esc(m.notes || '')}</textarea></label>
    <label>Acuerdos<textarea name="acuerdos" rows="3">${esc(m.acuerdos || '')}</textarea></label>
    <div class="lbl">Asuntos revisados (se agrega una línea a su historial)</div>
    ${open.map((i) => `<label class="cb"><input type="checkbox" name="items" value="${i.id}" ${(m.items || []).includes(i.id) ? 'checked' : ''}> ${esc(i.title)} <small class="muted">${esc(INTR_EST[i.status])}</small></label>`).join('') || '<p class="small muted">Sin asuntos abiertos.</p>'}
    <div class="btnrow savebar"><button class="btn primary grow">Guardar</button>${id ? `<button class="btn ghost danger-t" type="button" data-act="meeting-del" data-id="${id}">Eliminar</button>` : ''}</div>
  </form>`, { wide: true });
}

// ================= EVENTOS =================
function bindEvents() {
  document.addEventListener('click', onClick);
  document.addEventListener('submit', onSubmit);
  document.addEventListener('input', onInput);
  document.addEventListener('change', onChange);
  document.addEventListener('keydown', onKey);
  // Arrastrar y soltar (computadora)
  document.addEventListener('dragstart', (e) => {
    const el = e.target.closest?.('[data-drag-task]'); if (!el) return;
    e.dataTransfer.setData('text/plain', el.dataset.dragTask); e.dataTransfer.effectAllowed = 'move';
    el.classList.add('dragging');
    const r = el.getBoundingClientRect(); dragOffset = e.clientY - r.top;
  });
  document.addEventListener('dragend', (e) => { e.target.classList?.remove('dragging'); $$('.drop-over').forEach((x) => x.classList.remove('drop-over')); const g = $('.drop-ghost'); if (g) g.hidden = true; });
  document.addEventListener('dragover', (e) => {
    const z = e.target.closest?.('[data-drop-day],[data-drop-timeline],[data-drop-unsched]'); if (!z) return;
    e.preventDefault(); e.dataTransfer.dropEffect = 'move';
    $$('.drop-over').forEach((x) => x !== z && x.classList.remove('drop-over')); z.classList.add('drop-over');
    if (z.dataset.dropTimeline) {
      const m = timelineMinute(z, e.clientY - dragOffset);
      const g = $('.drop-ghost', z); g.hidden = false; g.style.top = ((m - +z.dataset.s0) * PX) + 'px'; g.textContent = toHHMM(m);
    }
  });
  document.addEventListener('drop', (e) => {
    const z = e.target.closest?.('[data-drop-day],[data-drop-timeline],[data-drop-unsched]'); if (!z) return;
    e.preventDefault();
    const id = e.dataTransfer.getData('text/plain'); if (!id) return;
    if (z.dataset.dropTimeline) { scheduleAt(id, state.date, timelineMinute(z, e.clientY - dragOffset)); }
    else if (z.dataset.dropDay) moveToDay(id, z.dataset.dropDay);
    else { const t = store.get('tasks', id); store.put('tasks', { ...t, scheduled_date: null, scheduled_start: null, status: t.status === 'programada' ? 'pendiente' : t.status }); }
  });
}
let dragOffset = 0;
function timelineMinute(z, clientY) {
  const r = z.getBoundingClientRect();
  return Math.max(+z.dataset.s0, Math.round(((clientY - r.top) / PX + +z.dataset.s0) / 15) * 15);
}

function onKey(e) {
  if (e.target.matches('[data-parse]') && e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); saveCapture(e.target.form); return; }
  if (e.key === 'Escape' && !$('#sheet').hidden) { closeSheet(); return; }
  if (e.target.matches('input,textarea,select') || e.metaKey || e.ctrlKey || e.altKey) return;
  if (!$('#sheet') || !$('#sheet').hidden) return;
  if (e.key === 'n') { e.preventDefault(); openCapture(); }
  if (e.key === 'i' || e.key === 'u') { e.preventDefault(); openBomberazo(); }
}

let parseT = null;
function onInput(e) {
  if (e.target.matches('[data-parse]')) { clearTimeout(parseT); parseT = setTimeout(updateParsePreview, 120); }
  if (e.target.dataset.input === 'docq') { segState.docQ = e.target.value; const pos = e.target.selectionStart; render(); const s = $('.search'); if (s) { s.focus(); s.setSelectionRange(pos, pos); } }
  if (e.target.dataset.input === 'q') { state.q = e.target.value; const pos = e.target.selectionStart; render(); const s = $('.search'); if (s) { s.focus(); s.setSelectionRange(pos, pos); } }
  if (e.target.dataset.cl && state.closure) state.closure[e.target.dataset.cl] = e.target.value;
  if (e.target.dataset.delegTo && state.closure) state.closure.delegate_to[e.target.dataset.delegTo] = e.target.value;
}
function onChange(e) {
  const el = e.target;
  if (SEG && SEG.onChange(el)) return;
  if (el.dataset.clcheck && state.closure) {
    const arr = state.closure[el.dataset.clcheck];
    if (el.checked) { if (!arr.includes(el.value)) arr.push(el.value); } else arr.splice(arr.indexOf(el.value), 1);
  }
  if (el.dataset.resched && state.closure) state.closure.resched[el.dataset.resched] = el.value;
  if (el.dataset.deleg && state.closure) {
    const arr = state.closure.delegate;
    if (el.checked) arr.push(el.dataset.deleg); else arr.splice(arr.indexOf(el.dataset.deleg), 1);
    render();
  }
  if (el.matches('[data-import]') && el.files[0]) {
    const r = new FileReader();
    r.onload = () => { try { store.importAll(r.result); toast('Respaldo importado'); } catch (er) { toast('Archivo no válido: ' + er.message); } };
    r.readAsText(el.files[0]);
  }
  if (el.name === 'theme') { localStorage.setItem('mdc:theme', el.value); applyTheme(); }
}

function onClick(e) {
  const el = e.target.closest('[data-act]'); if (!el) return;
  const a = el.dataset.act; const id = el.dataset.id;
  if (el.tagName === 'A' && el.getAttribute('href') === '#') e.preventDefault();
  // Evita que el clic en el check abra la tarea
  if (a !== 'toggle-done' && e.target.closest('[data-act="toggle-done"]')) return;
  const C = state.closure;
  const actions = {
    close: closeSheet,
    new: () => openCapture(),
    'new-prefill': () => openCapture(el.dataset.text),
    'new-at': () => openCapture('', { date: state.date, start: +el.dataset.s }),
    bomberazo: openBomberazo,
    task: () => openTask(id),
    'capture-detail': () => saveCapture(el.form, true),
    'toggle-done': () => {
      const t = store.get('tasks', id); const done = t.status !== 'completada';
      store.put('tasks', { ...t, status: done ? 'completada' : (t.scheduled_start ? 'programada' : 'pendiente'), completed_at: done ? new Date().toISOString() : null });
      if (!$('#sheet').hidden && $('form[data-form="task"]')) closeSheet();
      const follow = done ? SEG.onTaskDone(t) : null; // p. ej. avanzar el documento o anotar avance del seguimiento
      if (done) toast('Completada: ' + t.title, follow || { label: 'Deshacer', fn: () => store.put('tasks', { ...store.get('tasks', id), status: t.status, completed_at: null }) });
    },
    't-start': () => { const t = store.get('tasks', id); store.put('tasks', { ...t, status: 'en_curso' }); closeSheet(); },
    't-postpone': () => { const t = store.get('tasks', id); store.put('tasks', { ...t, status: 'pospuesta', postponed_count: (t.postponed_count || 0) + 1, scheduled_date: null, scheduled_start: null }); closeSheet(); moveToDay(id, nextWorkingDay(el.dataset.to, ctx().S, store, true)); const x = store.get('tasks', id); store.put('tasks', { ...x, status: 'pospuesta' }); },
    't-move': () => { closeSheet(); moveToDay(id, el.dataset.d); },
    't-unsched': () => { const t = store.get('tasks', id); store.put('tasks', { ...t, scheduled_date: null, scheduled_start: null, status: 'pendiente' }); closeSheet(); },
    't-suggest': () => {
      const t = store.get('tasks', id); const c = ctx();
      const masked = { ...c, store: { ...store, all: (k) => store.all(k).filter((x) => k !== 'tasks' || x.id !== id), get: store.get.bind(store) } };
      const p = autoPlan([{ ...t, scheduled_date: null, scheduled_start: null }], masked, { days: 10 });
      if (!p.length || p[0].unplaced) { toast('No hay espacio en los próximos días.'); return; }
      applyPlan(p, store); closeSheet(); toast(`Programada: ${fmtShort(p[0].date)} ${p[0].start}`);
    },
    't-split': () => {
      const t = store.get('tasks', id);
      openSheet('Dividir tarea', `<div class="stack"><p>${esc(t.title)} · ${fmtDur(t.duration_min || 60)}</p><div class="btnrow">${[2, 3, 4].map((n) => `<button class="btn" data-act="do-split" data-id="${id}" data-n="${n}">En ${n} partes (${fmtDur(Math.round((t.duration_min || 60) / n / 15) * 15 || 15)})</button>`).join('')}</div></div>`);
    },
    'do-split': () => { splitTask(store.get('tasks', id), +el.dataset.n, store); closeSheet(); toast('Tarea dividida. Las nuevas partes quedaron sin programar.'); },
    't-delete': () => { if (confirmInline(el, '¿Eliminar?')) { store.remove('tasks', id); closeSheet(); toast('Tarea eliminada'); } },
    'day-prev': () => { state.date = addDays(state.date, -1); render(); refreshOutlook(); },
    'day-next': () => { state.date = addDays(state.date, 1); render(); refreshOutlook(); },
    'day-today': () => { state.date = today(); render(); },
    'week-prev': () => { state.week = addDays(state.week, -7); render(); refreshOutlook(); },
    'week-next': () => { state.week = addDays(state.week, 7); render(); refreshOutlook(); },
    'week-today': () => { state.week = weekStart(today()); render(); },
    'goto-day': () => { state.date = el.dataset.d; location.hash = '#hoy'; },
    'day-settings': openDaySettings,
    'add-event': () => openEvent(state.date),
    'event-at': () => openEvent(state.date, +el.dataset.s, Math.min(+el.dataset.e, +el.dataset.s + 60)),
    'fix-conflicts': () => { const n = reflowAfterEvent(state.date); toast(n ? `${n} tarea(s) reacomodadas` : 'Sin cambios'); },
    'ev-del': () => { store.remove('events', id); closeSheet(); },
    slot: () => openSlot(+el.dataset.s, +el.dataset.e),
    place: () => { scheduleAt(id, state.date, +el.dataset.s); closeSheet(); },
    fixed: () => openFixed(id, el.dataset.d || state.date),
    'dg-extend': () => { const d = el.dataset.d; const D = dayRec(store, d) || { id: 'day-' + d, date: d }; const S = ctx().S; const r = S.recurring.find((x) => x.id === 'dg'); store.put('days', { ...D, dg_end: toHHMM(toMin(r.end) + +el.dataset.m) }); closeSheet(); const n = reflowAfterEvent(d); if (n) toast(`${n} tarea(s) reacomodadas`); },
    'rec-off': () => { const d = el.dataset.d; const D = dayRec(store, d) || { id: 'day-' + d, date: d }; store.put('days', { ...D, rec: { ...(D.rec || {}), [el.dataset.r]: false } }); closeSheet(); },
    'lunch-short': () => { const d = el.dataset.d; const D = dayRec(store, d) || { id: 'day-' + d, date: d }; store.put('days', { ...D, lunch_min: 30 }); closeSheet(); },
    filter: () => { state.filter = el.dataset.f; if (state.view !== 'tareas') location.hash = '#tareas'; else render(); },
    suggest: () => openSuggest(state.date < today() ? today() : state.date),
    'suggest-week': () => openSuggest(state.week < today() ? today() : state.week, true),
    'replan-stale': () => {
      const stale = store.all('tasks').filter((t) => isOpen(t) && t.scheduled_date && t.scheduled_date < today());
      stale.forEach((t) => store.put('tasks', { ...t, scheduled_date: null, scheduled_start: null, status: 'pendiente' }));
      openSuggest(today());
    },
    'apply-props': () => {
      const sel = $$('[data-prop]').filter((x) => x.checked).map((x) => window._props.filter((p) => !p.unplaced)[+x.dataset.prop]);
      applyPlan(sel, store); closeSheet(); toast(`${sel.length} bloque(s) programados`);
    },
    'bz-view': () => { const b = store.get('bomberazos', id); showBzResult(b); },
    'bz-accept': () => { const b = store.get('bomberazos', id); store.put('bomberazos', { ...b, state: 'aceptado' }); closeSheet(); toast('Reorganización aceptada'); },
    'bz-revert': () => {
      const b = store.get('bomberazos', id); applyChanges(b.changes, store, 'before');
      const u = store.get('tasks', b.task_id); if (u) store.put('tasks', { ...u, scheduled_date: null, scheduled_start: null, status: 'pendiente' });
      store.put('bomberazos', { ...b, state: 'revertido' }); closeSheet(); toast('Se restauró la agenda. El imprevisto quedó registrado sin horario.');
    },
    'intr-new': () => openIntramed(null),
    intr: () => openIntramed(id),
    'intr-tab': () => { state.intrTab = el.dataset.t; render(); },
    'intr-status': () => { state.intrStatus = el.dataset.s; render(); },
    'intr-tipo': () => { state.intrTipo = el.dataset.s; render(); },
    'intr-task': () => {
      const i = store.get('intramed', id);
      const t = store.put('tasks', { title: 'Intramed: ' + (i.next_action || 'Dar seguimiento a ' + i.title), notes: i.description || '', area: i.hospital || '', category: 'Intramed', priority: i.priority || 3, clasificacion: 'seguimiento', status: 'pendiente', duration_min: 30, due_date: i.due_date || null, intramed_id: i.id, source: 'intramed' });
      closeSheet(); openTask(t.id);
    },
    'intr-del': () => { if (confirmInline(el, '¿Eliminar?')) { store.remove('intramed', id); closeSheet(); } },
    'meeting-new': () => openMeeting(null),
    meeting: () => openMeeting(id),
    'meeting-del': () => { if (confirmInline(el, '¿Eliminar?')) { store.remove('intramed_meetings', id); closeSheet(); } },
    'cl-next': () => { C.step = Math.min(6, C.step + 1); render(); window.scrollTo(0, 0); },
    'cl-prev': () => { C.step = Math.max(0, C.step - 1); render(); window.scrollTo(0, 0); },
    'cl-finish': finishClosure,
    'cl-redo': () => { state.closureResult = null; renderRedo(); },
    copy: async () => { const t = $('#' + el.dataset.target); const ok = await copyText(t.value ?? t.textContent); toast(ok ? 'Copiado' : 'No se pudo copiar'); },
    'mail-summary': () => { const body = $('#sum-text').value; location.href = `mailto:?subject=${encodeURIComponent('Resumen ' + fmtShort(today()))}&body=${encodeURIComponent(body)}`; },
    period: () => { state.period = el.dataset.p; render(); },
    'ol-connect': () => outlook.connect(),
    'ol-disconnect': () => { outlook.disconnect(); render(); },
    'ol-refresh': async () => { await refreshOutlook(true); toast(outlook.error ? 'Outlook: ' + outlook.error : 'Calendario actualizado'); },
    export: () => { const blob = new Blob([store.exportAll()], { type: 'application/json' }); const a2 = document.createElement('a'); a2.href = URL.createObjectURL(blob); a2.download = `mi-dia-de-calidad-respaldo-${today()}.json`; a2.click(); },
    'sync-now': async () => { await store.sync(); toast('Sincronizado'); },
    logout: () => store.signOut(),
    'change-pass': openPasswordSheet,
  };
  if (actions[a]) { e.preventDefault(); actions[a](); }
  else if (SEG.actions[a]) { e.preventDefault(); SEG.actions[a](el); }
}
function renderRedo() {
  // Rehacer: ignora el cierre guardado hasta generar uno nuevo
  const d = today();
  const todays = store.all('tasks').filter((t) => t.scheduled_date === d || (t.status === 'completada' && (t.completed_at || '').slice(0, 10) === d));
  state.closure = { step: 0, date: d, done: todays.filter((t) => t.status === 'completada').map((t) => t.id), pending: [], blockers: '', urgent: '', missing: '', resched: {}, delegate: [], delegate_to: {} };
  render();
}
function confirmInline(el, msg) {
  if (el.dataset.confirm) return true;
  el.dataset.confirm = '1'; el.textContent = msg + ' Toca otra vez';
  setTimeout(() => { delete el.dataset.confirm; el.textContent = 'Eliminar'; }, 3000);
  return false;
}

function finishClosure() {
  const C = state.closure; const c = ctx();
  // 1. Completadas
  C.done.forEach((id) => { const t = store.get('tasks', id); if (t && t.status !== 'completada') store.put('tasks', { ...t, status: 'completada', completed_at: new Date().toISOString() }); });
  C.pending = C.pending.filter((id) => !C.done.includes(id));
  // 2. Delegación propuesta
  C.delegate.forEach((id) => { const t = store.get('tasks', id); if (t) store.put('tasks', { ...t, clasificacion: 'delegable', delegation_status: t.delegation_status === 'autorizada' ? 'autorizada' : 'propuesta', delegate_to: C.delegate_to[id] || t.delegate_to || '' }); });
  // 3. Generar textos (antes de mover, para que el resumen muestre el destino)
  const out = closureOutputs(C, c);
  // 4. Reprogramar
  const byDate = {};
  C.pending.forEach((id) => {
    const to = C.resched[id]; const t = store.get('tasks', id); if (!t || !to || to === 'keep') return;
    store.put('tasks', { ...t, scheduled_date: null, scheduled_start: null, status: 'pospuesta', postponed_count: (t.postponed_count || 0) + 1 });
    if (to !== 'none') (byDate[to] = byDate[to] || []).push(id);
  });
  Object.entries(byDate).forEach(([d, ids]) => {
    const p = autoPlan(ids.map((x) => store.get('tasks', x)), ctx(), { from: d, days: 5 });
    applyPlan(p, store);
  });
  out.plan = closureOutputs(C, ctx()).plan; // plan con la reprogramación aplicada
  store.put('closures', { id: 'closure-' + C.date, date: C.date, answers: C, outputs: out });
  state.closureResult = out; state.closure = null;
  render(); window.scrollTo(0, 0);
}

function onSubmit(e) {
  const f = e.target.closest('form[data-form]'); if (!f || f.dataset.form === 'login') return;
  e.preventDefault();
  const fd = new FormData(f);
  const kind = f.dataset.form;
  if (SEG.forms[kind]) { SEG.forms[kind](f, fd); return; }
  if (kind === 'capture') saveCapture(f);
  if (kind === 'task') {
    const id = f.dataset.id;
    const base = id ? store.get('tasks', id) : {};
    const t = readTaskForm(f, base);
    const saved = store.put('tasks', t);
    if (!id && fd.get('to_intramed')) linkIntramed(saved, parseTask(saved.notes || saved.title).intramed_tipo);
    closeSheet(); toast('Guardado');
  }
  if (kind === 'bomberazo') runBomberazo(f);
  if (kind === 'day') {
    const d = state.date; const D = dayRec(store, d) || { id: 'day-' + d, date: d };
    const rec = {}; ctx().S.recurring.forEach((r) => { const el = f.querySelector(`[name="rec_${r.id}"]`); if (el) rec[r.id] = el.checked; });
    const tipo = fd.get('tipo');
    store.put('days', { ...D, id: 'day-' + d, date: d, tipo, start: fd.get('start') || null, end: fd.get('end') || null, note: fd.get('note').trim(), lunch_min: fd.get('lunch_min') ? +fd.get('lunch_min') : null, dg_end: fd.get('dg_end') || null, rec });
    closeSheet();
    const n = reflowAfterEvent(d);
    toast(n ? `Día actualizado · ${n} tarea(s) reacomodadas` : 'Día actualizado');
  }
  if (kind === 'event') {
    const d = f.dataset.date;
    const s = fd.get('start'), en = fd.get('end');
    if (toMin(en) <= toMin(s)) { toast('La hora de fin debe ser posterior al inicio'); return; }
    store.put('events', { date: d, title: fd.get('title').trim(), start: s, end: en, kind: fd.get('kind') });
    closeSheet();
    if (fd.get('reflow')) { const n = reflowAfterEvent(d); toast(n ? `Reunión agregada · ${n} tarea(s) reacomodadas` : 'Reunión agregada'); }
  }
  if (kind === 'intramed') {
    const id = f.dataset.id; const base = id ? store.get('intramed', id) : { log: [] };
    const i = { ...base };
    ['title', 'tipo', 'origen', 'status', 'hospital', 'folio', 'requested_at', 'due_date', 'responsable', 'next_action', 'description'].forEach((k) => { i[k] = (fd.get(k) || '').toString().trim() || (['due_date', 'requested_at'].includes(k) ? null : ''); });
    i.priority = +fd.get('priority');
    const log = [...(base.log || [])];
    if (!id) log.push({ at: new Date().toISOString(), text: 'Registrado' });
    if (id && base.status !== i.status) log.push({ at: new Date().toISOString(), text: `Estatus: ${INTR_EST[base.status]} → ${INTR_EST[i.status]}` });
    const nl = (fd.get('newlog') || '').trim(); if (nl) log.push({ at: new Date().toISOString(), text: nl });
    i.log = log;
    if (i.status === 'cerrado' && !i.closed_at) i.closed_at = new Date().toISOString();
    store.put('intramed', i);
    closeSheet(); toast('Guardado');
  }
  if (kind === 'meeting') {
    const id = f.dataset.id; const base = id ? store.get('intramed_meetings', id) : {};
    const items = fd.getAll('items');
    const m = store.put('intramed_meetings', { ...base, date: fd.get('date'), notes: fd.get('notes').trim(), acuerdos: fd.get('acuerdos').trim(), items });
    const prev = base.items || [];
    items.filter((x) => !prev.includes(x)).forEach((iid) => {
      const it = store.get('intramed', iid);
      if (it) store.put('intramed', { ...it, log: [...(it.log || []), { at: new Date().toISOString(), text: `Revisado en reunión con Intramed (${fmtShort(m.date)})${m.acuerdos ? '. Acuerdos: ' + m.acuerdos : ''}` }] });
    });
    closeSheet(); toast('Nota guardada');
  }
  if (kind === 'settings') {
    const S = ctx().S;
    const n = { ...S };
    ['weekday_start', 'weekday_end', 'saturday_start', 'saturday_end', 'saturday_ref', 'lunch_start', 'directora_label'].forEach((k) => { if (fd.get(k)) n[k] = fd.get(k).trim(); });
    ['lunch_min', 'buffer_pct', 'block_min', 'block_max', 'focus_min', 'indicadores_dias', 'followup_days'].forEach((k) => { if (fd.get(k) !== null && fd.get(k) !== '') n[k] = +fd.get(k); });
    if (dow(n.saturday_ref) !== 6) { toast('El sábado de referencia debe ser sábado'); return; }
    n.hospitals = (fd.get('hospitals') || '').split('\n').map((x) => x.trim()).filter(Boolean);
    n.recurring = S.recurring.map((r) => ({ ...r, start: fd.get(`rec.${r.id}.start`) || r.start, end: fd.get(`rec.${r.id}.end`) || r.end, mode: fd.get(`rec.${r.id}.mode`) || r.mode, days: fd.getAll(`rec.${r.id}.days`).map(Number) }));
    n.outlook_hide_dupes = f.querySelector('[name=outlook_hide_dupes]') ? !!fd.get('outlook_hide_dupes') : S.outlook_hide_dupes;
    const wantNotif = !!fd.get('notifications');
    if (wantNotif && 'Notification' in window && Notification.permission !== 'granted') Notification.requestPermission();
    n.notifications = wantNotif;
    store.put('settings', { ...n, id: 'settings' });
    toast('Ajustes guardados');
  }
  if (kind === 'newpass') {
    if (fd.get('p1') !== fd.get('p2')) { toast('Las contraseñas no coinciden'); return; }
    store.updatePassword(fd.get('p1')).then(() => { closeSheet(); toast('Contraseña actualizada'); }).catch((er) => toast(er.message));
  }
}

// ---------------- Outlook y avisos ----------------
async function refreshOutlook(force = false) {
  if (!outlook.configured || !outlook.connected) return;
  const from = weekStart(state.view === 'semana' ? state.week : state.date);
  try { await outlook.fetchRange(from, addDays(from, 13)); } catch { /* error visible en la interfaz */ }
  void force; scheduleRender();
}
function tickNotifications() {
  const S = ctx().S;
  if (!S.notifications || !('Notification' in window) || Notification.permission !== 'granted') return;
  const info = dayInfo(today(), ctx()); if (!info.working) return;
  const n = nowMin();
  const blocks = [...info.fixed, ...info.tasks.filter((t) => isOpen(t.task))];
  blocks.forEach((b) => {
    if (b.kind === 'comida') return;
    const k = `mdc:notif:${today()}:${b.id}:${b.s}`;
    if (b.s - n <= 5 && b.s - n >= 0 && !sessionStorage.getItem(k)) {
      sessionStorage.setItem(k, '1');
      try { new Notification(`${toHHMM(b.s)} · ${b.title}`, { body: b.kind === 'task' ? 'Bloque de trabajo' : 'Reunión', tag: k }); } catch { /* */ }
    }
  });
  const kc = `mdc:notif:${today()}:cierre`;
  if (info.win.e - n <= 15 && info.win.e - n >= 0 && !sessionStorage.getItem(kc)) {
    sessionStorage.setItem(kc, '1');
    try { new Notification('Cierre del día', { body: 'Toma 3 minutos para cerrar la jornada y preparar mañana.' }); } catch { /* */ }
  }
}

boot();

// Exponer para pruebas
window.__mdc = { store, state, ctx };
void CONFIG; void score; void DEFAULT_SETTINGS; void uid;
