// Capa de datos: caché local + cola sin conexión + sincronización con Supabase en tiempo real.
// Cada registro remoto tiene: user_id, id, updated_at, deleted, data (jsonb).
import { CONFIG } from './config.js';
import { uid } from './util.js';

export const COLLS = ['tasks', 'intramed', 'intramed_meetings', 'events', 'days', 'closures', 'bomberazos', 'settings', 'projects', 'docs'];
const SUPABASE_CDN = 'https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2/+esm';

let sb = null;
let user = null;
let data = {};
let outbox = [];
let lastPull = {};
let listeners = new Set();
let status = 'local'; // local | online | offline | syncing | error
let flushing = false;
let channel = null;

const remote = () => !!(CONFIG.SUPABASE_URL && CONFIG.SUPABASE_ANON_KEY);
const key = (k) => `mdc:${user ? user.id : 'local'}:${k}`;
const ts = (s) => (s ? Date.parse(s) : 0);

function loadCache() {
  try { data = JSON.parse(localStorage.getItem(key('data'))) || {}; } catch { data = {}; }
  try { outbox = JSON.parse(localStorage.getItem(key('outbox'))) || []; } catch { outbox = []; }
  try { lastPull = JSON.parse(localStorage.getItem(key('pull'))) || {}; } catch { lastPull = {}; }
  COLLS.forEach((c) => { data[c] = data[c] || {}; });
}
let saveTimer = null;
function saveCache() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => {
    try {
      localStorage.setItem(key('data'), JSON.stringify(data));
      localStorage.setItem(key('outbox'), JSON.stringify(outbox));
      localStorage.setItem(key('pull'), JSON.stringify(lastPull));
    } catch (e) { console.warn('No se pudo guardar caché local', e); }
  }, 50);
}
function emit(info) { listeners.forEach((fn) => { try { fn(info); } catch (e) { console.error(e); } }); }
function setStatus(s) { if (status !== s) { status = s; emit({ status: s }); } }

const toRow = (coll, o) => {
  const { id, updated_at, deleted, ...rest } = o;
  return { user_id: user.id, id, updated_at, deleted: !!deleted, data: rest };
};
const fromRow = (r) => ({ ...(r.data || {}), id: r.id, updated_at: r.updated_at, deleted: r.deleted });

function merge(coll, obj) {
  const cur = data[coll][obj.id];
  if (!cur || ts(obj.updated_at) >= ts(cur.updated_at)) { data[coll][obj.id] = obj; return true; }
  return false;
}

async function pull() {
  if (!sb || !navigator.onLine) return;
  setStatus('syncing');
  for (const coll of COLLS) {
    let q = sb.from(coll).select('*');
    if (lastPull[coll]) q = q.gt('updated_at', lastPull[coll]);
    const { data: rows, error } = await q.order('updated_at', { ascending: true }).limit(5000);
    if (error) { console.error('pull', coll, error); setStatus('error'); return; }
    for (const r of rows) {
      merge(coll, fromRow(r));
      if (!lastPull[coll] || ts(r.updated_at) > ts(lastPull[coll])) lastPull[coll] = r.updated_at;
    }
  }
  saveCache();
  setStatus('online');
  emit({ remote: true });
}

async function flush() {
  if (!sb || flushing || !navigator.onLine || !outbox.length) { if (sb && navigator.onLine && !outbox.length && status !== 'error') setStatus('online'); return; }
  flushing = true;
  setStatus('syncing');
  try {
    while (outbox.length) {
      const batch = outbox.slice(0, 100);
      const byColl = {};
      batch.forEach((it) => { (byColl[it.coll] = byColl[it.coll] || []).push(toRow(it.coll, it.obj)); });
      for (const coll of Object.keys(byColl)) {
        const { error } = await sb.from(coll).upsert(byColl[coll], { onConflict: 'user_id,id' });
        if (error) throw error;
      }
      // Quitar solo lo enviado: si un registro cambió mientras se enviaba, su versión nueva queda en cola
      const sent = new Set(batch);
      outbox = outbox.filter((it) => !sent.has(it));
      saveCache();
    }
    setStatus('online');
  } catch (e) {
    console.error('flush', e);
    setStatus(navigator.onLine ? 'error' : 'offline');
  } finally { flushing = false; }
}

function subscribe() {
  if (!sb || channel) return;
  channel = sb.channel('mdc-' + user.id);
  COLLS.forEach((coll) => {
    channel.on('postgres_changes', { event: '*', schema: 'public', table: coll, filter: `user_id=eq.${user.id}` }, (p) => {
      if (!p.new || !p.new.id) return;
      if (merge(coll, fromRow(p.new))) { saveCache(); emit({ remote: true, coll }); }
    });
  });
  channel.subscribe((s) => { if (s === 'SUBSCRIBED') pull(); });
}

export const store = {
  get mode() { return remote() ? 'remote' : 'local'; },
  get status() { return status; },
  get user() { return user; },
  get pending() { return outbox.length; },

  async init() {
    if (!remote()) { user = null; loadCache(); setStatus('local'); return { ready: true }; }
    try {
      const mod = await import(SUPABASE_CDN);
      sb = mod.createClient(CONFIG.SUPABASE_URL, CONFIG.SUPABASE_ANON_KEY, { auth: { persistSession: true, autoRefreshToken: true } });
    } catch (e) {
      console.error(e);
      // Sin conexión en el primer arranque: si hay sesión previa, trabaja con la caché
      const last = localStorage.getItem('mdc:lastUser');
      if (last) { user = JSON.parse(last); loadCache(); setStatus('offline'); return { ready: true }; }
      return { ready: false, error: 'No se pudo cargar el servicio de sincronización. Revisa tu conexión.' };
    }
    const { data: s } = await sb.auth.getSession();
    sb.auth.onAuthStateChange((ev, session) => {
      if (ev === 'SIGNED_OUT') emit({ signedOut: true });
      if (ev === 'PASSWORD_RECOVERY') { this._recovery = true; emit({ recovery: true }); }
    });
    if (!s.session) return { ready: false, needLogin: true };
    await this._start(s.session.user);
    return { ready: true };
  },

  async _start(u) {
    user = { id: u.id, email: u.email };
    localStorage.setItem('mdc:lastUser', JSON.stringify(user));
    loadCache();
    setStatus(navigator.onLine ? 'syncing' : 'offline');
    window.addEventListener('online', () => { pull().then(flush); });
    window.addEventListener('offline', () => setStatus('offline'));
    document.addEventListener('visibilitychange', () => { if (!document.hidden) pull().then(flush); });
    // Arranque sin bloquear la interfaz
    pull().then(flush).then(subscribe).catch((e) => console.error(e));
  },

  async signIn(email, password) {
    const { data: d, error } = await sb.auth.signInWithPassword({ email, password });
    if (error) throw error;
    await this._start(d.user);
  },
  async resetPassword(email) {
    const { error } = await sb.auth.resetPasswordForEmail(email, { redirectTo: location.origin + location.pathname });
    if (error) throw error;
  },
  async updatePassword(password) {
    const { error } = await sb.auth.updateUser({ password });
    if (error) throw error;
  },
  async signOut() {
    if (sb) await sb.auth.signOut();
    localStorage.removeItem('mdc:lastUser');
    location.reload();
  },

  on(fn) { listeners.add(fn); return () => listeners.delete(fn); },
  all(coll) { return Object.values(data[coll] || {}).filter((o) => !o.deleted); },
  get(coll, id) { const o = (data[coll] || {})[id]; return o && !o.deleted ? o : null; },
  raw(coll, id) { return (data[coll] || {})[id] || null; }, // incluye eliminados

  put(coll, obj, { silent = false } = {}) {
    const o = { ...obj, id: obj.id || uid(), updated_at: new Date().toISOString() };
    if (!o.created_at) o.created_at = o.updated_at;
    data[coll][o.id] = o;
    if (remote() && user) {
      outbox = outbox.filter((it) => !(it.coll === coll && it.obj.id === o.id));
      outbox.push({ coll, obj: o });
    }
    saveCache();
    if (!silent) emit({ coll, id: o.id });
    flush();
    return o;
  },
  putMany(coll, list) { const out = list.map((o) => this.put(coll, o, { silent: true })); emit({ coll }); return out; },
  remove(coll, id) { const o = (data[coll] || {})[id]; if (o) this.put(coll, { ...o, deleted: true }); },

  exportAll() { return JSON.stringify({ exported_at: new Date().toISOString(), data }, null, 2); },
  importAll(json) {
    const parsed = JSON.parse(json);
    const src = parsed.data || parsed;
    COLLS.forEach((c) => Object.values(src[c] || {}).forEach((o) => this.put(c, o, { silent: true })));
    emit({});
  },
  sync() { return pull().then(flush); },
};
