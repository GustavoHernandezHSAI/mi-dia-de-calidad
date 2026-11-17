// Integración con Outlook (Microsoft 365) — SOLO LECTURA del calendario.
// OAuth 2.0 con PKCE desde el navegador (app registrada como "SPA" en Microsoft Entra ID).
// Permisos solicitados: Calendars.Read (+ openid / offline_access para mantener la sesión).
// Los eventos se leen en el navegador y NO se guardan en la base de datos de la nube.
import { CONFIG } from './config.js';
import { toMin, ymd } from './util.js';

const SCOPES = 'openid offline_access Calendars.Read';
const TOK = 'mdc:ol:token';
const CACHE = 'mdc:ol:cache';
const redirect = () => location.origin + location.pathname;
const base = () => `https://login.microsoftonline.com/${CONFIG.MS_TENANT_ID || 'organizations'}/oauth2/v2.0`;

export const outlook = {
  get configured() { return !!CONFIG.MS_CLIENT_ID; },
  get connected() { return !!localStorage.getItem(TOK); },
  events: {},   // { 'YYYY-MM-DD': [{id,s,e,title,showAs,allDay}] }
  lastSync: null,
  error: null,

  async connect() {
    const verifier = b64url(crypto.getRandomValues(new Uint8Array(32)));
    const challenge = b64url(new Uint8Array(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier))));
    const state = b64url(crypto.getRandomValues(new Uint8Array(16)));
    sessionStorage.setItem('mdc:ol:pkce', JSON.stringify({ verifier, state }));
    const p = new URLSearchParams({
      client_id: CONFIG.MS_CLIENT_ID, response_type: 'code', redirect_uri: redirect(),
      scope: SCOPES, code_challenge: challenge, code_challenge_method: 'S256', state, prompt: 'select_account',
    });
    location.href = `${base()}/authorize?${p}`;
  },

  // Llamar al arrancar: completa el inicio de sesión si venimos de Microsoft
  async handleRedirect() {
    const q = new URLSearchParams(location.search);
    if (!q.get('code') && !q.get('error')) return false;
    const saved = JSON.parse(sessionStorage.getItem('mdc:ol:pkce') || 'null');
    history.replaceState({}, '', redirect());
    if (q.get('error')) { this.error = q.get('error_description') || q.get('error'); return true; }
    if (!saved || saved.state !== q.get('state')) { this.error = 'Respuesta de Microsoft no válida (state).'; return true; }
    try {
      await this._token({ grant_type: 'authorization_code', code: q.get('code'), code_verifier: saved.verifier, redirect_uri: redirect() });
    } catch (e) { this.error = e.message; }
    sessionStorage.removeItem('mdc:ol:pkce');
    return true;
  },

  async _token(params) {
    const body = new URLSearchParams({ client_id: CONFIG.MS_CLIENT_ID, scope: SCOPES, ...params });
    const r = await fetch(`${base()}/token`, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body });
    const j = await r.json();
    if (!r.ok) throw new Error(j.error_description || j.error || 'Error al obtener token');
    localStorage.setItem(TOK, JSON.stringify({ access: j.access_token, refresh: j.refresh_token, exp: Date.now() + (j.expires_in - 120) * 1000 }));
  },

  async _access() {
    const t = JSON.parse(localStorage.getItem(TOK) || 'null');
    if (!t) throw new Error('No conectado');
    if (Date.now() < t.exp) return t.access;
    if (!t.refresh) throw new Error('Sesión de Outlook expirada');
    await this._token({ grant_type: 'refresh_token', refresh_token: t.refresh });
    return JSON.parse(localStorage.getItem(TOK)).access;
  },

  disconnect() { localStorage.removeItem(TOK); localStorage.removeItem(CACHE); this.events = {}; },

  loadCache() {
    try { const c = JSON.parse(localStorage.getItem(CACHE) || 'null'); if (c) { this.events = c.events || {}; this.lastSync = c.at; } } catch { /* */ }
  },

  // Lee eventos entre dos fechas (YYYY-MM-DD, inclusive)
  async fetchRange(from, to) {
    if (!this.configured || !this.connected) return;
    try {
      const token = await this._access();
      const start = `${from}T00:00:00`, endD = new Date(to + 'T00:00:00'); endD.setDate(endD.getDate() + 1);
      let url = `https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=${start}&endDateTime=${ymd(endD)}T00:00:00&$select=id,subject,start,end,showAs,isAllDay,isCancelled&$orderby=start/dateTime&$top=250`;
      const fresh = {};
      while (url) {
        const r = await fetch(url, { headers: { Authorization: 'Bearer ' + token, Prefer: `outlook.timezone="${CONFIG.TIMEZONE}"` } });
        if (r.status === 401) { localStorage.removeItem(TOK); throw new Error('La sesión de Outlook expiró. Vuelve a conectar.'); }
        const j = await r.json();
        if (!r.ok) throw new Error(j.error?.message || 'Error al leer el calendario');
        for (const ev of j.value) {
          if (ev.isCancelled) continue;
          const sd = ev.start.dateTime.slice(0, 10), ed = ev.end.dateTime.slice(0, 10);
          // Eventos de varios días: se reparten por día
          for (let d = new Date(sd + 'T00:00:00'); ymd(d) <= ed; d.setDate(d.getDate() + 1)) {
            const day = ymd(d);
            if (day === ed && ev.end.dateTime.slice(11, 16) === '00:00' && day !== sd) break;
            const s = day === sd && !ev.isAllDay ? toMin(ev.start.dateTime.slice(11, 16)) : 0;
            const e = day === ed && !ev.isAllDay ? toMin(ev.end.dateTime.slice(11, 16)) : 24 * 60;
            (fresh[day] = fresh[day] || []).push({ id: ev.id + '-' + day, s, e: Math.max(e, s + 15), title: ev.subject || '(Ocupado)', showAs: ev.showAs, allDay: ev.isAllDay });
          }
        }
        url = j['@odata.nextLink'] || null;
      }
      for (let d = new Date(from + 'T00:00:00'); ymd(d) <= to; d.setDate(d.getDate() + 1)) this.events[ymd(d)] = fresh[ymd(d)] || [];
      this.lastSync = new Date().toISOString();
      this.error = null;
      localStorage.setItem(CACHE, JSON.stringify({ at: this.lastSync, events: this.events }));
    } catch (e) { this.error = e.message; throw e; }
  },
};

function b64url(bytes) {
  let s = ''; bytes.forEach((b) => { s += String.fromCharCode(b); });
  return btoa(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}
