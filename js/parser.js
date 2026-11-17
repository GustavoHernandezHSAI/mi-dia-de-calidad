// Interpretación de texto libre con reglas locales (sin IA; el texto no sale de la app).
import { addDays, dow, norm, cap, today, ymd, toHHMM } from './util.js';

const CATS = [
  ['Intramed', ['intramed', 'incidencia', 'ticket', 'falla del sistema', 'error en el sistema', 'no funciona el sistema', 'soporte tecnico']],
  ['Certificación', ['mocebpass', 'cahusa', 'accreditation canada', 'acreditacion', 'iso 7101', 'certificacion', 'auditoria', 'evidencia', 'estandar']],
  ['Normatividad', ['nom-', 'nom ', 'norma oficial', 'normatividad', 'normativa', 'dof', 'cofepris', 'regulacion', 'ley general']],
  ['Indicadores', ['indicador', 'kpi', 'run chart', 'runchart', 'tablero', 'plan de calidad', 'desviacion', 'tasa de', 'meta ']],
  ['Formatos', ['formato', 'expediente', 'hoja de', 'registro clinico', 'chsai-for']],
  ['Documentos', ['documento', 'procedimiento', 'manual', 'politica', 'lineamiento', 'instructivo', 'emision', 'emitir', 'teams']],
  ['Coordinación', ['reunion', 'junta', 'sesion', 'grupo de trabajo', 'convocar', 'minuta']],
];
const PROGRAMAS = [['MOCEBPASS', 'mocebpass'], ['CAHUSA', 'cahusa'], ['Accreditation Canada', 'accreditation canada'], ['ISO 7101:2023', 'iso 7101']];
const VERBOS = ['revisar', 'actualizar', 'validar', 'emitir', 'elaborar', 'enviar', 'dar seguimiento', 'analizar', 'coordinar', 'preparar', 'solicitar', 'corregir', 'atender', 'verificar', 'programar', 'reportar', 'capacitar', 'integrar', 'cargar', 'subir', 'llamar', 'contestar', 'responder', 'documentar', 'registrar', 'publicar', 'convocar', 'presentar', 'entregar', 'confirmar', 'investigar', 'escalar', 'cerrar', 'definir', 'redactar', 'homologar', 'calcular', 'consolidar'];
const DEFAULT_DUR = { Intramed: 30, Documentos: 60, Formatos: 60, Normatividad: 60, Certificación: 60, Indicadores: 60, Coordinación: 60, General: 30 };
const DELIVERABLE = {
  Intramed: { reporte: 'Reporte documentado y con seguimiento registrado', necesidad: 'Necesidad analizada y enviada a Intramed', soporte: 'Soporte atendido y registrado' },
  Documentos: 'Documento revisado / actualizado',
  Formatos: 'Formato revisado / validado',
  Normatividad: 'Nota de verificación normativa',
  Certificación: 'Evidencia integrada',
  Indicadores: 'Análisis del indicador',
  Coordinación: 'Minuta / acuerdos',
  General: '',
};
const CRITERIO = {
  Intramed: { reporte: 'El solicitante confirma que quedó resuelto', necesidad: 'Necesidad registrada con Intramed y estatus actualizado', soporte: 'El usuario confirma que quedó resuelto' },
  Documentos: 'Documento validado y publicado en su repositorio',
  Formatos: 'Formato validado y entregado',
  Normatividad: 'Verificación registrada y comunicada',
  Certificación: 'Evidencia cargada y revisada',
  Indicadores: 'Análisis entregado y desviaciones con acción asignada',
  Coordinación: 'Acuerdos registrados y responsables definidos',
  General: '',
};
const DIAS_N = ['domingo', 'lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sabado'];
const MESES_N = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];

const has = (t, list) => list.some((w) => t.includes(w));

function parseDate(t, ref = today()) {
  if (/\bpasado manana\b/.test(t)) return addDays(ref, 2);
  if (/\bmanana\b/.test(t)) return addDays(ref, 1);
  if (/\bhoy\b|hoy mismo/.test(t)) return ref;
  if (/esta semana/.test(t)) { const d = dow(ref); return addDays(ref, d <= 5 ? 5 - d : 6); }
  if (/(proxima|siguiente) semana/.test(t)) { const d = dow(ref); return addDays(ref, ((8 - d) % 7) || 7); }
  if (/fin de mes/.test(t)) { const r = new Date(+ref.slice(0, 4), +ref.slice(5, 7), 0); return ymd(r); }
  let m = t.match(/\b(\d{1,2})[/-](\d{1,2})(?:[/-](\d{2,4}))?\b/);
  if (m) {
    let y = m[3] ? +m[3] : +ref.slice(0, 4); if (y < 100) y += 2000;
    const d = new Date(y, +m[2] - 1, +m[1]);
    if (!m[3] && ymd(d) < ref) d.setFullYear(d.getFullYear() + 1);
    return ymd(d);
  }
  m = t.match(/\b(\d{1,2}) de (enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)\b/);
  if (m) {
    const d = new Date(+ref.slice(0, 4), MESES_N.indexOf(m[2]), +m[1]);
    if (ymd(d) < ref) d.setFullYear(d.getFullYear() + 1);
    return ymd(d);
  }
  m = t.match(/\b(?:el|este|para el|antes del|a mas tardar el)?\s*(lunes|martes|miercoles|jueves|viernes|sabado)\b/);
  if (m) {
    const target = DIAS_N.indexOf(m[1]);
    let diff = (target - dow(ref) + 7) % 7; if (diff === 0) diff = 7;
    return addDays(ref, diff);
  }
  m = t.match(/\b(?:el|para el|antes del) (\d{1,2})\b(?! ?(?:min|h|hr|hora))/);
  if (m) {
    const n = +m[1];
    if (n >= 1 && n <= 31) {
      const d = new Date(+ref.slice(0, 4), +ref.slice(5, 7) - 1, n);
      if (ymd(d) < ref) d.setMonth(d.getMonth() + 1);
      return ymd(d);
    }
  }
  return null;
}

function parseDuration(t) {
  let m = t.match(/(\d+(?:[.,]\d+)?)\s*(?:h|hr|hrs|hora|horas)\b/);
  if (m) return Math.round(parseFloat(m[1].replace(',', '.')) * 60);
  m = t.match(/(\d+)\s*(?:min|mins|minutos)\b/);
  if (m) return +m[1];
  if (/media hora/.test(t)) return 30;
  if (/una hora|1 hora/.test(t)) return 60;
  if (/hora y media/.test(t)) return 90;
  if (/rapido|rapida|5 minutos/.test(t)) return 15;
  return null;
}

function findHospital(raw, t, hospitals) {
  for (const h of hospitals || []) {
    const n = norm(h);
    if (n && t.includes(n)) return h;
  }
  const m = raw.match(/\b(?:[Hh]ospital|[Uu]nidad|[Hh]osp\.?)\s+([A-ZÁÉÍÓÚÑ0-9][\wáéíóúñÁÉÍÓÚÑ]*(?:\s+(?:de\s+)?[A-ZÁÉÍÓÚÑ][\wáéíóúñÁÉÍÓÚÑ]*)?)/);
  return m ? m[0].replace(/^hosp\.(?=\s)/i, 'Hospital').trim() : '';
}

const STRIP = [
  /\b(antes del?|para el|a m[aá]s tardar el|el|este|pr[oó]ximo)\s+(lunes|martes|mi[eé]rcoles|jueves|viernes|s[aá]bado)\b/gi,
  /\b(antes del?|para el|el)\s+\d{1,2}([/-]\d{1,2}([/-]\d{2,4})?)?\b/gi,
  /\b\d{1,2}[/-]\d{1,2}([/-]\d{2,4})?\b/g,
  /\b(\d{1,2} de (enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre))\b/gi,
  /\b(hoy mismo|hoy|pasado ma[nñ]ana|ma[nñ]ana|esta semana|(pr[oó]xima|siguiente) semana|fin de mes)\b/gi,
  /,?\s*\b\d+(?:[.,]\d+)?\s*(h|hr|hrs|hora|horas|min|mins|minutos)\b/gi,
  /\b(media hora|hora y media|una hora)\b/gi,
  /\b(urgente|es urgente)\b/gi,
];
function cleanPhrase(s) {
  let out = s;
  STRIP.forEach((re) => { out = out.replace(re, ' '); });
  out = out.replace(/\s+/g, ' ').trim().split(' ').slice(0, 10).join(' ');
  out = out.replace(/(\s+(antes|el|para|del|de|a|en|y|que|con|la|los|las))+$/i, '').replace(/[,\s]+$/, '');
  return cap(out);
}

function verbPhrase(raw) {
  const t = norm(raw);
  let best = -1, verb = null;
  for (const v of VERBOS) {
    const re = new RegExp('\\b' + v + '\\b');
    const m = re.exec(t);
    if (m && (best === -1 || m.index < best)) { best = m.index; verb = v; }
  }
  if (best === -1) return null;
  // Los índices coinciden porque norm() conserva la longitud (solo quita acentos combinados)
  let rest = raw.normalize('NFC').slice(best);
  rest = rest.split(/[.;\n]|,\s*(?=y\b|luego|despu[eé]s|para)/)[0];
  return { verb, phrase: cleanPhrase(rest) };
}

export function parseTask(raw, { hospitals = [], ref = today() } = {}) {
  const text = (raw || '').normalize('NFC').trim();
  const t = norm(text);
  const r = { questions: [] };
  // Categoría
  r.category = 'General';
  for (const [c, kws] of CATS) if (has(t, kws)) { r.category = c; break; }
  const programa = PROGRAMAS.find(([, k]) => t.includes(k));
  if (programa) r.programa = programa[0];
  // Intramed: tipo
  r.intramed = r.category === 'Intramed' || /intramed/.test(t);
  if (r.intramed) r.category = 'Intramed';
  // reporte = algo falla o un hospital lo reporta; necesidad = algo nuevo que se requiere (de un hospital o de Calidad)
  r.intramed_tipo = /requerimiento|necesidad|solicitan|solicitud|cambio|agregar|nuevo campo|nueva funcion|mejora|desarrollo|configurar|crear (un|una|el|la)/.test(t) && !/problema|falla|error|no funciona|no permite|incidencia|caido|lento|reporto|reportan/.test(t) ? 'necesidad' : (/soporte/.test(t) ? 'soporte' : 'reporte');
  // Riesgo y prioridad
  r.riesgo = /paciente|seguridad|evento adverso|centinela|cuasifalla|riesgo|cofepris|nom-|normativ|auditori|certificac|mocebpass|cahusa|accreditation|iso 7101/.test(t);
  r.due_date = parseDate(t, ref);
  const urgentWord = /urgente|inmediat|hoy mismo|ya mismo|cuanto antes|asap|critico/.test(t);
  const dueDays = r.due_date ? Math.round((new Date(r.due_date) - new Date(ref)) / 86400000) : null;
  if (urgentWord || /evento adverso|centinela/.test(t) || dueDays === 0) r.priority = 1;
  else if (r.riesgo || (dueDays !== null && dueDays <= 2) || /importante|prioridad/.test(t) || (r.intramed && r.intramed_tipo === 'reporte')) r.priority = 2;
  else if (/cuando se pueda|sin prisa|fyi|para conocimiento|informativo|referencia/.test(t)) r.priority = 4;
  else r.priority = 3;
  r.impacto = r.riesgo || ['Certificación', 'Indicadores', 'Intramed'].includes(r.category) ? 'alto' : 'medio';
  // Clasificación
  if (/fyi|para conocimiento|solo referencia|informativo/.test(t)) r.clasificacion = 'referencia';
  else if (/esperar|en espera|pendiente de (respuesta|informacion|autorizacion|visto bueno|validacion)|falta informacion|requiere (autorizacion|visto bueno|decision|validacion de)|que me envien|que nos envien/.test(t)) r.clasificacion = 'espera';
  else if (/delegar|que lo haga|asignar a|gerente de calidad podria/.test(t)) r.clasificacion = 'delegable';
  else if (/^(dar )?seguimiento|monitorear|vigilar|seguimiento semanal/.test(t) && !urgentWord) r.clasificacion = 'seguimiento';
  else if (urgentWord || dueDays === 0) r.clasificacion = 'inmediata';
  else r.clasificacion = 'programable';
  // Duración
  const d = parseDuration(t);
  r.duration_min = d || DEFAULT_DUR[r.category];
  r.duration_estimated = !d;
  // Hospital / área
  r.area = findHospital(text, t, hospitals);
  // Título, siguiente acción, entregable, criterio
  const vp = verbPhrase(text);
  const head = text.split(/[,.;\n]/)[0].trim();
  let title = vp ? vp.phrase : cleanPhrase(head);
  if (r.area && !norm(title).includes(norm(r.area))) title += ' · ' + r.area;
  if (r.intramed && !/intramed/i.test(title)) title = 'Intramed: ' + title;
  r.title = title.slice(0, 90);
  r.next_action = vp ? vp.phrase : '';
  const del = DELIVERABLE[r.category];
  r.deliverable = typeof del === 'object' ? del[r.intramed_tipo] || del.reporte : del;
  const cri = CRITERIO[r.category];
  r.done_criteria = typeof cri === 'object' ? cri[r.intramed_tipo] || cri.reporte : cri;
  if (r.programa && r.category === 'Certificación') r.deliverable = `Evidencia ${r.programa} integrada`;
  r.notes = text;
  r.documental = ['Documentos', 'Formatos'].includes(r.category) && /nuevo|actualiz|emitir|emision|revisar|elaborar|diagrama/.test(t);
  // Preguntas en lugar de inventar
  if (r.intramed && r.intramed_tipo === 'reporte' && !r.area) r.questions.push({ field: 'area', q: '¿Qué hospital o unidad lo reporta? (vacío si es interno)' });
  if (r.priority <= 2 && !r.due_date) r.questions.push({ field: 'due_date', q: '¿Tiene fecha límite?' });
  if (r.clasificacion === 'espera') r.questions.push({ field: 'waiting_on', q: '¿De quién depende (información o decisión)?' });
  if (!vp) r.questions.push({ field: 'next_action', q: '¿Cuál es la siguiente acción concreta?' });
  return r;
}

export const _test = { parseDate, parseDuration, verbPhrase, toHHMM };
