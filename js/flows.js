// Flujos de trabajo de documentos y formatos (etapas, quién tiene la "cancha", rondas de revisión).
// Las etapas se pueden ajustar aquí si el proceso real cambia.

export const CANCHA = {
  calidad: 'Calidad (yo)',
  dueno: 'Dueño de proceso / responsable',
  consultor: 'Consultor de expediente clínico',
  proveedor: 'Proveedor de ECE',
  compras: 'Compras',
  autoriza: 'Autorización',
  nadie: '—',
};

// back: etapa a la que regresa cuando hay observaciones; round: etapa que cuenta rondas de revisión
// task: verbo de la tarea que se crea cuando la etapa queda en mi cancha; min: duración sugerida
export const FLOWS = {
  documento_nuevo: {
    label: 'Documento nuevo',
    stages: [
      { id: 'solicitud', label: 'Solicitud / reunión inicial', cancha: 'dueno' },
      { id: 'elaboracion', label: 'Elaboración por dueño de proceso', cancha: 'dueno' },
      { id: 'revision', label: 'Revisión por Calidad', cancha: 'calidad', round: true, back: 'correcciones', task: 'Revisar', min: 60 },
      { id: 'correcciones', label: 'Correcciones del dueño de proceso', cancha: 'dueno', next: 'revision' },
      { id: 'vobo', label: 'Visto bueno de Calidad', cancha: 'calidad', task: 'Dar visto bueno', min: 15 },
      { id: 'diagrama', label: 'Diagrama de flujo (Calidad)', cancha: 'calidad', task: 'Elaborar diagrama de flujo', min: 90 },
      { id: 'autorizacion', label: 'Autorización final', cancha: 'autoriza' },
      { id: 'intramed', label: 'Carga en Intramed (módulo documental)', cancha: 'calidad', task: 'Cargar en Intramed', min: 15 },
      { id: 'terminado', label: 'Terminado', cancha: 'nadie', done: true },
    ],
  },
  documento_actualizacion: {
    label: 'Actualización de documento',
    stages: [
      { id: 'envio', label: 'Enviado a responsable de actualizar', cancha: 'dueno' },
      { id: 'revision', label: 'Revisión por Calidad', cancha: 'calidad', round: true, back: 'correcciones', task: 'Revisar actualización', min: 60 },
      { id: 'correcciones', label: 'Correcciones del responsable', cancha: 'dueno', next: 'revision' },
      { id: 'vobo', label: 'Visto bueno de Calidad', cancha: 'calidad', task: 'Dar visto bueno', min: 15 },
      { id: 'diagrama', label: 'Diagrama de flujo (si aplica)', cancha: 'calidad', task: 'Actualizar diagrama de flujo', min: 60 },
      { id: 'autorizacion', label: 'Autorización final', cancha: 'autoriza' },
      { id: 'intramed', label: 'Carga en Intramed (módulo documental)', cancha: 'calidad', task: 'Cargar en Intramed', min: 15 },
      { id: 'terminado', label: 'Terminado', cancha: 'nadie', done: true },
    ],
  },
  formato_ece: {
    label: 'Formato de expediente clínico electrónico',
    stages: [
      { id: 'elaboracion', label: 'Elaboración del formato (Calidad)', cancha: 'calidad', task: 'Elaborar formato', min: 60 },
      { id: 'consultor', label: 'Revisión del consultor contra norma', cancha: 'consultor', round: true, back: 'correccion' },
      { id: 'correccion', label: 'Corrección de observaciones (Calidad)', cancha: 'calidad', next: 'consultor', task: 'Corregir observaciones del consultor', min: 60 },
      { id: 'proveedor', label: 'Enviado al proveedor de ECE', cancha: 'proveedor' },
      { id: 'compras', label: 'Enviado a Compras (formatos físicos)', cancha: 'compras' },
      { id: 'intramed', label: 'Carga en Intramed (módulo documental)', cancha: 'calidad', task: 'Cargar en Intramed', min: 15 },
      { id: 'terminado', label: 'Terminado', cancha: 'nadie', done: true },
    ],
  },
};

export const CLASES = ['Procedimiento', 'Manual', 'Política', 'Instructivo', 'Lineamiento', 'Guía', 'Formato', 'Otro'];

export const flowOf = (doc) => FLOWS[doc.flow] || FLOWS.documento_nuevo;
export const stageIndex = (doc) => Math.max(0, flowOf(doc).stages.findIndex((s) => s.id === doc.stage));
export const stageOf = (doc) => flowOf(doc).stages[stageIndex(doc)];
export const isDocDone = (doc) => !!stageOf(doc).done;
export const canchaOf = (doc) => doc.cancha_override || stageOf(doc).cancha;
export function progress(doc) {
  const st = flowOf(doc).stages;
  return Math.round((stageIndex(doc) / (st.length - 1)) * 100);
}
// Siguiente etapa "normal" (las de corrección regresan a revisión)
export function nextStage(doc) {
  const st = flowOf(doc).stages; const cur = stageOf(doc);
  if (cur.next) return st.find((s) => s.id === cur.next);
  const i = stageIndex(doc);
  // saltar etapas de corrección al avanzar desde revisión
  let j = i + 1;
  while (st[j] && st[j].next) j++;
  return st[j] || null;
}
export const backStage = (doc) => { const cur = stageOf(doc); return cur.back ? flowOf(doc).stages.find((s) => s.id === cur.back) : null; };

// Convierte el texto de estatus de un Excel a una etapa (o null si no se reconoce)
export function guessStage(flowId, text) {
  const t = String(text || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
  if (!t) return null;
  const has = (re) => re.test(t);
  const pick = (ids) => ids.find((id) => FLOWS[flowId].stages.some((s) => s.id === id));
  if (has(/termin|conclu|liberad|vigente|publicad|cerrad|emitid/)) return 'terminado';
  if (has(/intramed|cargar|carga|subir/)) return pick(['intramed']);
  if (has(/autoriz|firma/)) return pick(['autorizacion']);
  if (has(/diagrama|flujograma/)) return pick(['diagrama']);
  if (has(/visto bueno|vobo|validad/)) return pick(['vobo', 'proveedor']);
  if (has(/compras/)) return pick(['compras']);
  if (has(/proveedor/)) return pick(['proveedor']);
  if (has(/consultor/)) return pick(['consultor']);
  if (has(/correc|observ|ajust/)) return pick(['correcciones', 'correccion']);
  if (has(/revis/)) return pick(['revision', 'consultor']);
  if (has(/elabor|borrador|en proceso|proceso/)) return pick(['elaboracion', 'envio']);
  if (has(/solicit|reunion|pendiente|inicio/)) return pick(['solicitud', 'envio', 'elaboracion']);
  return null;
}
