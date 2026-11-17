// Lector mínimo de Excel (.xlsx) y CSV, sin librerías externas: funciona sin internet.
// Un .xlsx es un ZIP con XML; se descomprime con DecompressionStream del navegador.

async function unzip(buf) {
  const dv = new DataView(buf);
  let eocd = -1;
  for (let i = buf.byteLength - 22; i >= Math.max(0, buf.byteLength - 66000); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  if (eocd < 0) throw new Error('El archivo no parece un Excel .xlsx válido.');
  const count = dv.getUint16(eocd + 10, true);
  let p = dv.getUint32(eocd + 16, true);
  const files = {};
  const dec = new TextDecoder();
  for (let n = 0; n < count; n++) {
    if (dv.getUint32(p, true) !== 0x02014b50) break;
    const method = dv.getUint16(p + 10, true);
    const csize = dv.getUint32(p + 20, true);
    const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
    const off = dv.getUint32(p + 42, true);
    const name = dec.decode(new Uint8Array(buf, p + 46, nlen));
    files[name] = { method, csize, off };
    p += 46 + nlen + xlen + clen;
  }
  const read = async (name) => {
    const f = files[name]; if (!f) return null;
    const lnl = dv.getUint16(f.off + 26, true), lxl = dv.getUint16(f.off + 28, true);
    const data = new Uint8Array(buf, f.off + 30 + lnl + lxl, f.csize);
    if (f.method === 0) return dec.decode(data);
    if (f.method !== 8) throw new Error('Compresión no soportada en el archivo.');
    const ds = new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return await new Response(ds).text();
  };
  return { names: Object.keys(files), read };
}

const xml = (s) => new DOMParser().parseFromString(s, 'application/xml');
const colIndex = (ref) => { const m = ref.match(/^[A-Z]+/)[0]; let n = 0; for (const ch of m) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1; };

// Devuelve [{name, rows: [[celda,...],...]}]
export async function readXlsx(file) {
  const z = await unzip(await file.arrayBuffer());
  const shared = [];
  const ss = await z.read('xl/sharedStrings.xml');
  if (ss) xml(ss).querySelectorAll('si').forEach((si) => shared.push([...si.querySelectorAll('t')].map((t) => t.textContent).join('')));
  const wb = xml(await z.read('xl/workbook.xml'));
  const rels = xml(await z.read('xl/_rels/workbook.xml.rels'));
  const relMap = {};
  rels.querySelectorAll('Relationship').forEach((r) => { relMap[r.getAttribute('Id')] = r.getAttribute('Target'); });
  const sheets = [];
  for (const sh of wb.querySelectorAll('sheet')) {
    const rid = sh.getAttribute('r:id') || sh.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
    let target = relMap[rid] || '';
    target = target.replace(/^\/?xl\//, '').replace(/^\//, '');
    const s = await z.read('xl/' + target);
    if (!s) continue;
    const rows = [];
    xml(s).querySelectorAll('sheetData row').forEach((r) => {
      const idx = (+r.getAttribute('r') || rows.length + 1) - 1;
      const row = [];
      r.querySelectorAll('c').forEach((c) => {
        const t = c.getAttribute('t'); const v = c.querySelector('v'); const is = c.querySelector('is');
        let val = v ? v.textContent : (is ? [...is.querySelectorAll('t')].map((x) => x.textContent).join('') : '');
        if (t === 's') val = shared[+val] ?? '';
        else if (t === 'b') val = val === '1' ? 'Sí' : 'No';
        else if (!t || t === 'n') { if (val !== '' && !isNaN(+val)) val = +val; }
        row[colIndex(c.getAttribute('r') || 'A1')] = val;
      });
      rows[idx] = row;
    });
    sheets.push({ name: sh.getAttribute('name'), rows: Array.from(rows, (r) => Array.from(r || [], (x) => (x === undefined ? '' : x))) });
  }
  return sheets;
}

export async function readCsv(file) {
  let text = await file.text();
  text = text.replace(/^﻿/, '');
  const sep = (text.split('\n')[0].match(/;/g) || []).length > (text.split('\n')[0].match(/,/g) || []).length ? ';' : ',';
  const rows = []; let row = [], cell = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) { if (ch === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; } else cell += ch; continue; }
    if (ch === '"') q = true;
    else if (ch === sep) { row.push(cell); cell = ''; }
    else if (ch === '\n') { row.push(cell.replace(/\r$/, '')); rows.push(row); row = []; cell = ''; }
    else cell += ch;
  }
  if (cell || row.length) { row.push(cell); rows.push(row); }
  return [{ name: file.name, rows }];
}

export function toCsv(rows) {
  const esc = (v) => { const s = String(v ?? ''); return /[",;\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s; };
  return '﻿' + rows.map((r) => r.map(esc).join(',')).join('\r\n');
}

// Fecha desde Excel (número de serie, dd/mm/aaaa o aaaa-mm-dd) → 'AAAA-MM-DD' o null
export function toDate(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (typeof v === 'number' && v > 20000 && v < 80000) {
    const d = new Date(Date.UTC(1899, 11, 30) + v * 86400000);
    return d.toISOString().slice(0, 10);
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  m = s.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})/);
  if (m) { const y = m[3].length === 2 ? '20' + m[3] : m[3]; return `${y}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`; }
  return null;
}
