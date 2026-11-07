/**
 * Misión Chocó · Propuestas de equipo
 * Guarda cada ficha en la hoja "propuestas" y devuelve los datos a la web.
 * Implementar → Nueva implementación → App web · Ejecutar como: Yo · Acceso: Cualquier usuario
 */
const HOJA = 'propuestas';
const COLS = ['id','equipo','lider','dia','espacio','funcion','turno','estado','version','actualizado','enviado','tardia',
  'costo','cobertura_pct','idea','enfoque','tematica','minuto_a_minuto','quien_hace_que','materiales','necesita','ts','p_json','d_json'];

function hoja_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(HOJA);
  if (!s) {
    s = ss.insertSheet(HOJA);
    s.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet(e) {
  const s = hoja_(), v = s.getDataRange().getValues(); v.shift();
  const t = e && e.parameter && e.parameter.t;
  const C = n => COLS.indexOf(n);
  const rows = v.filter(r => r[0] && (!t || String(r[C('equipo')]) === String(t))).map(r => {
    let p = null, d = null;
    try { p = JSON.parse(r[C('p_json')] || 'null'); } catch (err) {}
    try { d = JSON.parse(r[C('d_json')] || 'null'); } catch (err) {}
    return { id: r[0], estado: r[C('estado')], rev: Number(r[C('version')]) || 0, ts: Number(r[C('ts')]) || 0, p: p, d: d };
  });
  return json_({ ok: true, now: Date.now(), rows: rows });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const b = JSON.parse(e.postData.contents), list = b.items || [b], s = hoja_();
    const ids = s.getRange(1, 1, Math.max(1, s.getLastRow()), 1).getValues().map(r => String(r[0]));
    list.forEach(x => {
      if (!/^t[1-5]-\w+$/.test(String(x.id || ''))) return;
      const p = x.p || {};
      const txt = (arr, f) => (arr || []).map(f).join('\n');
      const row = [
        x.id, x.t, x.ld, x.dia, x.sn, x.fn, x.turno || '', x.estado, x.rev || 0,
        new Date(x.ts), x.sentTs ? new Date(x.sentTs) : '', x.late ? 'sí' : '',
        p.tot || 0, p.cov || 0, p.idea || '', p.obj || '', p.tema || '',
        txt(p.mm, r => r.a + '–' + r.b + ' ' + r.q + (r.p ? ' (' + r.p + ')' : '')),
        txt(p.pe, r => r.p + ': ' + r.r),
        txt(p.ma, r => r.c + ' × ' + r.i + ' (' + r.t + ') · ' + ({ Q: 'Quibdó', B: 'Bogotá', T: 'Ya lo tenemos' }[r.s] || '') + (r.s === 'T' ? '' : ' · $' + r.u) + (r.re ? ' · reutilizable' : '')),
        p.nec || '', x.ts, JSON.stringify(p), JSON.stringify(x.d || null)
      ];
      let i = ids.indexOf(String(x.id));
      if (i > 0) {
        const old = Number(s.getRange(i + 1, COLS.indexOf('ts') + 1).getValue()) || 0;
        if (old > x.ts) return;
        s.getRange(i + 1, 1, 1, row.length).setValues([row]);
      } else {
        s.appendRow(row); ids.push(String(x.id));
      }
    });
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
