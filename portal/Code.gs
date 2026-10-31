/**
 * FC Portal – kho dữ liệu chung (Google Sheet) cho các app FC trên ruralwcm.github.io
 * Dán vào: Google Sheet "FC Portal - Dữ liệu chung" → Tiện ích mở rộng → Apps Script
 * Deploy: Web app · Execute as: Me · Who has access: Anyone
 */
const SHEET = 'kv';      // key | part | data | updatedAt
const PART = 45000;      // giới hạn 50.000 ký tự/ô → chia nhỏ

function sh_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) { sh = ss.insertSheet(SHEET); sh.getRange(1, 1, 1, 4).setValues([['key', 'part', 'data', 'updatedAt']]); sh.setFrozenRows(1); }
  sh.getRange('C:C').setNumberFormat('@');
  return sh;
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function readAll_(prefix) {
  const v = sh_().getDataRange().getValues(), m = {}, t = {};
  for (let i = 1; i < v.length; i++) {
    const k = String(v[i][0] || ''); if (!k || (prefix && k.indexOf(prefix) !== 0)) continue;
    (m[k] = m[k] || [])[Number(v[i][1]) || 0] = String(v[i][2]);
    t[k] = v[i][3];
  }
  const o = {}; for (const k in m) { try { o[k] = JSON.parse(m[k].join('')); } catch (e) { o[k] = null; } }
  return { data: o, updated: t };
}
function doGet(e) {
  const p = (e && e.parameter) || {}, a = p.action || 'all';
  try {
    if (a === 'ping') return out_({ ok: true, time: new Date().toISOString(), sheet: SpreadsheetApp.getActive().getName() });
    if (a === 'get') { const r = readAll_(p.key); return out_({ ok: true, value: r.data[p.key] ?? null }); }
    const r = readAll_(p.prefix || ''); return out_({ ok: true, data: r.data, updated: r.updated });
  } catch (err) { return out_({ ok: false, error: String(err) }); }
}
function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.action !== 'set' || !b.key) throw new Error('Yêu cầu không hợp lệ');
    const sh = sh_(), v = sh.getDataRange().getValues();
    for (let i = v.length - 1; i >= 1; i--) if (v[i][0] === b.key) sh.deleteRow(i + 1);
    if (b.value !== null && b.value !== undefined) {
      const s = JSON.stringify(b.value), now = new Date(), rows = [];
      for (let i = 0, k = 0; i < s.length; i += PART, k++) rows.push([b.key, k, s.slice(i, i + PART), now]);
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
    }
    return out_({ ok: true, key: b.key });
  } catch (err) { return out_({ ok: false, error: String(err) }); }
  finally { lock.releaseLock(); }
}
