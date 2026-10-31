/**
 * FC Portal – kho dữ liệu chung (Google Sheet) cho các app FC trên ruralwcm.github.io
 * Phiên bản 1.1 · 31/10/2026 · by HaoDT
 *
 * CÁCH CÀI (làm 1 lần):
 * 1. Mở Google Sheet "FC Portal - Dữ liệu chung" → Tiện ích mở rộng → Apps Script.
 * 2. Xóa hết code cũ, dán toàn bộ file này, bấm Lưu (biểu tượng đĩa).
 * 3. Chọn hàm "testSetup" trên thanh công cụ → Chạy → cấp quyền (Xem lại quyền → chọn tài khoản →
 *    Nâng cao → Đi tới … (không an toàn) → Cho phép). Nhật ký hiện "OK" là đúng.
 * 4. Triển khai → Quản lý các lần triển khai → biểu tượng bút chì → Phiên bản: "Phiên bản mới" → Triển khai.
 *    (Lần đầu: Triển khai → Triển khai mới → loại "Ứng dụng web" ·
 *     Thực thi dưới dạng: "Tôi" · Người có quyền truy cập: "Bất kỳ ai".)
 * 5. Mở link /exec?action=ping trên trình duyệt: thấy {"ok":true,...} là các app kết nối được.
 */
const SHEET = 'kv';        // key | part | data | updatedAt
const PART = 45000;        // 1 ô tối đa 50.000 ký tự → chia nhỏ
const SHEET_ID = '1v3ij5RLe86KtA-IFPuJlHDnEG4BQBXeFCkfd3mFUcuw'; // dùng khi script không gắn với Sheet
const VERSION = '1.1';

function book_() {
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { ss = null; }
  return ss || SpreadsheetApp.openById(SHEET_ID);
}
function sh_() {
  const ss = book_();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.getRange(1, 1, 1, 4).setValues([['key', 'part', 'data', 'updatedAt']]);
    sh.setFrozenRows(1);
    sh.getRange('C:C').setNumberFormat('@');
  }
  return sh;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function readAll_(prefix) {
  const sh = sh_(), n = sh.getLastRow();
  const m = {}, t = {};
  if (n >= 2) {
    const v = sh.getRange(2, 1, n - 1, 4).getValues();
    for (let i = 0; i < v.length; i++) {
      const k = String(v[i][0] || '');
      if (!k || (prefix && k.indexOf(prefix) !== 0)) continue;
      (m[k] = m[k] || [])[Number(v[i][1]) || 0] = String(v[i][2]);
      t[k] = v[i][3];
    }
  }
  const o = {};
  for (const k in m) { try { o[k] = JSON.parse(m[k].join('')); } catch (e) { o[k] = null; } }
  return { data: o, updated: t };
}
function doGet(e) {
  const p = (e && e.parameter) || {}, a = p.action || 'all';
  try {
    if (a === 'ping') return out_({ ok: true, version: VERSION, time: new Date().toISOString(), sheet: book_().getName() });
    if (a === 'get') { const r = readAll_(p.key); return out_({ ok: true, value: (p.key in r.data) ? r.data[p.key] : null }); }
    const r = readAll_(p.prefix || '');
    return out_({ ok: true, data: r.data, updated: r.updated });
  } catch (err) { return out_({ ok: false, error: String(err) }); }
}
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.action !== 'set' || !b.key) throw new Error('Yêu cầu không hợp lệ');
    setKey_(String(b.key), b.value);
    return out_({ ok: true, key: b.key });
  } catch (err) { return out_({ ok: false, error: String(err) }); }
  finally { lock.releaseLock(); }
}
function setKey_(key, value) {
  const sh = sh_(), n = sh.getLastRow();
  // xóa các dòng cũ của key (gộp các dòng liền nhau để xóa nhanh)
  if (n >= 2) {
    const keys = sh.getRange(2, 1, n - 1, 1).getValues();
    for (let i = keys.length - 1; i >= 0; i--) {
      if (String(keys[i][0]) !== key) continue;
      let j = i; while (j - 1 >= 0 && String(keys[j - 1][0]) === key) j--;
      sh.deleteRows(j + 2, i - j + 1); i = j;
    }
  }
  if (value === null || value === undefined) return;
  const s = JSON.stringify(value), now = new Date(), rows = [];
  for (let i = 0, k = 0; i < s.length; i += PART, k++) rows.push([key, k, s.slice(i, i + PART), now]);
  const start = sh.getLastRow() + 1;
  sh.getRange(start, 3, rows.length, 1).setNumberFormat('@');
  sh.getRange(start, 1, rows.length, 4).setValues(rows);
}
/** Chạy tay 1 lần trong trình soạn thảo để cấp quyền và kiểm tra. */
function testSetup() {
  setKey_('_test/ping', { at: new Date().toISOString() });
  const r = readAll_('_test/');
  setKey_('_test/ping', null);
  Logger.log(r.data['_test/ping'] ? 'OK – Sheet: ' + book_().getName() : 'Lỗi: không đọc lại được');
}
