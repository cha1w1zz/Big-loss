// Big Loss — ตัวเชื่อม Google Sheets (วางไฟล์นี้ที่ script.google.com หรือ Extensions > Apps Script)
function doGet() {
  const rows = sheet_().getDataRange().getValues().slice(1);
  return out_(rows.map(r => ({ id: String(r[0]), t: r[1], a: Number(r[2]), n: String(r[3]), d: Number(r[4]) })));
}

function doPost(e) {
  const b = JSON.parse(e.postData.contents);
  const s = sheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (b.op === 'add') {
      const a = Number(b.a);
      if (!(a > 0 && a < 1e7)) return out_({ ok: false });
      let n = String(b.n || '').slice(0, 200);
      if (/^[=+\-@]/.test(n)) n = "'" + n; // กันสูตรแปลกๆ ในชีต
      s.appendRow([Utilities.getUuid(), b.t === 'bet' ? 'bet' : 'smoke', a, n, Date.now()]);
    } else if (b.op === 'del') {
      const ids = s.getRange(1, 1, s.getLastRow(), 1).getValues();
      for (let i = ids.length - 1; i > 0; i--) if (String(ids[i][0]) === String(b.id)) { s.deleteRow(i + 1); break; }
    }
  } finally { lock.releaseLock(); }
  return out_({ ok: true });
}

function sheet_() {
  // ใช้ได้ทั้งแบบเปิดจากในชีต หรือสร้างที่ script.google.com (จะสร้างชีต "Big Loss" ให้เอง)
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    const props = PropertiesService.getScriptProperties();
    const id = props.getProperty('SHEET_ID');
    if (id) ss = SpreadsheetApp.openById(id);
    else { ss = SpreadsheetApp.create('Big Loss'); props.setProperty('SHEET_ID', ss.getId()); }
  }
  let s = ss.getSheetByName('data');
  if (!s) { s = ss.insertSheet('data'); s.appendRow(['id', 'type', 'amount', 'note', 'time']); }
  return s;
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
