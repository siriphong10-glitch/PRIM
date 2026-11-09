/**
 * PRIM — Project & Resource Integrated Management
 * ฐานข้อมูลบน Google Sheets + ระบบเข้าสู่ระบบ (Apps Script Web App)
 *
 * วิธีติดตั้ง (ทำครั้งเดียว):
 *  1) เปิด Google Sheet "PRIM Database" → เมนู ส่วนขยาย (Extensions) → Apps Script
 *  2) ลบโค้ดเดิมทั้งหมด แล้ววางโค้ดไฟล์นี้ → กดบันทึก (รูปแผ่นดิสก์)
 *  3) กด "การทำให้ใช้งานได้ (Deploy)" → "การทำให้ใช้งานได้รายการใหม่ (New deployment)"
 *     → เลือกประเภท "เว็บแอป (Web app)"
 *     → Execute as: "ฉัน (Me)" / Who has access: "ทุกคน (Anyone)"
 *     → กด Deploy → อนุญาตสิทธิ์ (Authorize) → คัดลอก URL ที่ลงท้ายด้วย /exec
 *  4) เปิดเว็บ PRIM → วาง URL ในหน้าตั้งค่าการเชื่อมต่อ → ผู้ที่เข้าสู่ระบบคนแรกจะเป็นผู้ดูแลระบบ (admin)
 *
 * แผ่นงานที่ใช้: data (ข้อมูลโครงการ/บุคลากร), users (ผู้ใช้งาน), log (ประวัติการแก้ไข)
 * อัปเดตโค้ด: วางโค้ดใหม่ → บันทึก → การทำให้ใช้งานได้ → จัดการการทำให้ใช้งานได้ → ✏️ แก้ไข → เวอร์ชัน "เวอร์ชันใหม่" → การทำให้ใช้งานได้ (URL เดิมใช้ต่อได้)
 */
const SHEET_ID = '';  // ว่างไว้ = สร้าง Google Sheet "PRIM Database" ให้อัตโนมัติในบัญชีที่ติดตั้งสคริปต์
const CHUNK = 45000;                       // ขนาดข้อความต่อเซลล์ (เซลล์รับได้ไม่เกิน 50,000 ตัวอักษร)
const TOKEN_DAYS = 14;                     // อายุการเข้าสู่ระบบ (วัน)

function ss_() {
  const p = PropertiesService.getScriptProperties(); const id = p.getProperty('sheetId') || SHEET_ID;
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  const act = SpreadsheetApp.getActiveSpreadsheet(); if (act) { p.setProperty('sheetId', act.getId()); return act; }
  const ss = SpreadsheetApp.create('PRIM Database — Project & Resource Integrated Management'); p.setProperty('sheetId', ss.getId()); return ss;
}
function sheet_(name, header) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh && name === 'data') { const f = ss.getSheets()[0]; if (f && String(f.getRange(1, 1).getValue()) === 'key') { f.setName('data'); sh = f; } }
  if (!sh) { sh = ss.insertSheet(name); sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold'); sh.setFrozenRows(1); }
  else if (sh.getLastColumn() < header.length) sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  return sh;
}
const DATA_H = ['key', 'updatedAt', 'updatedBy', 'deleted', 'json'];
const USER_H = ['username', 'displayName', 'role', 'salt', 'passHash', 'active', 'createdAt', 'lastLogin', 'status', 'note'];
const LOG_H = ['time', 'user', 'action', 'key'];

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doGet(e) { return out_({ ok: true, app: 'PRIM', time: Date.now(), hint: 'POST only' }); }
function doPost(e) {
  let req = {};
  try { req = JSON.parse(e.postData.contents || '{}'); } catch (err) { return out_({ ok: false, error: 'bad_json' }); }
  try { return out_(route_(req)); } catch (err) { return out_({ ok: false, error: String(err && err.message || err) }); }
}

function hash_(salt, pass) {
  const d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pass, Utilities.Charset.UTF_8);
  return d.map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}
function users_() {
  const sh = sheet_('users', USER_H); const v = sh.getDataRange().getValues(); const out = [];
  for (let i = 1; i < v.length; i++) if (v[i][0]) out.push({ row: i + 1, username: String(v[i][0]), displayName: String(v[i][1]), role: String(v[i][2] || 'editor'), salt: String(v[i][3]), passHash: String(v[i][4]), active: v[i][5] !== false && String(v[i][5]).toUpperCase() !== 'FALSE', createdAt: v[i][6], lastLogin: v[i][7], status: String(v[i][8] || ''), note: String(v[i][9] || '') });
  return out;
}
function tokGet_(tok) { const v = PropertiesService.getScriptProperties().getProperty('tk_' + tok); try { return v ? JSON.parse(v) : null; } catch (e) { return null; } }
function tokPut_(tok, u) {
  const p = PropertiesService.getScriptProperties(); const now = Date.now(); const all = p.getProperties();
  Object.keys(all).forEach(k => { if (k.indexOf('tk_') === 0) { try { if (JSON.parse(all[k]).exp < now) p.deleteProperty(k); } catch (e) { p.deleteProperty(k); } } });
  p.setProperty('tk_' + tok, JSON.stringify({ u: u, exp: now + TOKEN_DAYS * 864e5 }));
}
function auth_(token) {
  if (!token) throw new Error('not_logged_in');
  const t = tokGet_(String(token)); if (!t || t.exp < Date.now()) throw new Error('not_logged_in');
  const u = users_().find(x => x.username === t.u && x.active); if (!u) throw new Error('not_logged_in');
  return u;
}
function log_(user, action, key) { try { sheet_('log', LOG_H).appendRow([new Date(), user, action, key || '']); } catch (e) {} }
function pub_(u) { return { username: u.username, displayName: u.displayName, role: u.role, active: u.active, lastLogin: u.lastLogin, status: u.status || '', note: u.note || '', createdAt: u.createdAt }; }

function readAll_(since) {
  const sh = sheet_('data', DATA_H); const v = sh.getDataRange().getValues(); const docs = [];
  for (let i = 1; i < v.length; i++) {
    const r = v[i]; if (!r[0]) continue; const ts = Number(r[1]) || 0;
    if (since && ts <= since) continue;
    let json = ''; for (let c = 4; c < r.length; c++) json += String(r[c] || '');
    docs.push({ key: String(r[0]), updatedAt: ts, updatedBy: String(r[2] || ''), deleted: r[3] === true || String(r[3]).toUpperCase() === 'TRUE', data: json && !(r[3] === true) ? JSON.parse(json) : null });
  }
  return docs;
}
function write_(key, data, by, del) {
  const sh = sheet_('data', DATA_H); const keys = sh.getRange(1, 1, Math.max(1, sh.getLastRow()), 1).getValues().map(r => String(r[0]));
  let row = keys.indexOf(key) + 1; if (row < 1) row = sh.getLastRow() + 1;
  const json = del ? '' : JSON.stringify(data); const parts = [];
  for (let i = 0; i < json.length; i += CHUNK) parts.push(json.slice(i, i + CHUNK));
  const ts = Date.now();
  const width = Math.max(sh.getLastColumn(), 4 + Math.max(1, parts.length));
  const vals = [key, ts, by, !!del].concat(parts.length ? parts : ['']); while (vals.length < width) vals.push('');
  sh.getRange(row, 1, 1, width).setNumberFormat('@').setValues([vals.map(String)]);
  sh.getRange(row, 2).setNumberFormat('0').setValue(ts); sh.getRange(row, 4).setValue(!!del);
  return ts;
}

// ข้อมูลเดิมที่ย้ายมาจากเวอร์ชันบน claude.ai (นำเข้าอัตโนมัติครั้งแรกเมื่อแผ่นงาน data ยังว่าง)
const SEED = [{"key":"projects/p_mw7gna2jg6cmh","data":{"code":"DW690801","color":"#2f6fed","createdAt":1794120205805,"dataDate":"2026-11-08","holidays":[],"id":"p_mw7gna2jg6cmh","name":"แผนงานโครงการตรวจสอบอาคารและโครงสร้าง เพื่อจัดทำแบบการดัดแปลงอาคาร โรงงาน บริษัท ไทยฮอนด้า จำกัด","nextTask":26,"owner":"","progressLog":[],"startDate":"2026-11-16","tasks":[{"actualFinish":"","actualStart":"","assignees":[],"daily":{},"dur":0,"id":7,"name":"งานเตรียมการ","note":"","parent":null,"pct":0,"plannedStart":"","preds":[],"type":"summary","wbs":"","weight":""},{"actualFinish":"","actualStart":"","assignees":[],"daily":{},"dur":10,"id":1,"name":"งานรวบรวมเอกสารและวางแผนโครงการ","note":"","parent":7,"pct":0,"plannedStart":"2026-11-16","preds":[],"type":"task","wbs":"1.1","weight":""},{"actualFinish":"","actualStart":"","assignees":[],"daily":{},"dur":14,"id":2,"name":"จัดเตรียมไฟล์ Auto CAD สำหรับการตั้งต้นเขียนแบบอาคาร (ฝ่ายโรงงานเป็นผู้จัดเตรียมและส่งมอบ)","note":"","parent":7,"pct":0,"plannedStart":"","preds":[{"id":1,"lag":0,"type":"FS"}],"type":"task","wbs":"1.2","weight":""},{"actualFinish":"","actualStart":"","assignees":[],"daily":{},"dur":0,"id":8,"name":"งานสำรวจส่วนอาคาร Sokie","note":"","parent":null,"pct":0,"plannedStart":"","preds":[],"type":"summary","wbs":"","weight":""},{"actualFinish":"","actualStart":"","assignees":[],"daily":{},"dur":21,"id":3,"name":"สำรวจการจัดทำระยะอาคาร-ความสูงอาคาร ส่วนต่อเติม Sokie","note":"","parent":8,"pct":0,"plannedStart":"","preds":[{"id":2,"lag":-1,"type":"FS"}],"type":"task","wbs":"2.1","weight":""},{"actualFinish":"","actualStart":"","assignees":[],"daily":{},"dur":14,"id":4,"name":"งานสำรวจรูปแบบโครงสร้างต่อเติม (ส่วนเหนือดิน-ฝ้า) ของส่วนอาคารต่อเติม Sokie","note":"","parent":8,"pct":0,"plannedStart":"","preds":[{"id":3,"lag":1,"type":"FS"}],"type":"task","wbs":"2.2","weight":""},{"actualFinish":"","actualStart":"","assignees":[],"daily":{},"dur":3,"id":5,"name":"งานจาะสำรวจดินข้างอาคาร Sokie","note":"","parent":8,"pct":0,"plannedStart":"","preds":[{"id":4,"lag":3,"type":"FS"}],"type":"task","wbs":"2.3","weight":""},{"actualFinish":"","actualStart":"","assignees":[],"daily":{},"dur":14,"id":10,"name":"งานขุดสำรวจรูปแบบโครงสร้างส่วนใต้ดินของส่วนต่อเติม","note":"","parent":8,"pct":0,"plannedStart":"","preds":[{"id":4,"lag":0,"type":"FS"}],"type":"task","wbs":"","weight":""},{"actualFinish":"","actualStart":"","assignees":[],"daily":{},"dur":7,"id":11,"name":"งานสำรวจระบบดับเพลิงเดิม เพื่อวางแผนจัดการหรือออกแบบการต่อเติมระบบ Sokie","note":"","parent":8,"pct":0,"plannedStart":"","preds":[{"id":3,"lag":-5,"type":"FS"}],"type":"task","wbs":"","weight":""},{"actualFinish":"","actualStart":"","assignees":[],"daily":{},"dur":0,"id":12,"name":"งานสำรวจส่วนอาคาร Factory 3","note":"","parent":null,"pct":0,"plannedStart":"","preds":[],"type":"summary","wbs":"","weight":""},{"actualFinish":"","actualStart":"","assignees":[],"daily":{},"dur":21,"id":13,"name":"สำรวจการจัดทำระยะอาคาร-ความสูงอาคาร ส่วนต่อเติม Factory 3","note":"","parent":12,"pct":0,"plannedStart":"","preds":[{"id":10,"lag":7,"type":"FS"}],"type":"task","wbs":"","weight":""},{"actualFinish":"","actualStart":"","assignees":[],"daily":{},"dur":14,"id":14,"name":"งานสำรวจรูปแบบโครงสร้างต่อเติม (ส่วนเหนือดิน-ฝ้า) ของส่วนอาคารต่อเติม Factory 3","note":"","parent":12,"pct":0,"plannedStart":"","preds":[{"id":13,"lag":0,"type":"FS"}],"type":"task","wbs":"","weight":""},{"actualFinish":"","actualStart":"","assignees":[],"daily":{},"dur":4,"id":15,"name":"งานเจาะสำรวจดินข้างอาคาร Factory 3","note":"","parent":12,"pct":0,"plannedStart":"","preds":[{"id":14,"lag":0,"type":"FS"}],"type":"task","wbs":"","weight":""},{"actualFinish":"","actualStart":"","assignees":[],"daily":{},"dur":7,"id":16,"name":"งานขุดและสำรวจรูปแบบโครงสร้างส่วนใต้ดินของส่วนต่อเติม อาคาร Factory 3","note":"","parent":12,"pct":0,"plannedStart":"","preds":[{"id":14,"lag":0,"type":"FS"}],"type":"task","wbs":"","weight":""},{"actualFinish":"","actualStart":"","assignees":[],"daily":{},"dur":0,"id":9,"name":"งานจัดทำแบบสำรวจ ออกแบบระบบ และเอกสารประกอบ","note":"","parent":null,"pct":0,"plannedStart":"","preds":[],"type":"summary","wbs":"","weight":""},{"actualFinish":"","actualStart":"","assignees":[],"daily":{},"dur":30,"id":6,"name":"งานออกแบบการจัดทำระบบดับเพลิงเพิ่มเติมสำหรับพื้นที่ต่อเติมของอาคาร Sokie","note":"","parent":9,"pct":0,"plannedStart":"","preds":[{"id":2,"lag":7,"type":"FS"}],"type":"task","wbs":"2.4","weight":""},{"actualFinish":"","actualStart":"","assignees":[],"daily":{},"dur":7,"id":17,"name":"งานสรุปรูปแบบโครงสร้างและแบบต่อเติมอาคาร Sokie","note":"","parent":9,"pct":0,"plannedStart":"","preds":[{"id":11,"lag":0,"type":"FS"}],"type":"task","wbs":"","weight":""},{"actualFinish":"","actualStart":"","assignees":[],"daily":{},"dur":5,"id":18,"name":"งานสรุปแบบโครงสร้างและส่วนต่อเติมอาคาร Factory 3","note":"","parent":9,"pct":0,"plannedStart":"","preds":[{"id":16,"lag":4,"type":"FS"}],"type":"task","wbs":"","weight":""},{"actualFinish":"","actualStart":"","assignees":[],"daily":{},"dur":28,"id":19,"name":"งานจัดทำแบบส่วนต่อเติมอาคาร Sokie","note":"","parent":9,"pct":0,"plannedStart":"","preds":[{"id":17,"lag":0,"type":"FS"}],"type":"task","wbs":"","weight":""},{"actualFinish":"","actualStart":"","assignees":[],"daily":{},"dur":28,"id":20,"name":"งานจัดทำแบบส่วนต่อเติมอาคาร Factory 3","note":"","parent":9,"pct":0,"plannedStart":"","preds":[{"id":18,"lag":0,"type":"FS"}],"type":"task","wbs":"","weight":""},{"actualFinish":"","actualStart":"","assignees":[],"daily":{},"dur":21,"id":21,"name":"งานจัดทำรายการคำนวณออกแบบชิ้นส่วนโครงสร้างอาคาร Sokie (เอกสารประกอบ)","note":"","parent":9,"pct":0,"plannedStart":"","preds":[{"id":17,"lag":0,"type":"FS"}],"type":"task","wbs":"","weight":""},{"actualFinish":"","actualStart":"","assignees":[],"daily":{},"dur":21,"id":22,"name":"งานจัดทำรายการคำนวณออกแบบชิ้นส่วนโครงสร้างอาคาร Factory 3 (เอกสารประกอบ)","note":"","parent":9,"pct":0,"plannedStart":"","preds":[{"id":18,"lag":0,"type":"FS"}],"type":"task","wbs":"","weight":""},{"actualFinish":"","actualStart":"","assignees":[],"daily":{},"dur":14,"id":23,"name":"ส่งแบบและเอกสารประกอบให้แก่โรงงาน เพื่อตรวจสอบความถูกต้องรอบที่ 1","note":"","parent":9,"pct":0,"plannedStart":"","preds":[{"id":20,"lag":0,"type":"FS"},{"id":19,"lag":0,"type":"FS"}],"type":"task","wbs":"","weight":""},{"actualFinish":"","actualStart":"","assignees":[],"daily":{},"dur":21,"id":24,"name":"รวบรวมเอกสารและข้อมูลประกอบการขออนุญาตดัดแปลง","note":"","parent":9,"pct":0,"plannedStart":"","preds":[{"id":23,"lag":0,"type":"FS"}],"type":"task","wbs":"","weight":""},{"actualFinish":"","actualStart":"","assignees":[],"daily":{},"dur":14,"id":25,"name":"ส่งแบบแก้ไขและยื่นเอกสารขออนุญาตต่อหน่วยงานราชการ","note":"","parent":9,"pct":0,"plannedStart":"","preds":[{"id":24,"lag":5,"type":"FS"}],"type":"task","wbs":"","weight":""}],"updatedAt":1794124368969,"workdays":[1,2,3,4,5,6,0]}},{"key":"team/main","data":{"members":[{"active":true,"color":"#c4ed31","dept":"โครงสร้าง","email":"","id":"m_mw7h7hcxo1iqa","leaves":[],"name":"ปฏิภาณ มูลศรี","phone":"","role":"วิศวกรโยธา"},{"active":true,"color":"#1f509e","dept":"โครงสร้าง","email":"","id":"m_mw7h848m6kmqa","leaves":[],"name":"ราชย์ มณี","phone":"","role":"วิศวกรโยธา"},{"active":true,"color":"#e67c73","dept":"โครงสร้าง","email":"","id":"m_mw7k1tufkl05y","leaves":[],"name":"สิริพงศ์ เกิดบุญมา","phone":"","role":"ผู้จัดการวิศวกรรมโครงสร้าง"}],"updatedAt":1794125807008}}];
function seedIfEmpty_() {
  const p = PropertiesService.getScriptProperties(); if (p.getProperty('seeded')) return;
  const sh = sheet_('data', DATA_H);
  if (sh.getLastRow() <= 1) SEED.forEach(d => write_(d.key, d.data, 'migrated', false));
  p.setProperty('seeded', '1');
}

/* ═══════ ปฏิทิน Google (ลงนัดเมื่อผู้ใช้กดสั่งเท่านั้น) ═══════ */
const CAL_H = ['key', 'eventId', 'hash', 'batch', 'createdBy', 'updatedAt', 'title', 'start'];
const CAL_NAME = 'PRIM — ตารางงาน';
function authorizeCalendar() { cal_(); return 'ok'; }   // กด "เรียกใช้" ฟังก์ชันนี้ 1 ครั้งเพื่ออนุญาตสิทธิ์ปฏิทิน
function cal_() {
  const p = PropertiesService.getScriptProperties(); const id = p.getProperty('calendarId');
  if (id) { const c = CalendarApp.getCalendarById(id); if (c) return c; }
  const found = CalendarApp.getCalendarsByName(CAL_NAME)[0];
  const c = found || CalendarApp.createCalendar(CAL_NAME, { timeZone: 'Asia/Bangkok', color: CalendarApp.Color.BLUE });
  p.setProperty('calendarId', c.getId()); return c;
}
function calRows_() {
  const sh = sheet_('calendar', CAL_H); const v = sh.getDataRange().getValues(); const out = {};
  for (let i = 1; i < v.length; i++) if (v[i][0]) out[String(v[i][0])] = { row: i + 1, key: String(v[i][0]), eventId: String(v[i][1]), hash: String(v[i][2]), batch: String(v[i][3]), title: String(v[i][6]), start: String(v[i][7]) };
  return out;
}
function calStatus_(prefix) { const r = calRows_(); return Object.values(r).filter(x => x.key.indexOf(prefix) === 0).map(x => ({ key: x.key, hash: x.hash, batch: x.batch, title: x.title, start: x.start })); }
function setGuests_(ev, emails) {
  const want = new Set((emails || []).map(e => String(e).toLowerCase())); const have = ev.getGuestList().map(g => g.getEmail().toLowerCase());
  have.forEach(e => { if (!want.has(e)) ev.removeGuest(e); }); want.forEach(e => { if (have.indexOf(e) < 0) ev.addGuest(e); });
}
function calPush_(q, user) {
  const c = cal_(); const sh = sheet_('calendar', CAL_H); const rows = calRows_(); const batch = 'b' + Date.now();
  const res = { created: 0, updated: 0, deleted: 0, batch: batch, errors: [] };
  (q.deletes || []).forEach(key => {
    const r = rows[key]; if (!r) return;
    try { const ev = c.getEventById(r.eventId); if (ev) ev.deleteEvent(); } catch (e) {}
    sh.getRange(r.row, 1, 1, CAL_H.length).clearContent(); res.deleted++;
  });
  (q.events || []).slice(0, 300).forEach(e => {
    try {
      const st = new Date(e.start), en = new Date(e.end); const r = rows[e.key];
      let ev = null; if (r && r.eventId) { try { ev = c.getEventById(r.eventId); } catch (x) { ev = null; } }
      if (ev) {
        ev.setTitle(e.title); ev.setTime(st, en); ev.setDescription(e.description || ''); if (q.invite) setGuests_(ev, e.guests);
        sh.getRange(r.row, 1, 1, CAL_H.length).setValues([[e.key, ev.getId(), e.hash, r.batch, user, new Date(), e.title, e.start]]); res.updated++;
      } else {
        const opt = { description: e.description || '' }; if (q.invite && e.guests && e.guests.length) { opt.guests = e.guests.join(','); opt.sendInvites = true; }
        ev = c.createEvent(e.title, st, en, opt);
        const vals = [[e.key, ev.getId(), e.hash, batch, user, new Date(), e.title, e.start]];
        if (r) sh.getRange(r.row, 1, 1, CAL_H.length).setValues(vals); else sh.appendRow(vals[0]);
        res.created++;
      }
    } catch (err) { res.errors.push(e.key + ': ' + err.message); }
  });
  return res;
}
function calUndo_(batch) {
  const c = cal_(); const sh = sheet_('calendar', CAL_H); let n = 0;
  Object.values(calRows_()).filter(r => r.batch === batch).forEach(r => { try { const ev = c.getEventById(r.eventId); if (ev) ev.deleteEvent(); } catch (e) {} sh.getRange(r.row, 1, 1, CAL_H.length).clearContent(); n++; });
  return n;
}

function route_(q) {
  const a = q.action;
  if (a === 'ping') return { ok: true, time: Date.now(), needSetup: users_().length === 0 };
  if (a === 'login') {
    const lock = LockService.getScriptLock(); lock.waitLock(10000);
    try {
      const us = users_(); const name = String(q.username || '').trim().toLowerCase(); const pass = String(q.password || '');
      if (!name || pass.length < 4) return { ok: false, error: 'กรอกชื่อผู้ใช้และรหัสผ่าน (อย่างน้อย 4 ตัวอักษร)' };
      let u = us.find(x => x.username.toLowerCase() === name);
      if (!us.length) { // first user becomes admin
        const salt = Utilities.getUuid(); sheet_('users', USER_H).appendRow([name, q.displayName || name, 'admin', salt, hash_(salt, pass), true, new Date(), new Date()]);
        u = users_()[0]; log_(name, 'create_admin');
      } else if (!u || hash_(u.salt, pass) !== u.passHash) return { ok: false, error: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
      else if (u.status === 'pending') return { ok: false, error: 'บัญชีนี้ส่งคำขอแล้ว กำลังรอผู้ดูแลระบบอนุมัติ' };
      else if (u.status === 'rejected') return { ok: false, error: 'คำขอสมัครใช้งานนี้ไม่ได้รับอนุมัติ — ติดต่อผู้ดูแลระบบ' };
      else if (!u.active) return { ok: false, error: 'บัญชีนี้ถูกปิดการใช้งาน — ติดต่อผู้ดูแลระบบ' };
      const tok = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, ''); tokPut_(tok, u.username);
      sheet_('users', USER_H).getRange(u.row, 8).setValue(new Date()); log_(u.username, 'login');
      return { ok: true, token: tok, user: pub_(u) };
    } finally { lock.releaseLock(); }
  }
  if (a === 'register') {
    const lock = LockService.getScriptLock(); lock.waitLock(10000);
    try {
      const us = users_(); if (!us.length) return { ok: false, error: 'ระบบยังไม่มีผู้ดูแล — กรุณาสร้างบัญชีผู้ดูแลระบบจากหน้าเข้าสู่ระบบก่อน' };
      const name = String(q.username || '').trim().toLowerCase(); const dn = String(q.displayName || '').trim().slice(0, 80);
      if (!/^[a-z0-9._-]{2,40}$/.test(name)) return { ok: false, error: 'ชื่อผู้ใช้ใช้ได้เฉพาะ a-z 0-9 . _ - (2–40 ตัว)' };
      if (!dn) return { ok: false, error: 'กรุณากรอกชื่อ-นามสกุล' };
      if (String(q.password || '').length < 4) return { ok: false, error: 'รหัสผ่านต้องมีอย่างน้อย 4 ตัวอักษร' };
      if (us.some(x => x.username.toLowerCase() === name)) return { ok: false, error: 'ชื่อผู้ใช้นี้มีผู้ใช้แล้ว กรุณาเลือกชื่ออื่น' };
      if (us.filter(x => x.status === 'pending').length >= 50) return { ok: false, error: 'มีคำขอรออนุมัติจำนวนมาก กรุณาติดต่อผู้ดูแลระบบ' };
      const salt = Utilities.getUuid();
      sheet_('users', USER_H).appendRow([name, dn, 'viewer', salt, hash_(salt, String(q.password)), false, new Date(), '', 'pending', String(q.note || '').slice(0, 200)]);
      log_(name, 'register'); return { ok: true };
    } finally { lock.releaseLock(); }
  }
  const u = auth_(q.token);
  if (a === 'me') return { ok: true, user: pub_(u) };
  if (a === 'pendingCount') return { ok: true, count: u.role === 'admin' ? users_().filter(x => x.status === 'pending').length : 0 };
  if (a === 'logout') { PropertiesService.getScriptProperties().deleteProperty('tk_' + q.token); return { ok: true }; }
  if (a === 'all') { seedIfEmpty_(); }
  if (a === 'all') return { ok: true, time: Date.now(), docs: readAll_(0), user: pub_(u) };
  if (a === 'since') return { ok: true, time: Date.now(), docs: readAll_(Number(q.since) || 0) };
  if (a === 'changePassword') {
    if (hash_(u.salt, String(q.oldPassword || '')) !== u.passHash) return { ok: false, error: 'รหัสผ่านเดิมไม่ถูกต้อง' };
    if (String(q.newPassword || '').length < 4) return { ok: false, error: 'รหัสผ่านใหม่ต้องมีอย่างน้อย 4 ตัวอักษร' };
    const salt = Utilities.getUuid(); sheet_('users', USER_H).getRange(u.row, 4, 1, 2).setValues([[salt, hash_(salt, String(q.newPassword))]]); log_(u.username, 'change_password');
    return { ok: true };
  }
  if (a === 'set' || a === 'delete') {
    if (u.role === 'viewer') return { ok: false, error: 'read_only' };
    if (!/^(projects\/[\w.:@+~-]+|team\/main)$/.test(String(q.key))) return { ok: false, error: 'bad_key' };
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try { const ts = write_(String(q.key), q.data, u.username, a === 'delete'); log_(u.username, a, q.key); return { ok: true, updatedAt: ts }; }
    finally { lock.releaseLock(); }
  }
  if (a === 'calStatus') return { ok: true, items: calStatus_(String(q.prefix || '')), calendar: CAL_NAME };
  if (a === 'calPush' || a === 'calUndo') {
    if (u.role === 'viewer') return { ok: false, error: 'read_only' };
    const lock = LockService.getScriptLock(); lock.waitLock(25000);
    try {
      if (a === 'calPush') { const r = calPush_(q, u.username); log_(u.username, 'calendar_push', (q.events || []).length + ' events / ' + (q.deletes || []).length + ' deletes'); return Object.assign({ ok: true, calendar: CAL_NAME }, r); }
      const n = calUndo_(String(q.batch || '')); log_(u.username, 'calendar_undo', q.batch); return { ok: true, deleted: n };
    } finally { lock.releaseLock(); }
  }
  if (a === 'users') { if (u.role !== 'admin') return { ok: false, error: 'admin_only' }; return { ok: true, users: users_().map(pub_) }; }
  if (a === 'saveUser') {
    if (u.role !== 'admin') return { ok: false, error: 'admin_only' };
    const name = String(q.username || '').trim().toLowerCase(); if (!/^[a-z0-9._-]{2,40}$/.test(name)) return { ok: false, error: 'ชื่อผู้ใช้ใช้ได้เฉพาะ a-z 0-9 . _ - (2–40 ตัว)' };
    const role = ['admin', 'editor', 'viewer'].includes(q.role) ? q.role : 'editor';
    const sh = sheet_('users', USER_H); const ex = users_().find(x => x.username === name);
    if (ex) {
      if (ex.username === u.username && role !== 'admin') return { ok: false, error: 'เปลี่ยนสิทธิ์ของตัวเองไม่ได้' };
      sh.getRange(ex.row, 2, 1, 2).setValues([[q.displayName || name, role]]); sh.getRange(ex.row, 6).setValue(q.active !== false);
      if (q.decision === 'approve') { sh.getRange(ex.row, 6).setValue(true); sh.getRange(ex.row, 9).setValue(''); }
      else if (q.decision === 'reject') { sh.getRange(ex.row, 6).setValue(false); sh.getRange(ex.row, 9).setValue('rejected'); }
      else if (ex.status === 'pending') sh.getRange(ex.row, 6).setValue(false);
      if (q.password) { if (String(q.password).length < 4) return { ok: false, error: 'รหัสผ่านอย่างน้อย 4 ตัวอักษร' }; const salt = Utilities.getUuid(); sh.getRange(ex.row, 4, 1, 2).setValues([[salt, hash_(salt, String(q.password))]]); }
    } else {
      if (String(q.password || '').length < 4) return { ok: false, error: 'ต้องกำหนดรหัสผ่านเริ่มต้น (อย่างน้อย 4 ตัวอักษร)' };
      const salt = Utilities.getUuid(); sh.appendRow([name, q.displayName || name, role, salt, hash_(salt, String(q.password)), true, new Date(), '', '', '']);
    }
    log_(u.username, q.decision ? q.decision + '_user' : 'save_user', name); return { ok: true, users: users_().map(pub_) };
  }
  return { ok: false, error: 'unknown_action' };
}
