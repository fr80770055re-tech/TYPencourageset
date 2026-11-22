/**
 * 班級加分系統（校內帳號版）
 *
 * 老師、學生、行政都用學校 Google 帳號打開同一個網頁，系統依 Email 判斷身分：
 *   typs學號@typ.kh.edu.tw  → 學生，只看自己的點數
 *   「設定」裡的導師 Email   → 導師，給分＋看自己班＋撤回自己班任何一筆
 *   「設定」裡的行政 Email   → 行政，給分＋看所有班＋撤回自己給的
 *   其他 typt…@typ.kh.edu.tw → 科任，給分＋看、撤回自己給的
 *
 * 工作表：
 *   學生名單  班級 | 座號 | 姓名 | 學號
 *   設定      角色（導師／行政）| Email | 班級（導師才填）
 *   常用事由  給分畫面上的快速按鈕
 *   老師稱呼  老師第一次使用時自己填，學生看到的就是這個
 *   紀錄      每位學生一列；撤回只改「狀態」，不刪資料
 *
 * 部署：網頁應用程式，執行身分「我」，存取權「typ.kh.edu.tw 網域內的所有人」。
 */

const DOMAIN = 'typ.kh.edu.tw';
const STEP = 10; // 點數一律是 10 的倍數
const TZ = 'Asia/Taipei';

const SHEET = { roster: '學生名單', staff: '設定', reasons: '常用事由', names: '老師稱呼', records: '紀錄' };
const REC_HEADER = ['編號', '時間', '日期', '班級', '學號', '座號', '姓名', '點數', '事由', '給分者', '狀態', '撤回者', '撤回時間'];
const R = { id: 0, time: 1, date: 2, cls: 3, sid: 4, seat: 5, name: 6, points: 7, reason: 8, by: 9, status: 10, wBy: 11, wAt: 12 };
const WITHDRAWN = '已撤回';

// ───────────── 安裝 ─────────────

function onOpen() {
  SpreadsheetApp.getUi().createMenu('加分系統').addItem('初始設定（只需執行一次）', 'setup').addToUi();
}

// 建立缺少的工作表；已經存在的不會被動到，重跑也安全。
function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TZ);
  ensureSheet_(ss, SHEET.roster, ['班級', '座號', '姓名', '學號'], [], 'A:D');
  ensureSheet_(ss, SHEET.staff, ['角色', 'Email', '班級'],
    [['導師', Session.getEffectiveUser().getEmail(), '']], 'C:C');
  ensureSheet_(ss, SHEET.reasons, ['常用事由'],
    [['喜閱網闖關分數'], ['每週秩序獎'], ['每週整潔獎'], ['每週禮貌獎']]);
  ensureSheet_(ss, SHEET.names, ['Email', '稱呼'], []);
  ensureSheet_(ss, SHEET.records, REC_HEADER, [], 'C:F');
  const first = ss.getSheets()[0];
  if (first.getName() === '工作表1' || first.getName() === 'Sheet1') {
    if (first.getLastRow() === 0) ss.deleteSheet(first);
  }
  ss.toast('完成。請填「學生名單」，並在「設定」補上導師班級。', '加分系統', 8);
}

function ensureSheet_(ss, name, header, rows, textCols) {
  if (ss.getSheetByName(name)) return;
  const sh = ss.insertSheet(name);
  if (textCols) sh.getRange(textCols).setNumberFormat('@');
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  sh.setFrozenRows(1);
  if (rows.length) sh.getRange(2, 1, rows.length, header.length).setValues(rows);
}

// ───────────── 網頁 ─────────────

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('高雄市桃源國小全校加分系統')
    .setFaviconUrl('https://lh3.googleusercontent.com/d/1OUlbDSphilBFg7k1C_CoLjOkSxeuy-7C=s192#icon.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 網頁打開時呼叫一次，依身分回傳需要的資料。
function boot() {
  const u = whoAmI_();
  if (u.kind === 'student') return Object.assign({ kind: 'student' }, studentData_(u.student));
  if (u.kind !== 'staff') return { kind: 'none', email: u.email };
  const roster = readRoster_();
  const byClass = {};
  roster.forEach(s => { (byClass[s.cls] = byClass[s.cls] || []).push({ id: s.id, seat: s.seat, name: s.name }); });
  return {
    kind: 'staff',
    email: u.email,
    name: readNames_()[u.email] || '',
    reasons: readColumn_(SHEET.reasons),
    step: STEP,
    roster: byClass,
    viewClasses: u.viewClasses,
    homeroomClasses: u.homeroomClasses,
    today: today_(),
  };
}

// 老師第一次使用時設定學生看到的稱呼。
function setName(name) {
  const u = requireStaff_();
  name = String(name || '').trim().slice(0, 20);
  if (!name) throw new Error('請輸入稱呼。');
  withLock_(() => {
    const sh = sheet_(SHEET.names);
    const emails = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0]).toLowerCase()) : [];
    const i = emails.indexOf(u.email);
    if (i >= 0) sh.getRange(i + 2, 2).setValue(name);
    else sh.appendRow([u.email, name]);
  });
  return name;
}

// 給分：一次可給多位學生，每人一列紀錄。
function give(studentIds, reason, points) {
  const u = requireStaff_();
  reason = String(reason || '').trim().slice(0, 100);
  points = Number(points);
  if (!reason) throw new Error('請選擇或輸入事由。');
  if (!Number.isInteger(points) || points < STEP || points % STEP !== 0) throw new Error('點數必須是 ' + STEP + ' 的倍數。');
  const ids = [...new Set((studentIds || []).map(String))];
  if (!ids.length) throw new Error('請至少選一位學生。');
  const roster = readRoster_();
  const students = ids.map(id => roster.find(s => s.id === id));
  if (students.some(s => !s)) throw new Error('名單裡找不到部分學生，請重新整理頁面。');

  const now = new Date();
  const date = today_();
  const rows = students.map(s => [newId_(), now, date, s.cls, s.id, s.seat, s.name, points, reason, u.email, '', '', '']);
  withLock_(() => {
    const sh = sheet_(SHEET.records);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, REC_HEADER.length).setValues(rows);
  });
  return { count: rows.length, points: points };
}

// 自己給過的紀錄（最新的在前，最多 100 筆）。
function mine() {
  const u = requireStaff_();
  const names = readNames_();
  return readRecords_().filter(r => r.by === u.email).reverse().slice(0, 100).map(r => publicRecord_(r, names));
}

// 全班檢視：導師看自己班，行政看全部。
function classView(cls) {
  const u = requireStaff_();
  cls = String(cls || '');
  if (u.viewClasses.indexOf(cls) < 0) throw new Error('你沒有查看這個班的權限。');
  const names = readNames_();
  const students = readRoster_().filter(s => s.cls === cls);
  const records = readRecords_().filter(r => r.cls === cls);
  const active = records.filter(r => r.status !== WITHDRAWN);
  return {
    cls: cls,
    months: [...new Set(active.map(r => r.date.slice(0, 7)))].sort().reverse(),
    students: students.map(s => {
      const mineR = active.filter(r => r.sid === s.id);
      const byMonth = {};
      mineR.forEach(r => { const m = r.date.slice(0, 7); byMonth[m] = (byMonth[m] || 0) + r.points; });
      return { id: s.id, seat: s.seat, name: s.name, total: mineR.reduce((a, r) => a + r.points, 0), byMonth: byMonth };
    }),
    records: records.reverse().map(r => publicRecord_(r, names)),
  };
}

// 撤回：給分者本人，或該生的導師。只標記，不刪除。
function withdraw(recordId) {
  const u = requireStaff_();
  return withLock_(() => {
    const sh = sheet_(SHEET.records);
    const rec = readRecords_().find(r => r.id === String(recordId));
    if (!rec) throw new Error('找不到這筆紀錄。');
    if (rec.status === WITHDRAWN) return true;
    if (rec.by !== u.email && u.homeroomClasses.indexOf(rec.cls) < 0) throw new Error('只能撤回自己給的分。');
    sh.getRange(rec.row, R.status + 1, 1, 3).setValues([[WITHDRAWN, u.email, new Date()]]);
    return true;
  });
}

// ───────────── 身分 ─────────────

function whoAmI_() {
  const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  const parts = email.split('@');
  if (parts.length !== 2 || parts[1] !== DOMAIN) return { kind: 'none', email: email };
  const local = parts[0];

  const staff = readStaff_().filter(s => s.email === email);
  const homeroomClasses = staff.filter(s => s.role === '導師' && s.cls).map(s => s.cls);
  const isAdmin = staff.some(s => s.role === '行政');
  if (staff.length || /^typt/.test(local)) {
    const allClasses = [...new Set(readRoster_().map(s => s.cls))].sort();
    return {
      kind: 'staff', email: email, homeroomClasses: homeroomClasses,
      viewClasses: isAdmin ? allClasses : homeroomClasses,
    };
  }
  const m = /^typs(\d+)$/.exec(local);
  if (m) {
    const student = readRoster_().find(s => s.id === m[1]);
    if (student) return { kind: 'student', email: email, student: student };
  }
  return { kind: 'none', email: email };
}

function requireStaff_() {
  const u = whoAmI_();
  if (u.kind !== 'staff') throw new Error('只有老師可以使用這個功能。');
  return u;
}

// ───────────── 資料 ─────────────

function studentData_(s) {
  const names = readNames_();
  const items = readRecords_().filter(r => r.sid === s.id && r.status !== WITHDRAWN).reverse();
  return {
    seat: s.seat, name: s.name, cls: s.cls,
    total: items.reduce((a, r) => a + r.points, 0),
    items: items.map(r => ({ date: r.date, reason: r.reason, points: r.points, by: nameOf_(r.by, names) })),
  };
}

function publicRecord_(r, names) {
  return {
    id: r.id, date: r.date, cls: r.cls, sid: r.sid, seat: r.seat, name: r.name,
    points: r.points, reason: r.reason, byEmail: r.by, by: nameOf_(r.by, names),
    withdrawn: r.status === WITHDRAWN, wBy: r.wBy ? nameOf_(r.wBy, names) : '',
  };
}

function readRoster_() {
  return rows_(SHEET.roster)
    .map(r => ({ cls: str_(r[0]), seat: str_(r[1]), name: str_(r[2]), id: str_(r[3]) }))
    .filter(s => s.id && s.name);
}

function readStaff_() {
  return rows_(SHEET.staff)
    .map(r => ({ role: str_(r[0]), email: str_(r[1]).toLowerCase(), cls: str_(r[2]) }))
    .filter(s => s.email);
}

function readNames_() {
  const map = {};
  rows_(SHEET.names).forEach(r => { const e = str_(r[0]).toLowerCase(); if (e) map[e] = str_(r[1]); });
  return map;
}

function readColumn_(name) {
  return rows_(name).map(r => str_(r[0])).filter(Boolean);
}

function readRecords_() {
  return rows_(SHEET.records).map((r, i) => ({
    row: i + 2,
    id: str_(r[R.id]),
    date: r[R.date] instanceof Date ? Utilities.formatDate(r[R.date], TZ, 'yyyy/MM/dd') : str_(r[R.date]),
    cls: str_(r[R.cls]), sid: str_(r[R.sid]), seat: str_(r[R.seat]), name: str_(r[R.name]),
    points: Number(r[R.points]) || 0, reason: str_(r[R.reason]),
    by: str_(r[R.by]).toLowerCase(), status: str_(r[R.status]), wBy: str_(r[R.wBy]).toLowerCase(),
  })).filter(r => r.id);
}

function nameOf_(email, names) {
  return names[email] || String(email).split('@')[0];
}

function rows_(name) {
  const sh = sheet_(name);
  const n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, sh.getLastColumn()).getValues() : [];
}

function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('找不到工作表「' + name + '」，請先執行初始設定。');
  return sh;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function str_(v) { return String(v == null ? '' : v).trim(); }
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd'); }
function newId_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 10); }
