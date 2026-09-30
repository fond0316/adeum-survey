/**
 * CLASS BOARD 사용자 등록부 GAS — 아듬샘 전용 (Made by Cheon_R)
 * 다른 선생님들의 앱이 처음 로그인할 때 여기로 등록·확인 요청을 보내요.
 * 재배포 시 반드시 [배포 관리] → [수정] → 새 버전 (새 배포 금지)
 *
 * [시트 사용법]
 *  - '사용자' 탭: 승인 ☑ 체크 = 사용 허가 (승인제 ON일 때만 의미 있음)
 *                차단 ☑ 체크 = 사용 중지 (승인제와 상관없이 항상 적용)
 *  - '시트연결' 칸: 미연결 = 이름·학교만 등록하고 아직 구글 시트를 연결하지 않은 선생님 / 연결됨 = 시트 연결까지 끝낸 선생님 (자동 기록)
 *  - '설정' 탭: 승인제 → ON / OFF
 *               관리자설치ID → 아듬샘 앱 ⚙️ 탭에 보이는 '설치 ID'를 붙여넣기
 */

var TZ = 'Asia/Seoul';
var USERS = { name: '사용자', headers: ['설치ID', '등록일', '이름', '학교', '학급', '마지막사용', '승인', '차단', '메모', '시트연결'] };
var CONF = { name: '설정', headers: ['항목', '값'] };
var COL = { id: 1, reg: 2, name: 3, school: 4, cls: 5, last: 6, ok: 7, block: 8, conn: 10 };

function doGet(e) {
  var p = (e && e.parameter) || {};
  var cb = p.callback || 'callback';
  var out;
  try {
    setup_();
    if (p.action === 'check') out = check_(p);
    else if (p.action === 'register') out = register_(p);
    else if (p.action === 'ownerList') out = ownerList_(p);
    else if (p.action === 'ownerApproval') out = ownerApproval_(p);
    else if (p.action === 'ownerUser') out = ownerUser_(p);
    else out = { ok: false, error: '알 수 없는 요청이에요.' };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(cb + '(' + JSON.stringify(out) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function setup_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var u = ss.getSheetByName(USERS.name);
  if (!u) {
    u = ss.insertSheet(USERS.name);
    u.getRange(1, 1, 1, USERS.headers.length).setValues([USERS.headers]).setFontWeight('bold');
    u.setFrozenRows(1);
  }
  var c = ss.getSheetByName(CONF.name);
  if (!c) {
    c = ss.insertSheet(CONF.name);
    c.getRange(1, 1, 1, 2).setValues([CONF.headers]).setFontWeight('bold');
    c.getRange(2, 1, 2, 2).setValues([['승인제', 'OFF'], ['관리자설치ID', '']]);
    c.setFrozenRows(1);
  }
  // 예전에 만든 시트에도 '시트연결' 머리글을 자동으로 추가 (기존 데이터는 그대로)
  if (String(u.getRange(1, COL.conn).getValue()).trim() !== '시트연결') u.getRange(1, COL.conn).setValue('시트연결').setFontWeight('bold');
}

function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }
function conf_(key) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONF.name);
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (String(v[i][0]).trim() === key) return String(v[i][1]).trim();
  return '';
}
function usersSheet_() { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS.name); }
function findRow_(id) {
  var sh = usersSheet_();
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, COL.id, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]).trim() === id) return i + 2;
  return -1;
}
function isTrue_(v) { return v === true || String(v).toUpperCase() === 'TRUE' || v === 'Y' || v === 'O'; }

/** 상태 판단: owner / ok / pending / blocked / unregistered */
function status_(id) {
  if (id && id === conf_('관리자설치ID')) return 'owner';
  var row = findRow_(id);
  if (row < 0) return 'unregistered';
  var sh = usersSheet_();
  var r = sh.getRange(row, 1, 1, USERS.headers.length).getValues()[0];
  sh.getRange(row, COL.last).setValue(now_());
  if (id.indexOf('pre-') !== 0 && String(r[COL.conn - 1]).trim() !== '연결됨') sh.getRange(row, COL.conn).setValue('연결됨');
  if (isTrue_(r[COL.block - 1])) return 'blocked';
  if (conf_('승인제').toUpperCase() === 'ON' && !isTrue_(r[COL.ok - 1])) return 'pending';
  return 'ok';
}

/** 시트 연결 전 임시 ID('pre-…')로 등록해 둔 줄을 진짜 설치 ID로 바꿔 끼움 (승인·차단 상태 유지) */
function migrate_(id, old) {
  old = String(old || '').trim();
  if (!old || old === id || old.indexOf('pre-') !== 0 || id.indexOf('pre-') === 0) return;
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var oldRow = findRow_(old);
    if (oldRow < 0) return;
    if (findRow_(id) < 0) usersSheet_().getRange(oldRow, COL.id).setValue(id);
    else usersSheet_().deleteRow(oldRow);   // 이미 진짜 ID 줄이 있으면 임시 줄만 정리
  } finally { lock.releaseLock(); }
}

function check_(p) {
  var id = String(p.id || '').trim();
  if (!id) throw new Error('설치 ID가 없어요.');
  migrate_(id, p.old);
  var res = { ok: true, status: status_(id) };
  var row = findRow_(id);
  if (row > 0) {   // 다른 기기에서도 문의 메일 제목에 쓸 수 있게 등록 정보 반환
    var r = usersSheet_().getRange(row, 1, 1, USERS.headers.length).getValues()[0];
    res.name = String(r[COL.name - 1]); res.school = String(r[COL.school - 1]); res.cls = String(r[COL.cls - 1]);
  }
  return res;
}

/* ---------- 👑 아듬샘 전용 (관리자설치ID로만 가능) ---------- */
function owner_(p) {
  var id = String(p.id || '').trim(), o = conf_('관리자설치ID');
  if (!o || id !== o) throw new Error('관리자만 할 수 있어요.');
}
function setConf_(key, val) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONF.name);
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (String(v[i][0]).trim() === key) { sh.getRange(i + 1, 2).setValue(val); return; }
  sh.appendRow([key, val]);
}
function ownerList_(p) {
  owner_(p);
  var sh = usersSheet_(), last = sh.getLastRow(), users = [];
  if (last >= 2) {
    users = sh.getRange(2, 1, last - 1, USERS.headers.length).getValues()
      .filter(function (r) { return String(r[0]).trim(); })
      .map(function (r) {
        var f = function (v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm') : String(v); };
        return { id: String(r[0]).trim(), reg: f(r[1]), name: String(r[2]), school: String(r[3]), cls: String(r[4]),
          last: f(r[5]), ok: isTrue_(r[6]), block: isTrue_(r[7]), conn: String(r[COL.conn - 1]).trim() !== '미연결' };
      });
  }
  return { ok: true, approval: conf_('승인제').toUpperCase() === 'ON', users: users };
}
function ownerApproval_(p) {
  owner_(p);
  setConf_('승인제', p.on === '1' ? 'ON' : 'OFF');
  return { ok: true };
}
function ownerUser_(p) {
  owner_(p);
  var row = findRow_(String(p.target || '').trim());
  if (row < 0) throw new Error('선생님을 찾지 못했어요.');
  var col = p.field === 'block' ? COL.block : p.field === 'ok' ? COL.ok : 0;
  if (!col) throw new Error('잘못된 요청이에요.');
  usersSheet_().getRange(row, col).setValue(p.value === '1');
  return { ok: true };
}

function register_(p) {
  var id = String(p.id || '').trim();
  var name = String(p.name || '').trim(), school = String(p.school || '').trim(), cls = String(p.cls || '').trim();
  if (!id || !name || !school) throw new Error('이름과 학교를 입력해 주세요.');
  var isPre = id.indexOf('pre-') === 0;   // 시트 연결 전 사전 등록
  if (!isPre) migrate_(id, p.old);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = usersSheet_();
    var row = findRow_(id);
    if (row < 0) {
      sh.appendRow([id, now_(), name, school, cls, now_(), false, false, '', isPre ? '미연결' : '연결됨']);
      row = sh.getLastRow();
      sh.getRange(row, COL.ok, 1, 2).insertCheckboxes();
    } else {
      sh.getRange(row, COL.name, 1, 3).setValues([[name, school, cls]]);
    }
  } finally { lock.releaseLock(); }
  return { ok: true, status: isPre ? 'pre' : status_(id) };
}
