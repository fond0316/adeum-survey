/**
 * 학생공지 GAS 백엔드 — Made by Cheon_R
 * 모든 요청은 JSONP(doGet + callback) 방식
 * 재배포 시 반드시 [배포 관리] → [수정] → 새 버전 (새 배포 금지)
 */

var TZ = 'Asia/Seoul';
var SHEETS = {
  settings: { name: '설정', headers: ['항목', '값'] },
  students: { name: '학생명단', headers: ['학번', '이름'] },
  notices: { name: '공지', headers: ['ID', '유형', '날짜', '제목', '내용', '대상메모', '중요', '팝업', '작성자', '작성시각', '수정시각'] },
  targets: { name: '호출대상', headers: ['공지ID', '학번', '이름'] },
  confirms: { name: '확인기록', headers: ['공지ID', '학번', '이름', '확인시각'] },
  requests: { name: '수정요청', headers: ['ID', '공지ID', '학번', '이름', '내용', '작성시각', '처리', '처리시각'] }
};
var PRESIDENT_TYPES = ['확인', '교과', '일반'];
var DEFAULT_PW = 'classboard';   // 초기 비밀번호 — 첫 로그인 때 변경 강제
var DEFAULT_SETTINGS = [['학년도', String(new Date().getFullYear())], ['학급명', '우리 반'], ['관리자비번', DEFAULT_PW], ['회장코드', '']];

/* ---------- 진입점 ---------- */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var cb = p.callback || 'callback';
  var out;
  try {
    setup_();
    out = route_(p);
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService
    .createTextOutput(cb + '(' + JSON.stringify(out) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function route_(p) {
  switch (p.action) {
    case 'ping': return { ok: true, className: getSetting_('학급명'), year: year_() };
    case 'load': return load_(p);
    case 'who': return who_(p);
    case 'confirm': return confirm_(p);
    case 'request': return request_(p);
    case 'resolveRequest': return resolveRequest_(p);
    case 'login': return login_(p);
    case 'adminLoad': return adminLoad_(p);
    case 'saveNotice': return saveNotice_(p);
    case 'saveNotices': return saveNotices_(p);
    case 'deleteNotice': return deleteNotice_(p);
    case 'saveStudents': return saveStudents_(p);
    case 'saveSettings': return saveSettings_(p);
    default: return { ok: false, error: '알 수 없는 요청이에요.' };
  }
}

/* ---------- 시트 준비 (탭이 없으면 헤더 포함 자동 생성) ---------- */
function setup_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(function (k) {
    var def = SHEETS[k];
    var sh = ss.getSheetByName(def.name);
    if (!sh) {
      sh = ss.insertSheet(def.name);
      sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold');
      sh.getRange(1, 1, 1000, def.headers.length).setNumberFormat('@'); // 날짜·학번 자동변환 방지
      sh.setFrozenRows(1);
      if (k === 'settings') sh.getRange(2, 1, DEFAULT_SETTINGS.length, 2).setValues(DEFAULT_SETTINGS);
    }
  });
}

/** 탭 헤더가 지워졌을 때 수동 실행용 */
function 헤더복구() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(function (k) {
    var def = SHEETS[k];
    var sh = ss.getSheetByName(def.name);
    if (sh) sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold');
  });
  setup_();
}

/* ---------- 공통 유틸 ---------- */
function sh_(k) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS[k].name); }
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function str_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return v === null || v === undefined ? '' : String(v).trim();
}
function rows_(k) {
  var sh = sh_(k);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var n = SHEETS[k].headers.length;
  return sh.getRange(2, 1, last - 1, n).getValues().map(function (r) { return r.map(str_); });
}
function getSetting_(key) {
  var r = rows_('settings').filter(function (x) { return x[0] === key; })[0];
  return r ? r[1] : '';
}
function setSetting_(key, val) {
  var sh = sh_('settings');
  var data = rows_('settings');
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === key) { sh.getRange(i + 2, 2).setValue(String(val)); return; }
  }
  sh.appendRow([key, String(val)]);
}
/** 설치 ID — 이 시트(선생님)를 구분하는 고유값, 없으면 자동 생성 */
function installId_() {
  var v = getSetting_('설치ID');
  if (!v) { v = Utilities.getUuid(); setSetting_('설치ID', v); }
  return v;
}
function login_(p) {
  var role = auth_(p.pw);
  if (!role) return { ok: true, role: '' };
  return { ok: true, role: role, installId: installId_(), className: getSetting_('학급명'),
    mustChange: role === 'admin' && String(p.pw) === DEFAULT_PW };
}
function year_() { return getSetting_('학년도') || Utilities.formatDate(new Date(), TZ, 'yyyy'); }
function auth_(pw) {
  pw = String(pw || '');
  if (!pw) return '';
  if (pw === getSetting_('관리자비번')) return 'admin';
  var code = getSetting_('회장코드');
  if (code && pw === code) return 'president';
  return '';
}
function need_(p, roles) {
  var role = auth_(p.pw);
  if (roles.indexOf(role) < 0) throw new Error('권한이 없어요. 비밀번호(코드)를 확인해 주세요.');
  return role;
}
function lock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function noticeObj_(r, targetMap) {
  return {
    id: r[0], type: r[1], date: r[2], title: r[3], body: r[4], memo: r[5],
    important: r[6] === 'Y', popup: r[7] === 'Y', author: r[8], created: r[9], updated: r[10],
    targets: targetMap[r[0]] || []
  };
}
function targetMap_() {
  var m = {};
  rows_('targets').forEach(function (r) {
    (m[r[0]] = m[r[0]] || []).push({ sid: r[1], name: r[2] });
  });
  return m;
}
function daysAgo_(n) {
  var d = new Date(Date.now() - n * 86400000);
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}

/* ---------- 학생용 ---------- */
function load_(p) {
  var sid = String(p.sid || '');
  var from = daysAgo_(60);
  var tm = targetMap_();
  var notices = rows_('notices')
    .filter(function (r) { return r[0] && r[2] >= from; })
    .map(function (r) { return noticeObj_(r, tm); });
  var mine = sid ? rows_('confirms')
    .filter(function (r) { return r[1] === sid; })
    .map(function (r) { return r[0]; }) : [];
  return { ok: true, className: getSetting_('학급명'), year: year_(), notices: notices, confirmed: mine, serverTime: now_() };
}

function who_(p) {
  var sid = String(p.sid || '');
  var r = rows_('students').filter(function (x) { return x[0] === sid; })[0];
  if (!r) return { ok: false, error: '명단에 없는 학번이에요. 담임선생님께 확인해 주세요.' };
  return { ok: true, sid: r[0], name: r[1] };
}

function confirm_(p) {
  var sid = String(p.sid || ''), nid = String(p.nid || '');
  var st = rows_('students').filter(function (x) { return x[0] === sid; })[0];
  if (!st) return { ok: false, error: '명단에 없는 학번이에요.' };
  return lock_(function () {
    var dup = rows_('confirms').some(function (r) { return r[0] === nid && r[1] === sid; });
    if (!dup) sh_('confirms').appendRow([nid, sid, st[1], now_()]);
    return { ok: true };
  });
}

/** 학생 수정요청·질문 — 이름은 명단에서 자동으로 붙임 (사칭 방지) */
function request_(p) {
  var sid = String(p.sid || ''), nid = String(p.nid || ''), text = String(p.text || '').trim();
  if (!text) return { ok: false, error: '내용을 입력해 주세요.' };
  if (text.length > 300) text = text.slice(0, 300);
  var st = rows_('students').filter(function (x) { return x[0] === sid; })[0];
  if (!st) return { ok: false, error: '명단에 없는 학번이에요. 학번 등록을 다시 해 주세요.' };
  var exists = rows_('notices').some(function (r) { return r[0] === nid; });
  if (!exists) return { ok: false, error: '삭제된 공지예요.' };
  return lock_(function () {
    sh_('requests').appendRow(['R' + Date.now(), nid, sid, st[1], text, now_(), '', '']);
    return { ok: true, name: st[1] };
  });
}

function resolveRequest_(p) {
  var role = need_(p, ['admin', 'president']);
  var rid = String(p.rid || ''), undo = p.undo === '1';
  return lock_(function () {
    var sh = sh_('requests');
    var data = rows_('requests');
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === rid) {
        if (role === 'president') {
          var n = rows_('notices').filter(function (r) { return r[0] === data[i][1]; })[0];
          if (!n || n[8] !== '회장') throw new Error('회장이 쓴 공지의 요청만 처리할 수 있어요.');
        }
        sh.getRange(i + 2, 7, 1, 2).setValues([[undo ? '' : 'Y', undo ? '' : now_()]]);
        return { ok: true };
      }
    }
    throw new Error('요청을 찾지 못했어요.');
  });
}

/* ---------- 관리자·회장용 ---------- */
function adminLoad_(p) {
  var role = need_(p, ['admin', 'president']);
  var tm = targetMap_();
  var notices = rows_('notices').filter(function (r) { return r[0]; })
    .map(function (r) { return noticeObj_(r, tm); });
  var res = { ok: true, role: role, className: getSetting_('학급명'), year: year_(), notices: notices };
  var presNids = {};
  notices.forEach(function (n) { if (n.author === '회장') presNids[n.id] = true; });
  res.requests = rows_('requests').filter(function (r) { return r[0] && (role === 'admin' || presNids[r[1]]); })
    .map(function (r) { return { id: r[0], nid: r[1], sid: r[2], name: r[3], text: r[4], at: r[5], done: r[6] === 'Y', doneAt: r[7] }; });
  if (role === 'admin') {
    res.students = rows_('students').filter(function (r) { return r[0]; })
      .map(function (r) { return { sid: r[0], name: r[1] }; });
    res.confirms = rows_('confirms').map(function (r) { return { nid: r[0], sid: r[1], name: r[2], at: r[3] }; });
    res.presidentCode = getSetting_('회장코드');
    res.installId = installId_();
  }
  return res;
}

function saveNotice_(p) {
  var role = need_(p, ['admin', 'president']);
  var n = JSON.parse(p.data || '{}');
  var author = role === 'admin' ? '담임' : '회장';
  if (role === 'president') {
    if (PRESIDENT_TYPES.indexOf(n.type) < 0) throw new Error('회장은 확인사항·교과공지·일반만 작성할 수 있어요.');
    n.popup = false; n.targets = [];
  }
  if (!n.title || !n.date || !n.type) throw new Error('유형, 날짜, 제목은 꼭 입력해 주세요.');

  return lock_(function () {
    var sh = sh_('notices');
    var data = rows_('notices');
    var idx = -1;
    if (n.id) {
      for (var i = 0; i < data.length; i++) if (data[i][0] === String(n.id)) { idx = i; break; }
      if (idx < 0) throw new Error('수정할 공지를 찾지 못했어요.');
      if (role === 'president' && data[idx][8] !== '회장') throw new Error('본인이 쓴 공지만 수정할 수 있어요.');
    }
    var id = n.id ? String(n.id) : 'N' + Date.now();
    var row = [id, n.type, n.date, n.title, n.body || '', n.memo || '',
      n.important ? 'Y' : '', n.popup ? 'Y' : '',
      idx >= 0 ? data[idx][8] : author,
      idx >= 0 ? data[idx][9] : now_(),
      idx >= 0 ? now_() : ''];
    if (idx >= 0) sh.getRange(idx + 2, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);

    replaceTargets_(id, n.targets || []);
    return { ok: true, id: id };
  });
}

/** 목록 붙여넣기 일괄 등록 (호출 대상·팝업 없음) */
function saveNotices_(p) {
  var role = need_(p, ['admin', 'president']);
  var list = JSON.parse(p.data || '[]');
  var author = role === 'admin' ? '담임' : '회장';
  var allowed = role === 'admin' ? ['긴급', '확인', '호출', '교과', '일반'] : PRESIDENT_TYPES;
  var rows = list.map(function (n, i) {
    if (allowed.indexOf(n.type) < 0) throw new Error('올릴 수 없는 유형이 있어요.');
    if (!n.title || !n.date) throw new Error('날짜와 제목이 빠진 줄이 있어요.');
    return ['N' + Date.now() + '_' + i, n.type, n.date, n.title, n.body || '', '', n.important ? 'Y' : '', '', author, now_(), ''];
  });
  if (!rows.length) return { ok: true, count: 0 };
  return lock_(function () {
    var sh = sh_('notices');
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setNumberFormat('@').setValues(rows);
    return { ok: true, count: rows.length };
  });
}

function replaceTargets_(nid, targets) {
  var sh = sh_('targets');
  var data = rows_('targets');
  for (var i = data.length - 1; i >= 0; i--) if (data[i][0] === nid) sh.deleteRow(i + 2);
  var roster = {};
  rows_('students').forEach(function (r) { roster[r[0]] = r[1]; });
  targets.forEach(function (sid) {
    sid = String(sid);
    if (roster[sid] !== undefined) sh.appendRow([nid, sid, roster[sid]]);
  });
}

function deleteNotice_(p) {
  var role = need_(p, ['admin', 'president']);
  var nid = String(p.nid || '');
  return lock_(function () {
    var sh = sh_('notices');
    var data = rows_('notices');
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === nid) {
        if (role === 'president' && data[i][8] !== '회장') throw new Error('본인이 쓴 공지만 삭제할 수 있어요.');
        sh.deleteRow(i + 2);
        replaceTargets_(nid, []);
        return { ok: true };
      }
    }
    throw new Error('삭제할 공지를 찾지 못했어요.');
  });
}

function saveStudents_(p) {
  need_(p, ['admin']);
  var list = JSON.parse(p.data || '[]');
  return lock_(function () {
    var sh = sh_('students');
    var last = sh.getLastRow();
    if (last >= 2) sh.getRange(2, 1, last - 1, 2).clearContent();
    if (list.length) {
      var vals = list.map(function (s) { return [String(s.sid), String(s.name)]; });
      sh.getRange(2, 1, vals.length, 2).setNumberFormat('@').setValues(vals);
    }
    return { ok: true, count: list.length };
  });
}

function saveSettings_(p) {
  need_(p, ['admin']);
  var s = JSON.parse(p.data || '{}');
  return lock_(function () {
    if (s.year !== undefined) setSetting_('학년도', s.year);
    if (s.className !== undefined) setSetting_('학급명', s.className);
    if (s.presidentCode !== undefined) setSetting_('회장코드', s.presidentCode);
    if (s.newPw) {
      if (String(s.newPw) === DEFAULT_PW) throw new Error('초기 비밀번호와 다른 비밀번호로 정해 주세요.');
      if (String(s.newPw).length < 4) throw new Error('비밀번호는 4자 이상으로 정해 주세요.');
      setSetting_('관리자비번', s.newPw);
    }
    return { ok: true };
  });
}
