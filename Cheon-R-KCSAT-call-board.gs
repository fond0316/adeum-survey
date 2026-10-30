/**
 * KCSAT 감독 호출 보드 GAS 백엔드 — Made by Cheon_R
 * 모든 요청은 JSONP(doGet + callback) 방식
 * 재배포 시 반드시 [배포 관리] → [수정] → 새 버전 (새 배포 금지)
 *
 * 시트 탭(자동 생성): 설정 / 설정목록 / 호출 / 전체공지
 * ※ 시트 탭은 통째로 삭제하지 마세요. 행만 지우거나 앱의 [기록 지우기]를 쓰세요.
 *   (실수로 지웠다면 아래 '헤더복구' 함수를 한 번 실행하세요)
 */

/** GAS 코드 버전 — 코드를 고칠 때마다 1씩 올려요. 화면이 이 번호로 업데이트가 필요한지 알려줘요. */
var GAS_VER = 5;

var TZ = 'Asia/Seoul';
var DEFAULT_PW = 'kcsatcallboard';   // 초기 비밀번호 — 첫 로그인 때 변경 강제

var SHEETS = {
  settings: { name: '설정', headers: ['항목', '값'] },
  lists: { name: '설정목록', headers: ['종류', '순서', '값'] },
  calls: { name: '호출', headers: ['호출ID', '그룹ID', '연습', '교시', '고사실', '역할', '이름', '지참물', '메모', '메모공개', '긴급', '상태', '알림번호', '재호출수', '호출시각', '확인시각', '처리시각', '취소시각', '호출ms', '최근호출ms', '확인ms', '처리ms', '문구'] },
  notices: { name: '전체공지', headers: ['공지ID', '내용', '긴급', '연습', '상태', '알림번호', '게시시각', '내린시각', '게시ms', '종료ms'] }
};
var C = { id: 0, group: 1, practice: 2, period: 3, room: 4, role: 5, name: 6, supplies: 7, memo: 8, memoPub: 9, urgent: 10, status: 11, alert: 12, recalls: 13, at: 14, confirmAt: 15, doneAt: 16, cancelAt: 17, ms: 18, lastMs: 19, confirmMs: 20, doneMs: 21, msg: 22 };
var N = { id: 0, text: 1, urgent: 2, practice: 3, state: 4, alert: 5, at: 6, offAt: 7, ms: 8, endMs: 9 };

var DEFAULT_SETTINGS = [['관리자비번', DEFAULT_PW], ['알림번호', '0'], ['메모기본공개', 'N'], ['지참물형식', '지참물: {items}'], ['고사실수', '30']];
var DEFAULT_LISTS = {
  period: ['1교시 국어', '2교시 수학', '3교시 영어', '4교시 한국사·탐구', '5교시 제2외국어·한문'],
  role: ['제1감독관', '제2감독관', '복도감독관'],
  supply: ['도장', '필기구', '감독관 명찰'],
  phrase: ['확인 사항이 있으니 본부로 와 주시기 바랍니다', '예비 시험지를 가지러 와 주시기 바랍니다', '물품 수령을 위해 본부로 와 주시기 바랍니다', '예비 감독관 투입을 위해 해당 고사실로 가 주시기 바랍니다', '감독 교대를 위해 본부로 와 주시기 바랍니다', '긴급 상황입니다. 즉시 본부로 와 주시기 바랍니다'],
  apphrase: ['잠시 대기해 주세요, 곧 안내드리겠습니다', '전체 감독관은 본부로 모여 주세요', '점심시간 안내: 본부 확인 후 이동해 주세요', '시험 종료 후 안내 사항이 있으니 대기해 주세요'],
  roster: []
};
var LIST_KINDS = ['period', 'role', 'supply', 'phrase', 'apphrase', 'roster'];
var STATUS = ['미처리', '확인됨', '처리완료', '취소'];

/* ---------- 진입점 ---------- */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var cb = String(p.callback || 'callback');
  if (!/^[\w.$]+$/.test(cb)) cb = 'callback';
  var out;
  try {
    ready_();
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
    case 'ping': return { ok: true, gasVer: GAS_VER };
    case 'poll': return poll_(p);
    case 'confirm': return confirm_(p);
    case 'login': return login_(p);
    case 'adminLoad': return adminLoad_(p);
    case 'call': return call_(p);
    case 'status': return status_(p);
    case 'announce': return announce_(p);
    case 'unannounce': return unannounce_(p);
    case 'saveLists': return saveLists_(p);
    case 'saveSetting': return saveSetting_(p);
    case 'changePw': return changePw_(p);
    case 'clear': return clear_(p);
    default: return { ok: false, error: '알 수 없는 요청이에요.' };
  }
}

/* ---------- 시트 준비 (탭이 없으면 헤더 포함 자동 생성) ---------- */
function ready_() {
  var cache = CacheService.getScriptCache();
  if (cache.get('ready')) return;
  setup_();
  cache.put('ready', '1', 600);
}
function setup_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(function (k) {
    var def = SHEETS[k];
    var sh = ss.getSheetByName(def.name);
    if (!sh) {
      sh = ss.insertSheet(def.name);
      sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold');
      sh.getRange(1, 1, 1000, def.headers.length).setNumberFormat('@');
      sh.setFrozenRows(1);
      if (k === 'settings') sh.getRange(2, 1, DEFAULT_SETTINGS.length, 2).setValues(DEFAULT_SETTINGS);
      if (k === 'lists') {
        var rows = [];
        LIST_KINDS.forEach(function (kind) { DEFAULT_LISTS[kind].forEach(function (v, i) { rows.push([kind, String(i + 1), v]); }); });
        if (rows.length) sh.getRange(2, 1, rows.length, 3).setValues(rows);
      }
    }
  });
}
/** 탭 헤더가 지워졌을 때 수동 실행용 */
function 헤더복구() {
  setup_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(function (k) {
    var def = SHEETS[k];
    var sh = ss.getSheetByName(def.name);
    if (sh) sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold');
  });
}

/* ---------- 공통 유틸 ---------- */
function sh_(k) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS[k].name); }
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function todayStr_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function isDate_(v) { return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime()); }
function str_(v) {
  if (isDate_(v)) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm:ss');
  return v === null || v === undefined ? '' : String(v).trim();
}
function clip_(v, n) { return str_(v).slice(0, n); }
function ensureRows_(sh, upTo) {
  var max = sh.getMaxRows();
  if (upTo > max) sh.insertRowsAfter(max, upTo - max + 100);
}
function writeRow_(sh, rowNum, row) {
  ensureRows_(sh, rowNum);
  sh.getRange(rowNum, 1, 1, row.length).setNumberFormat('@').setValues([row]);
}
function rows_(k) {
  var sh = sh_(k);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var n = SHEETS[k].headers.length;
  return sh.getRange(2, 1, last - 1, n).getValues().map(function (r) { return r.map(str_); });
}
/** 탭의 데이터 행을 통째로 바꿔 씀 */
function replaceRows_(k, rows) {
  var sh = sh_(k), n = SHEETS[k].headers.length;
  var last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, 1, last - 1, n).clearContent();
  if (rows.length) {
    ensureRows_(sh, rows.length + 1);
    sh.getRange(2, 1, rows.length, n).setNumberFormat('@').setValues(rows);
  }
}
function getSetting_(key) {
  var r = rows_('settings').filter(function (x) { return x[0] === key; })[0];
  return r ? r[1] : '';
}
function setSetting_(key, val) {
  var sh = sh_('settings');
  var data = rows_('settings');
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === key) { sh.getRange(i + 2, 2).setNumberFormat('@').setValue(String(val)); return; }
  }
  writeRow_(sh, sh.getLastRow() + 1, [key, String(val)]);
}
function installId_() {
  var v = getSetting_('설치ID');
  if (!v) { v = Utilities.getUuid(); setSetting_('설치ID', v); }
  return v;
}
function nextAlert_() {
  var n = (Number(getSetting_('알림번호')) || 0) + 1;
  setSetting_('알림번호', n);
  return n;
}
function auth_(pw) { pw = String(pw || ''); return !!pw && pw === getSetting_('관리자비번'); }
function need_(p) { if (!auth_(p.pw)) throw new Error('권한이 없어요. 비밀번호를 확인해 주세요.'); }
function json_(s, d) { try { return JSON.parse(s); } catch (e) { return d; } }

/* 쓰기 작업: 한 번에 하나씩 + 끝나면 대기실 화면용 기억(캐시)을 비움 */
function cbump_() {
  try {
    var pr = PropertiesService.getScriptProperties();
    pr.setProperty('gen', String((Number(pr.getProperty('gen')) || 0) + 1));
  } catch (e) {}
}
function lock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { cbump_(); lock.releaseLock(); }
}

/* ---------- 로그인 / 비밀번호 ---------- */
function login_(p) {
  if (!auth_(p.pw)) return { ok: true, role: '' };
  return { ok: true, role: 'admin', gasVer: GAS_VER, installId: installId_(), mustChange: String(p.pw) === DEFAULT_PW };
}
function changePw_(p) {
  need_(p);
  var np = String(p.newPw || '');
  if (np.length < 4) throw new Error('새 비밀번호는 4자 이상으로 해 주세요.');
  if (np === DEFAULT_PW) throw new Error('처음 비밀번호와 다른 비밀번호로 정해 주세요.');
  return lock_(function () { setSetting_('관리자비번', np); return { ok: true }; });
}

/* ---------- 목록 / 설정 ---------- */
function lists_() {
  var out = {};
  LIST_KINDS.forEach(function (k) { out[k] = []; });
  rows_('lists').sort(function (a, b) { return Number(a[1]) - Number(b[1]); })
    .forEach(function (r) { if (out[r[0]]) out[r[0]].push(r[2]); });
  return out;
}
function saveLists_(p) {
  need_(p);
  var kind = String(p.kind || '');
  if (LIST_KINDS.indexOf(kind) < 0) throw new Error('알 수 없는 목록이에요.');
  var items = json_(p.items, null);
  if (!Array.isArray(items)) throw new Error('목록 형식이 올바르지 않아요.');
  if (String(p.mode) === 'add') {   // 이름 명단처럼 많은 항목은 여러 번에 나눠 '추가'로 보내요 (기존 항목은 그대로 두고 뒤에 붙임)
    items = rows_('lists').filter(function (r) { return r[0] === kind; }).map(function (r) { return String(r[2]); }).concat(items);
  }
  var seen = {}, clean = [];
  items.forEach(function (v) {
    v = clip_(v, 60);
    if (v && !seen[v]) { seen[v] = 1; clean.push(v); }
  });
  if (clean.length > (kind === 'roster' ? 300 : 80)) throw new Error('항목은 최대 ' + (kind === 'roster' ? 300 : 80) + '개까지예요.');
  return lock_(function () {
    var keep = rows_('lists').filter(function (r) { return r[0] !== kind; });
    clean.forEach(function (v, i) { keep.push([kind, String(i + 1), v]); });
    replaceRows_('lists', keep);
    return { ok: true, items: clean };
  });
}
var SETTING_KEYS = ['메모기본공개', '지참물형식', '고사실수', '학교명', '인사말', '대기실소리'];
function saveSetting_(p) {
  need_(p);
  var key = String(p.key || ''), val = String(p.value === undefined ? '' : p.value);
  if (SETTING_KEYS.indexOf(key) < 0) throw new Error('알 수 없는 설정이에요.');
  if (key === '메모기본공개') val = val === 'Y' ? 'Y' : 'N';
  if (key === '대기실소리') val = val === 'N' ? 'N' : 'Y';
  if (key === '지참물형식') {
    val = val.trim().slice(0, 60);
    if (val.indexOf('{items}') < 0) throw new Error('문구에 {items} 가 들어 있어야 지참물이 채워져요.');
  }
  if (key === '학교명') val = val.replace(/[\r\n\t]+/g, ' ').trim().slice(0, 30);
  if (key === '인사말') val = val.replace(/[\r\n\t]+/g, ' ').trim().slice(0, 120);
  if (key === '고사실수') {
    var n = Math.round(Number(val));
    if (!(n >= 1 && n <= 99)) throw new Error('고사실 수는 1~99 사이로 해 주세요.');
    val = String(n);
  }
  return lock_(function () { setSetting_(key, val); return { ok: true }; });
}
function settingsOut_() {
  return {
    memoDefault: getSetting_('메모기본공개') === 'Y',
    suppliesFormat: getSetting_('지참물형식') || '지참물: {items}',
    roomMax: Number(getSetting_('고사실수')) || 30,
    school: getSetting_('학교명'),
    greet: getSetting_('인사말'),
    roomSound: getSetting_('대기실소리') !== 'N'
  };
}

/* ---------- 호출 ---------- */
function callObj_(r) {
  return {
    id: r[C.id], group: r[C.group], practice: r[C.practice] === 'Y', period: r[C.period], room: r[C.room], role: r[C.role],
    name: r[C.name], msg: r[C.msg] || '본부로 와 주시기 바랍니다', supplies: r[C.supplies], memo: r[C.memo], memoPublic: r[C.memoPub] === 'Y', urgent: r[C.urgent] === 'Y',
    status: r[C.status], alert: Number(r[C.alert]) || 0, recalls: Number(r[C.recalls]) || 0,
    at: r[C.at], confirmAt: r[C.confirmAt], doneAt: r[C.doneAt], cancelAt: r[C.cancelAt],
    ms: Number(r[C.ms]) || 0, since: Number(r[C.lastMs]) || Number(r[C.ms]) || 0, confirmMs: Number(r[C.confirmMs]) || 0, doneMs: Number(r[C.doneMs]) || 0
  };
}
function call_(p) {
  need_(p);
  var list = json_(p.rows, null);
  if (!Array.isArray(list) || !list.length) throw new Error('호출할 대상이 없어요.');
  if (list.length > 40) throw new Error('한 번에 40명까지 부를 수 있어요.');
  var practice = p.practice === '1' ? 'Y' : 'N';
  var supplies = '';
  var items = json_(p.supplies, []);
  if (Array.isArray(items) && items.length) {
    var clean = items.map(function (v) { return clip_(v, 30); }).filter(function (v) { return v; }).slice(0, 10);
    if (clean.length) supplies = (getSetting_('지참물형식') || '지참물: {items}').replace('{items}', clean.join(', '));
  }
  var msg = clip_(p.msg, 80) || '본부로 와 주시기 바랍니다';
  var memo = clip_(p.memo, 200), memoPub = p.memoPublic === '1' ? 'Y' : 'N', urgent = p.urgent === '1' ? 'Y' : 'N';
  return lock_(function () {
    var sh = sh_('calls');
    var exist = {};
    rows_('calls').forEach(function (r) { exist[r[C.id]] = true; });
    var group = clip_(p.group, 40) || ('g' + Date.now() + Math.floor(Math.random() * 1000));
    var ms = Date.now(), at = now_();
    var add = [];
    list.forEach(function (x, i) {
      var id = clip_(x.id, 40);
      if (!id || exist[id]) return;     // 같은 호출을 두 번 눌러도 한 번만 저장
      var name = clip_(x.name, 20), room = clip_(x.room, 20);
      if (!name && !room) return;
      add.push([id, group, practice, clip_(x.period, 40), room, clip_(x.role, 30), name, supplies, memo, memoPub, urgent, '미처리', '', '0', at, '', '', '', String(ms), String(ms), '', '', msg]);
    });
    if (!add.length) return { ok: true, added: 0, group: group };
    var a = nextAlert_();
    add.forEach(function (r) { r[C.alert] = String(a); });
    var start = sh.getLastRow() + 1;
    ensureRows_(sh, start + add.length);
    sh.getRange(start, 1, add.length, SHEETS.calls.headers.length).setNumberFormat('@').setValues(add);
    return { ok: true, added: add.length, group: group, alert: a };
  });
}
/** 대기실 선생님이 [확인했어요] — 누르지 않아도 되는 보조 버튼 */
function confirm_(p) {
  var id = String(p.id || '');
  if (!id) throw new Error('호출 정보가 없어요.');
  return lock_(function () {
    var sh = sh_('calls'), data = rows_('calls');
    for (var i = 0; i < data.length; i++) {
      if (data[i][C.id] === id) {
        if (data[i][C.status] === '미처리') {
          data[i][C.status] = '확인됨'; data[i][C.confirmAt] = now_(); data[i][C.confirmMs] = String(Date.now());
          writeRow_(sh, i + 2, data[i]);
        }
        return { ok: true };
      }
    }
    return { ok: true };
  });
}
/** 본부: 처리완료 / 취소 / 되돌리기 / 다시 부르기 */
function status_(p) {
  need_(p);
  var ids = json_(p.ids, null);
  if (!Array.isArray(ids) || !ids.length) throw new Error('대상이 없어요.');
  var to = String(p.to || '');
  if (['done', 'cancel', 'undo', 'recall'].indexOf(to) < 0) throw new Error('알 수 없는 처리예요.');
  var want = {};
  ids.forEach(function (x) { want[String(x)] = true; });
  return lock_(function () {
    var sh = sh_('calls'), data = rows_('calls'), n = 0, a = 0;
    if (to === 'recall') a = nextAlert_();
    for (var i = 0; i < data.length; i++) {
      var r = data[i];
      if (!want[r[C.id]]) continue;
      var st = r[C.status];
      if (to === 'done') {
        if (st !== '미처리' && st !== '확인됨') continue;
        r[C.status] = '처리완료'; r[C.doneAt] = now_(); r[C.doneMs] = String(Date.now());
      } else if (to === 'cancel') {
        if (st !== '미처리' && st !== '확인됨') continue;
        r[C.status] = '취소'; r[C.cancelAt] = now_();
      } else if (to === 'undo') {
        if (st !== '처리완료' && st !== '취소') continue;
        r[C.status] = r[C.confirmMs] ? '확인됨' : '미처리'; r[C.doneAt] = ''; r[C.doneMs] = ''; r[C.cancelAt] = '';
      } else {   // recall
        r[C.status] = '미처리'; r[C.confirmAt] = ''; r[C.confirmMs] = ''; r[C.doneAt] = ''; r[C.doneMs] = ''; r[C.cancelAt] = '';
        r[C.alert] = String(a); r[C.recalls] = String((Number(r[C.recalls]) || 0) + 1); r[C.lastMs] = String(Date.now());
      }
      writeRow_(sh, i + 2, r); n++;
    }
    return { ok: true, changed: n };
  });
}

/* ---------- 전체 공지 ---------- */
function announce_(p) {
  need_(p);
  var text = clip_(p.text, 300);
  if (!text) throw new Error('공지 내용을 입력해 주세요.');
  var minutes = Math.max(0, Math.min(600, Math.round(Number(p.minutes) || 0)));
  return lock_(function () {
    var sh = sh_('notices');
    var id = clip_(p.id, 40) || ('n' + Date.now() + Math.floor(Math.random() * 1000));
    var dup = rows_('notices').filter(function (r) { return r[N.id] === id; })[0];
    if (dup) return { ok: true, id: id };
    var ms = Date.now();
    var row = [id, text, p.urgent === '1' ? 'Y' : 'N', p.practice === '1' ? 'Y' : 'N', '게시중', String(nextAlert_()), now_(), '', String(ms), minutes ? String(ms + minutes * 60000) : ''];
    writeRow_(sh, sh.getLastRow() + 1, row);
    return { ok: true, id: id };
  });
}
function unannounce_(p) {
  need_(p);
  var id = String(p.id || '');
  return lock_(function () {
    var sh = sh_('notices'), data = rows_('notices');
    for (var i = 0; i < data.length; i++) {
      if (data[i][N.id] === id && data[i][N.state] === '게시중') {
        data[i][N.state] = '내림'; data[i][N.offAt] = now_();
        writeRow_(sh, i + 2, data[i]);
      }
    }
    return { ok: true };
  });
}
function noticeObj_(r) {
  return { id: r[N.id], text: r[N.text], urgent: r[N.urgent] === 'Y', practice: r[N.practice] === 'Y', state: r[N.state], alert: Number(r[N.alert]) || 0, at: r[N.at], offAt: r[N.offAt], ms: Number(r[N.ms]) || 0, endMs: Number(r[N.endMs]) || 0 };
}
function activeNotice_(o, nowMs) { return o.state === '게시중' && (!o.endMs || o.endMs > nowMs); }

/* ---------- 대기실 화면용 (비밀번호 없이 읽기만) ---------- */
/* ---------- 대기실 생존 신호 (대기실 화면이 5초마다 poll 하면서 함께 남김) ---------- */
function hbTouch_(p) {
  var rid = String(p.rid || '').replace(/[^\w-]/g, '').slice(0, 24);
  if (!rid) return;
  try {
    var cache = CacheService.getScriptCache();
    cache.put('hb_' + rid, JSON.stringify({ t: Date.now(), seen: Number(p.seen) || 0, au: String(p.au) === '1' ? 1 : 0 }), 21600);
    var ids = (json_(cache.get('hbids'), []) || []);
    if (ids.indexOf(rid) < 0) { ids.push(rid); if (ids.length > 10) ids = ids.slice(-10); cache.put('hbids', JSON.stringify(ids), 21600); }
  } catch (e) {}
}
/** 대기실 상태: exists(한 번이라도 연결됐나) / ageMs(가장 최근 신호가 몇 ms 전) / seen(살아 있는 대기실이 화면에 그린 가장 큰 알림번호) */
function hbRead_() {
  var out = { exists: false, ageMs: 0, seen: 0, audio: false };
  try {
    var cache = CacheService.getScriptCache();
    var ids = (json_(cache.get('hbids'), []) || []);
    if (!ids.length) return out;
    var all = cache.getAll(ids.map(function (k) { return 'hb_' + k; })), now = Date.now(), best = null;
    Object.keys(all).forEach(function (k) {
      var v = json_(all[k], null); if (!v) return;
      var age = now - Number(v.t);
      if (best === null || age < best) best = age;
      if (age <= 20000 && Number(v.seen) > out.seen) out.seen = Number(v.seen);
      if (age <= 20000 && Number(v.au) === 1) out.audio = true;   // 살아 있는 대기실 중 한 대라도 소리 준비(화면을 한 번 눌렀음)
    });
    if (best !== null) { out.exists = true; out.ageMs = best; }
  } catch (e) {}
  return out;
}
function poll_(p) {
  hbTouch_(p);
  var cache = CacheService.getScriptCache(), gen = '0';
  try { gen = PropertiesService.getScriptProperties().getProperty('gen') || '0'; } catch (e) {}
  var key = 'poll_' + gen, hit = cache.get(key), out;
  if (hit) {
    out = JSON.parse(hit);
  } else {
    out = {
      ok: true, gasVer: GAS_VER,
      calls: rows_('calls').filter(function (r) { return r[C.status] === '미처리' || r[C.status] === '확인됨'; }).map(function (r) {
        var o = callObj_(r);
        return { id: o.id, group: o.group, practice: o.practice, period: o.period, room: o.room, role: o.role, name: o.name, msg: o.msg, supplies: o.supplies,
          memo: o.memoPublic ? o.memo : '', urgent: o.urgent, status: o.status, alert: o.alert, since: o.since, confirmMs: o.confirmMs };
      }),
      notices: rows_('notices').map(noticeObj_).filter(function (o) { return o.state === '게시중'; }),
      greet: { school: getSetting_('학교명'), text: getSetting_('인사말') },
      sound: getSetting_('대기실소리') !== 'N'   // 대기실 화면 인사말 (학교 이름 · 문구)
    };
    try { cache.put(key, JSON.stringify(out), 10); } catch (e) {}
  }
  var t = Date.now();
  out.now = t;
  out.notices = out.notices.filter(function (o) { return activeNotice_(o, t); });   // 시간이 지난 공지는 자동으로 빠짐
  return out;
}

/* ---------- 본부 화면용 ---------- */
function adminLoad_(p) {
  need_(p);
  var t = Date.now();
  return {
    ok: true, gasVer: GAS_VER, now: t,
    calls: rows_('calls').map(callObj_),
    notices: rows_('notices').map(noticeObj_).map(function (o) { o.active = activeNotice_(o, t); return o; }),
    lists: lists_(), settings: settingsOut_(), room: hbRead_()
  };
}

/* ---------- 기록 지우기 ---------- */
function clear_(p) {
  need_(p);
  var scope = String(p.scope || '');
  if (['today', 'all', 'practice'].indexOf(scope) < 0) throw new Error('알 수 없는 범위예요.');
  return lock_(function () {
    var today = todayStr_(), removed = 0;
    var keepCalls = rows_('calls').filter(function (r) {
      var drop = scope === 'all' || (scope === 'practice' && r[C.practice] === 'Y') || (scope === 'today' && String(r[C.at]).slice(0, 10) === today);
      if (drop) removed++;
      return !drop;
    });
    var keepNotices = rows_('notices').filter(function (r) {
      var drop = scope === 'all' || (scope === 'practice' && r[N.practice] === 'Y') || (scope === 'today' && String(r[N.at]).slice(0, 10) === today);
      if (drop) removed++;
      return !drop;
    });
    replaceRows_('calls', keepCalls);
    replaceRows_('notices', keepNotices);
    return { ok: true, removed: removed };
  });
}
