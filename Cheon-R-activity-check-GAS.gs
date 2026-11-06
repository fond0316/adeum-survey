/**
 * ACTIVITY CHECK (학생 활동 자체점검) GAS 백엔드 — Made by Cheon_R
 * 모든 요청은 JSONP(doGet + callback) 방식
 * 재배포 시 반드시 [배포 관리] → [수정] → 새 버전 (새 배포 금지)
 */

/** GAS 코드 버전 — 코드를 고칠 때마다 1씩 올려요. 관리 화면이 이 번호로 업데이트가 필요한지 알려줘요. */
var GAS_VER = 7;

var TZ = 'Asia/Seoul';
var DEFAULT_PW = 'activitycheck';   // 초기 비밀번호 — 첫 로그인 때 변경 강제
var CATS = ['자율', '진로', '교과', '동아리'];
var SCATS = ['자율', '진로', '모르겠어요'];   // 학생이 직접 추가하는 활동의 분류 ('모르겠어요'는 담임이 정해 줘요)
var MAX_SADD = 10;                  // 학생 1명이 추가할 수 있는 활동 수
var ATYPES = ['독서', '탐구보고서', '실험', '토론', '발표', '프로젝트', '기타'];   // 교과 활동 유형
var RTYPES = ['해성에듀', '학급자율활동앱', '직접링크', '없음'];
var MAX_FAIL = 5;                   // 학생 PIN을 이만큼 틀리면 10분 동안 잠깐 막아요
var MAX_ADMIN_FAIL = 8;             // 담임 비밀번호를 이만큼 틀리면 10분 동안 잠깐 막아요
var C_FRESH = 60;                   // 명단·항목·설정은 60초 동안 기억해 두고 써요 (접속이 몰려도 시트를 덜 읽게)

var SHEETS = {
  settings: { name: '설정', headers: ['항목', '값'] },
  students: { name: '학생명단', headers: ['학번', '이름', 'PIN', '진로희망', 'PIN설정시각', '동아리'] },
  items: { name: '항목', headers: ['ID', '분류', '항목명', '작성유의사항', '보고서종류', '기본링크', '담임제공', '순서', '사용', '교사용기재유의', '문구제공', '학기', '보낼곳', '확인필요', '마감일', '추가한학생'] },
  records: { name: '기록', headers: ['학번', '항목ID', '번호', '참여', '보고서', '링크', '주제', '진로연결', '요약', '수정시각', '교사확인', '교사메모', '활동유형', '학기', '자료'] }
};
var DEFAULT_SETTINGS = [['학년도', String(new Date().getFullYear())], ['학급명', '우리 반'], ['관리자비번', DEFAULT_PW]];

/* ---------- 진입점 ---------- */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var cb = String(p.callback || 'callback');
  if (!/^[A-Za-z_$][\w$.]{0,60}$/.test(cb)) cb = 'callback';   // 콜백 이름은 글자·숫자만
  var out;
  try {
    setup_();
    out = route_(p);
  } catch (err) {
    var m = String(err && err.message || err);
    out = m === 'BUSY' ? { ok: false, busy: true, error: '지금 접속이 많아서 잠깐 밀렸어요. 자동으로 다시 시도해요.' } : { ok: false, error: m };
  }
  return ContentService
    .createTextOutput(cb + '(' + JSON.stringify(out) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function route_(p) {
  switch (p.action) {
    case 'ping': return { ok: true, className: getSetting_('학급명'), year: year_(), gasVer: GAS_VER };
    case 'sCheck': return sCheck_(p);
    case 'sSetPin': return sSetPin_(p);
    case 'sLoad': return sLoad_(p);
    case 'sSave': return sSave_(p);
    case 'sDelete': return sDelete_(p);
    case 'sCareer': return sCareer_(p);
    case 'sClub': return sClub_(p);
    case 'sAddItem': return sAddItem_(p);
    case 'sDelItem': return sDelItem_(p);
    case 'promoteSItem': return promoteSItem_(p);
    case 'mergeSItem': return mergeSItem_(p);
    case 'delSItem': return delSItem_(p);
    case 'login': return login_(p);
    case 'adminLoad': return adminLoad_(p);
    case 'saveItem': return saveItem_(p);
    case 'saveItems': return saveItems_(p);
    case 'saveGuide': return saveGuide_(p);
    case 'deleteItem': return deleteItem_(p);
    case 'deleteItems': return deleteItems_(p);
    case 'saveImpLog': return saveImpLog_(p);
    case 'reorderItems': return reorderItems_(p);
    case 'saveStudents': return saveStudents_(p);
    case 'resetPin': return resetPin_(p);
    case 'saveSettings': return saveSettings_(p);
    case 'tReview': return tReview_(p);
    case 'yearReset': return yearReset_(p);
    default: return { ok: false, error: '알 수 없는 요청이에요.' };
  }
}

/* ---------- 시트 준비 (탭이 없으면 헤더 포함 자동 생성) ---------- */
function setup_() {
  var cache = CacheService.getScriptCache(), key = 'setup:' + GAS_VER;
  if (cache.get(key)) return;
  setupSheets_();
  cache.put(key, '1', 21600);
}
function setupSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(function (k) {
    var def = SHEETS[k];
    var sh = ss.getSheetByName(def.name);
    if (!sh) {
      sh = ss.insertSheet(def.name);
      sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold');
      sh.getRange(1, 1, 1000, def.headers.length).setNumberFormat('@'); // 학번·날짜 자동변환 방지
      sh.setFrozenRows(1);
      if (k === 'settings') sh.getRange(2, 1, DEFAULT_SETTINGS.length, 2).setValues(DEFAULT_SETTINGS);
    } else {
      // 예전 버전 시트: 새로 생긴 칸의 머리글만 뒤에 채워 넣어요 (기존 데이터는 그대로)
      var cur = sh.getRange(1, 1, 1, def.headers.length).getValues()[0], first = -1;
      for (var i = 0; i < def.headers.length; i++) if (!cur[i]) { cur[i] = def.headers[i]; if (first < 0) first = i; }
      if (first >= 0) {
        sh.getRange(1, 1, 1, def.headers.length).setValues([cur]).setFontWeight('bold');
        sh.getRange(1, first + 1, 1000, def.headers.length - first).setNumberFormat('@');
      }
    }
  });
}
/** 탭 헤더가 지워졌을 때 수동 실행용 */
function 헤더복구() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(function (k) {
    var def = SHEETS[k], sh = ss.getSheetByName(def.name);
    if (sh) sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold');
  });
  setupSheets_();
}

/* ---------- 공통 유틸 ---------- */
function sh_(k) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS[k].name); }
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function isDate_(v) { return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime()); }
function str_(v) {
  if (isDate_(v)) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm:ss');
  return v === null || v === undefined ? '' : String(v).trim();
}
function ensureRows_(sh, upTo) {
  var max = sh.getMaxRows();
  if (upTo > max) sh.insertRowsAfter(max, upTo - max + 100);
}
/** 행을 글자 형식으로 고정해서 쓰기 (자동 변환 방지) */
function writeRow_(sh, rowNum, row) {
  ensureRows_(sh, rowNum);
  sh.getRange(rowNum, 1, 1, row.length).setNumberFormat('@').setValues([row]);
}
function rows_(k) {
  var sh = sh_(k), last = sh.getLastRow();
  if (last < 2) return [];
  var n = SHEETS[k].headers.length;
  return sh.getRange(2, 1, last - 1, n).getValues().map(function (r) { return r.map(str_); });
}
/** 잠깐 기억해 두기 (접속이 몰릴 때 시트를 덜 읽어요) — 값이 바뀌는 곳에서는 cdel_로 지워요 */
function cached_(key, loader, force) {
  var c = CacheService.getScriptCache(), o = null;
  if (!force) {
    try { var t = c.get(key); o = t ? JSON.parse(t) : null; } catch (e) { o = null; }
  }
  if (o) {
    if (Date.now() - o.t < C_FRESH * 1000) return o.v;
    if (c.get('rf_' + key)) return o.v;          // 누군가 새로 읽는 중이면 직전 값을 받아요
    c.put('rf_' + key, '1', 15);
  }
  var v = loader();
  try { c.put(key, JSON.stringify({ t: Date.now(), v: v }), 21600); } catch (e) {}
  return v;
}
function cdel_(key) { CacheService.getScriptCache().remove(key); }
function settingsTable_(fresh) { return cached_('c_set', function () { return rows_('settings'); }, fresh); }
function getSetting_(key) {
  var r = settingsTable_().filter(function (x) { return x[0] === key; })[0];
  return r ? r[1] : '';
}
function setSetting_(key, val) {
  var sh = sh_('settings'), data = rows_('settings');
  cdel_('c_set');
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === key) { sh.getRange(i + 2, 2).setValue(String(val)); return; }
  }
  sh.appendRow([key, String(val)]);
}
function installId_() {
  var v = getSetting_('설치ID');
  if (!v) { v = Utilities.getUuid(); setSetting_('설치ID', v); }
  return v;
}
function year_() { return getSetting_('학년도') || Utilities.formatDate(new Date(), TZ, 'yyyy'); }
function lock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(9000)) throw new Error('BUSY');   // 너무 몰리면 화면이 알아서 다시 시도해요
  try { return fn(); } finally { lock.releaseLock(); }
}
/** 길이 제한 (칸이 모두 글자 형식이라 =, - 로 시작해도 수식으로 실행되지 않아요) */
function cut_(v, n) {
  return String(v === null || v === undefined ? '' : v).replace(/\r\n?/g, '\n').trim().substring(0, n);
}
/** 기록 시트에서 (학번, 항목ID, 번호) 줄 위치만 빠르게 찾기 — 앞 3칸만 읽어요 */
function recIdx_(sh, sid, iid, n) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var d = sh.getRange(2, 1, last - 1, 3).getValues();
  for (var i = 0; i < d.length; i++) if (str_(d[i][0]) === sid && str_(d[i][1]) === iid && str_(d[i][2]) === String(n)) return i;
  return -1;
}
function okUrl_(u) { return /^https?:\/\/\S+$/i.test(u) && u.length <= 300; }
function hex_(bytes) {
  return bytes.map(function (b) { b = (b < 0 ? b + 256 : b); return ('0' + b.toString(16)).slice(-2); }).join('');
}
function hashPin_(sid, pin) {
  return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, installId_() + ':' + sid + ':' + pin));
}

/* ---------- 담임 로그인 ---------- */
function auth_(pw) {
  pw = String(pw || '');
  if (!pw) return '';
  var cache = CacheService.getScriptCache(), bad = Number(cache.get('bad_admin')) || 0;
  if (bad >= MAX_ADMIN_FAIL) throw new Error('비밀번호를 여러 번 틀려서 10분 동안 잠시 막았어요. 10분 뒤에 다시 해 보세요.');
  if (pw === getSetting_('관리자비번')) { if (bad) cache.remove('bad_admin'); return 'admin'; }
  cache.put('bad_admin', String(bad + 1), 600);
  return '';
}
function need_(p) {
  if (auth_(p.pw) !== 'admin') throw new Error('권한이 없어요. 비밀번호를 확인해 주세요.');
}
function login_(p) {
  var role = auth_(p.pw);
  if (!role) return { ok: true, role: '' };
  return { ok: true, role: role, gasVer: GAS_VER, installId: installId_(), className: getSetting_('학급명'),
    mustChange: String(p.pw) === DEFAULT_PW };
}

/* ---------- 항목 · 기록 변환 ---------- */
function itemObj_(r) {
  return { id: r[0], cat: r[1], title: r[2], note: r[3], rtype: r[4] || '없음', url: r[5], prov: r[6] === 'Y', order: Number(r[7]) || 0, on: r[8] !== 'N',
    tnote: r[9] || '', give: r[10] === 'Y', term: r[11] || '', dest: r[12] || '', chk: r[13] === 'Y', due: r[14] || '' };
}
/** 학생에게 보내는 항목 — 교사용 기재 유의·보낼곳·확인필요는 절대 넣지 않아요 */
function itemObjS_(r) {
  return { id: r[0], cat: r[1], title: r[2], note: r[3], rtype: r[4] || '없음', url: r[5], prov: r[6] === 'Y', order: Number(r[7]) || 0, on: true,
    give: r[10] === 'Y', term: r[11] || '', due: r[14] || '' };
}
function itemsS_(fresh) {
  return cached_('c_items_s', function () {
    return sortItems_(rows_('items').filter(function (r) { return r[0] && r[8] !== 'N' && !r[15]; }).map(itemObjS_));
  }, fresh);
}
/** 학생이 직접 추가한 활동 (추가한학생 칸에 학번이 있는 줄) — 담임이 정식 항목으로 올리기 전까지 그 학생에게만 보여요 */
function ownedS_(fresh) {
  return cached_('c_items_u', function () {
    return rows_('items').filter(function (r) { return r[0] && r[15]; }).map(function (r) { return { id: r[0], cat: r[1], title: r[2], by: r[15] }; });
  }, fresh);
}
function sortItems_(list) {
  return list.sort(function (a, b) { return (CATS.indexOf(a.cat) - CATS.indexOf(b.cat)) || (a.order - b.order); });
}
function recObj_(r) {
  return { sid: r[0], iid: r[1], n: Number(r[2]) || 1, part: r[3], report: r[4], link: r[5], topic: r[6], career: r[7], summary: r[8], at: r[9], tchk: r[10] === 'Y', tmemo: r[11], atype: r[12] || '', rterm: r[13] || '', ref: r[14] || '' };
}
function findRec_(data, sid, iid, n) {
  for (var i = 0; i < data.length; i++) if (data[i][0] === sid && data[i][1] === iid && data[i][2] === String(n)) return i;
  return -1;
}

/* ---------- 학생용 ---------- */
function normName_(s) { return String(s || '').replace(/\s+/g, ''); }
function stuTable_(fresh) { return cached_('c_stu', function () { return rows_('students'); }, fresh); }
function stuRow_(sid, fresh) {
  sid = String(sid || '').trim();
  var data = stuTable_(fresh);
  for (var i = 0; i < data.length; i++) if (data[i][0] === sid) return { row: i + 2, v: data[i] };
  return null;
}
/** 학번+이름 확인 (PIN을 처음 만들 수 있는지도 알려줘요) */
function sCheck_(p) {
  var st = stuRow_(p.sid);
  if (!st || normName_(st.v[1]) !== normName_(p.name) || !st.v[2]) st = stuRow_(p.sid, true);   // 기억해 둔 값이 오래됐을 수 있어 한 번 더 확인
  if (!st || normName_(st.v[1]) !== normName_(p.name)) return { ok: false, error: '학번과 이름이 맞지 않아요. 담임선생님께 확인해 주세요.' };
  return { ok: true, name: st.v[1], needPin: !st.v[2] };
}
function sSetPin_(p) {
  var st = stuRow_(p.sid, true), pin = String(p.pin || '');
  if (!st || normName_(st.v[1]) !== normName_(p.name)) return { ok: false, error: '학번과 이름이 맞지 않아요.' };
  if (!/^\d{4}$/.test(pin)) return { ok: false, error: 'PIN은 숫자 4자리로 만들어 주세요.' };
  return lock_(function () {
    var cur = stuRow_(p.sid, true);
    if (cur.v[2]) return { ok: false, error: '이미 PIN이 만들어져 있어요. 기억나지 않으면 담임선생님께 PIN 초기화를 부탁해 주세요.' };
    var sh = sh_('students');
    sh.getRange(cur.row, 3).setNumberFormat('@').setValue(hashPin_(cur.v[0], pin));
    sh.getRange(cur.row, 5).setNumberFormat('@').setValue(now_());
    cdel_('c_stu');
    return { ok: true, name: cur.v[1] };
  });
}
/** PIN 확인 — 여러 번 틀리면 10분 동안 막아요 */
function stuAuth_(sid, pin) {
  sid = String(sid || '').trim(); pin = String(pin || '');
  var cache = CacheService.getScriptCache(), ck = 'bad_' + sid;
  var bad = Number(cache.get(ck)) || 0;
  if (bad >= MAX_FAIL) throw new Error('PIN을 여러 번 틀려서 잠시 막았어요. 10분 뒤에 다시 해 보거나 담임선생님께 PIN 초기화를 부탁해 주세요.');
  var st = stuRow_(sid) || stuRow_(sid, true);   // 못 찾으면 시트를 한 번 더 확인
  if (!st) throw new Error('명단에 없는 학번이에요.');
  if (!st.v[2]) { st = stuRow_(sid, true); if (!st || !st.v[2]) { var e = new Error('PIN을 먼저 만들어 주세요.'); throw e; } }
  if (st.v[2] !== hashPin_(sid, pin)) {
    var fr = stuRow_(sid, true);
    if (!fr || fr.v[2] !== hashPin_(sid, pin)) {
      cache.put(ck, String(bad + 1), 600);
      throw new Error('PIN이 맞지 않아요.' + (bad + 1 >= MAX_FAIL ? ' 잠시 뒤에 다시 해 주세요.' : ''));
    }
    st = fr;
  }
  if (bad) cache.remove(ck);
  return st;
}
function sLoad_(p) {
  var st = stuRow_(p.sid);
  if (st && !st.v[2]) st = stuRow_(p.sid, true);
  if (st && !st.v[2] && normName_(st.v[1]) === normName_(p.name || st.v[1])) return { ok: true, needPin: true, name: st.v[1] };
  st = stuAuth_(p.sid, p.pin);
  var sid = st.v[0];
  var items = itemsS_();
  var recs = rows_('records').filter(function (r) { return r[0] === sid; }).map(recObj_);
  return { ok: true, name: st.v[1], career: st.v[3], club: st.v[5] || '', className: getSetting_('학급명'), year: year_(), items: items, mine: ownedS_().filter(function (o) { return o.by === sid; }), records: recs, gasVer: GAS_VER };
}
function sSave_(p) {
  var st = stuAuth_(p.sid, p.pin), sid = st.v[0];
  var d = JSON.parse(p.data || '{}');
  var iid0 = String(d.itemId);
  var byId = function (x) { return x.id === iid0; };
  var item = itemsS_().filter(byId)[0] || itemsS_(true).filter(byId)[0];
  if (!item) {   // 내가 직접 추가한 활동이면 그것으로 저장해요
    var ow = ownedS_().filter(byId)[0] || ownedS_(true).filter(byId)[0];
    if (ow && ow.by === sid) item = { id: ow.id, cat: ow.cat };
  }
  if (!item) throw new Error('없어진 항목이에요. 화면을 새로고침해 주세요.');
  if (item.cat === '동아리' && st.v[5] !== item.id) throw new Error('내 동아리로 고른 항목만 기록할 수 있어요.');
  var n = (item.cat === '교과' || item.cat === '동아리') ? Math.max(1, Math.min(9, parseInt(d.n, 10) || 1)) : 1;
  var part = (d.part === 'yes' || d.part === 'no') ? d.part : '';
  var report = (d.report === 'done' || d.report === 'todo') ? d.report : '';
  var link = String(d.link || '').trim();
  if (link.length > 300) throw new Error('링크·경로는 300자 이내로 적어 주세요.');   // 주소(https://…)나 경로 글(예: 해성에듀 > 내 보고서) 모두 가능
  var topic = cut_(d.topic, 60), career = cut_(d.career, 80), summary = cut_(d.summary, 500);
  var atype = ATYPES.indexOf(d.atype) >= 0 ? d.atype : '', rterm = (d.rterm === '1학기' || d.rterm === '2학기') ? d.rterm : '', ref = cut_(d.ref, 200);
  return lock_(function () {
    var sh = sh_('records'), idx = recIdx_(sh, sid, item.id, n);
    var memo = idx >= 0 ? str_(sh.getRange(idx + 2, 12).getValue()) : '';
    var at = now_();
    writeRow_(sh, idx >= 0 ? idx + 2 : sh.getLastRow() + 1, [sid, item.id, String(n), part, report, link, topic, career, summary, at, '', memo, atype, rterm, ref]);   // 학생이 고치면 교사확인은 풀려요
    return { ok: true, rec: { sid: sid, iid: item.id, n: n, part: part, report: report, link: link, topic: topic, career: career, summary: summary, at: at, tchk: false, tmemo: memo, atype: atype, rterm: rterm, ref: ref } };
  });
}
/** 학생이 놓친 활동을 직접 추가 (id가 있으면 분류만 바꿔요) */
function sAddItem_(p) {
  var st = stuAuth_(p.sid, p.pin), sid = st.v[0];
  var d = JSON.parse(p.data || '{}');
  var cat = String(d.cat || '');
  if (SCATS.indexOf(cat) < 0) throw new Error('분류는 자율·진로·모르겠어요 중에서 골라 주세요.');
  var id = String(d.id || '');
  return lock_(function () {
    var sh = sh_('items'), data = rows_('items'), idx = -1;
    if (id) {
      for (var i = 0; i < data.length; i++) if (data[i][0] === id) { idx = i; break; }
      if (idx < 0 || data[idx][15] !== sid) throw new Error('내가 추가한 활동만 고칠 수 있어요.');
      sh.getRange(idx + 2, 2).setNumberFormat('@').setValue(cat);
      cdel_('c_items_u');
      return { ok: true, item: { id: id, cat: cat, title: data[idx][2], by: sid } };
    }
    var title = cut_(d.title, 60);
    if (!title) throw new Error('활동 이름을 적어 주세요.');
    var key = normName_(title).toLowerCase(), cnt = 0;
    data.forEach(function (r) {
      if (!r[0]) return;
      if (r[15] === sid) { cnt++; if (normName_(r[2]).toLowerCase() === key) throw new Error('이미 내가 추가한 활동이에요.'); }
      else if (!r[15] && r[8] !== 'N' && r[1] !== '동아리' && normName_(r[2]).toLowerCase() === key) throw new Error('이미 목록에 있는 활동이에요. 자율·진로 목록에서 찾아보세요.');
    });
    if (cnt >= MAX_SADD) throw new Error('직접 추가는 ' + MAX_SADD + '개까지 할 수 있어요. 선생님께 말씀해 주세요.');
    var nid = 'U' + Date.now().toString(36) + Math.floor(Math.random() * 1296).toString(36);
    writeRow_(sh, sh.getLastRow() + 1, [nid, cat, title, '', '직접링크', '', '', '0', '', '', '', '', '', '', '', sid]);
    cdel_('c_items_u');
    return { ok: true, item: { id: nid, cat: cat, title: title, by: sid } };
  });
}
/** 내가 추가한 활동 지우기 (적어 둔 기록도 함께 지워져요) */
function sDelItem_(p) {
  var st = stuAuth_(p.sid, p.pin), sid = st.v[0], id = String(p.id || '');
  return lock_(function () {
    var sh = sh_('items'), data = rows_('items'), found = false;
    for (var i = 0; i < data.length; i++) if (data[i][0] === id && data[i][15] === sid) { sh.deleteRow(i + 2); found = true; break; }
    if (!found) throw new Error('내가 추가한 활동만 지울 수 있어요.');
    delRecsOf_(id);
    cdel_('c_items_u');
    return { ok: true };
  });
}
/** 어떤 항목의 학생 기록을 모두 지워요 (위에서부터 지우면 줄 번호가 밀려서 아래부터) */
function delRecsOf_(iid) {
  var sh = sh_('records'), last = sh.getLastRow();
  if (last < 2) return 0;
  var d = sh.getRange(2, 2, last - 1, 1).getValues(), n = 0;
  for (var i = d.length - 1; i >= 0; i--) if (str_(d[i][0]) === iid) { sh.deleteRow(i + 2); n++; }
  return n;
}
function sDelete_(p) {
  var st = stuAuth_(p.sid, p.pin), sid = st.v[0], iid = String(p.iid || ''), n = parseInt(p.n, 10) || 1;
  return lock_(function () {
    var sh = sh_('records'), idx = recIdx_(sh, sid, iid, n);
    if (idx >= 0) sh.deleteRow(idx + 2);
    return { ok: true };
  });
}
/** 내 동아리 고르기 (학생당 1개 · 바꿔도 돼요 · 비우려면 빈 값) */
function sClub_(p) {
  var st = stuAuth_(p.sid, p.pin), cid = String(p.club || '');
  if (cid) {
    var byId = function (x) { return x.id === cid && x.cat === '동아리'; };
    if (!itemsS_().filter(byId)[0] && !itemsS_(true).filter(byId)[0]) throw new Error('없는 동아리예요. 화면을 새로고침해 주세요.');
  }
  return lock_(function () {
    sh_('students').getRange(stuRow_(st.v[0], true).row, 6).setNumberFormat('@').setValue(cid);
    cdel_('c_stu');
    return { ok: true, club: cid };
  });
}
function sCareer_(p) {
  var st = stuAuth_(p.sid, p.pin), career = cut_(p.career, 80);
  return lock_(function () {
    sh_('students').getRange(stuRow_(st.v[0], true).row, 4).setNumberFormat('@').setValue(career);
    cdel_('c_stu');
    return { ok: true, career: career };
  });
}

/* ---------- 담임용 ---------- */
function adminLoad_(p) {
  need_(p);
  var allIt = rows_('items').filter(function (r) { return r[0]; });
  var items = sortItems_(allIt.filter(function (r) { return !r[15]; }).map(itemObj_));
  var nm = {}; rows_('students').forEach(function (r) { nm[r[0]] = r[1]; });
  var sItems = allIt.filter(function (r) { return r[15]; }).map(function (r) { return { id: r[0], cat: r[1], title: r[2], by: r[15], byName: nm[r[15]] || '' }; });
  var students = rows_('students').filter(function (r) { return r[0]; }).map(function (r) { return { sid: r[0], name: r[1], hasPin: !!r[2], career: r[3], club: r[5] || '' }; });
  var records = rows_('records').filter(function (r) { return r[0]; }).map(recObj_);
  return { ok: true, role: 'admin', gasVer: GAS_VER, className: getSetting_('학급명'), year: year_(), grade: getSetting_('학년'), guide: guide_(), impLog: impLog_(),
    items: items, sItems: sItems, students: students, records: records, installId: installId_() };
}
/** 불러온 파일 기록 (엑셀·붙여넣기로 등록한 묶음) — 설정 탭에 JSON으로 보관 */
function impLog_() {
  try { var a = JSON.parse(getSetting_('불러온파일') || '[]'); return Array.isArray(a) ? a : []; } catch (e) { return []; }
}
function saveImpLog_(p) {
  need_(p);
  var a = JSON.parse(p.data || '[]');
  if (!Array.isArray(a)) throw new Error('기록 형식이 올바르지 않아요.');
  var KINDS = ['items', 'club', 'tt', 'paste'];
  var out = a.slice(-20).map(function (x) {
    return { id: cut_(x.id, 20), name: cut_(x.name, 80), kind: KINDS.indexOf(x.kind) >= 0 ? x.kind : 'items', count: Number(x.count) || 0, at: cut_(x.at, 20),
      ids: (Array.isArray(x.ids) ? x.ids : []).slice(0, 80).map(function (i) { return cut_(i, 30); }) };
  });
  var s = JSON.stringify(out);
  while (s.length > 40000 && out.length > 1) { out.shift(); s = JSON.stringify(out); }
  return lock_(function () { setSetting_('불러온파일', s); return { ok: true }; });
}
var GUIDE_KEYS = ['공통', '자율', '진로', '행발'];
function guide_() {
  var g = {};
  GUIDE_KEYS.forEach(function (k) { g[k] = getSetting_('기재지침_' + k); });
  return g;
}
/** 영역별 기재 지침 (교사만 봐요) */
function saveGuide_(p) {
  need_(p);
  var k = String(p.key || ''), text = cut_(p.text, 500);
  if (GUIDE_KEYS.indexOf(k) < 0) throw new Error('지침 종류가 올바르지 않아요.');
  return lock_(function () { setSetting_('기재지침_' + k, text); return { ok: true }; });
}
function cleanItem_(it) {
  var cat = String(it.cat || '');
  if (CATS.indexOf(cat) < 0) throw new Error('분류는 자율·진로·교과·동아리 중에서 골라 주세요.');
  var title = cut_(it.title, 60);
  if (!title) throw new Error('항목명이 비어 있어요.');
  var rtype = RTYPES.indexOf(it.rtype) >= 0 ? it.rtype : '없음';
  var url = String(it.url || '').trim();
  if (url && !okUrl_(url)) throw new Error('"' + title + '"의 링크는 https://로 시작하는 300자 이내 주소여야 해요.');
  var term = (it.term === '1학기' || it.term === '2학기') ? it.term : '';
  var due = /^\d{4}-\d{2}-\d{2}$/.test(String(it.due || '')) ? String(it.due) : '';   // 마감일(선택)
  var dest = (it.dest === '자율' || it.dest === '진로' || it.dest === '행발') ? it.dest : '';
  return { cat: cat, title: title, note: cut_(it.note, 300), rtype: rtype, url: url, prov: it.prov ? 'Y' : '', on: it.on === false ? 'N' : '',
    tnote: cut_(it.tnote, 300), give: it.give ? 'Y' : '', term: term, dest: dest, chk: it.chk ? 'Y' : '', due: due };
}
function maxOrder_(data) {
  var m = 0;
  data.forEach(function (r) { m = Math.max(m, Number(r[7]) || 0); });
  return m;
}
function saveItem_(p) {
  need_(p);
  var it = JSON.parse(p.data || '{}'), c = cleanItem_(it);
  return lock_(function () {
    var sh = sh_('items'), data = rows_('items'), idx = -1, id = String(it.id || '');
    if (id) for (var i = 0; i < data.length; i++) if (data[i][0] === id) { idx = i; break; }
    if (id && idx < 0) throw new Error('수정할 항목을 찾지 못했어요.');
    if (idx < 0) id = 'I' + Date.now().toString(36) + Math.floor(Math.random() * 1296).toString(36);
    var order = idx >= 0 ? (Number(data[idx][7]) || 0) : maxOrder_(data) + 1;
    writeRow_(sh, idx >= 0 ? idx + 2 : sh.getLastRow() + 1, [id, c.cat, c.title, c.note, c.rtype, c.url, c.prov, String(order), c.on, c.tnote, c.give, c.term, c.dest, c.chk, c.due]);
    cdel_('c_items_s');
    return { ok: true, id: id };
  });
}
/** 엑셀·붙여넣기로 한꺼번에 — mode: append(추가) / replace(처음 묶음에서 기존 항목을 모두 지움) / replaceCat(한 분류만 지움) */
function saveItems_(p) {
  need_(p);
  var d = JSON.parse(p.data || '{}'), list = d.items || [];
  if (!Array.isArray(list) || list.length > 20) throw new Error('한 번에 20개까지만 올릴 수 있어요.');
  var cleaned = list.map(cleanItem_);
  return lock_(function () {
    var sh = sh_('items');
    if (d.mode === 'replace') {
      var allr = rows_('items');   // 학생이 직접 추가한 활동(추가한학생 칸이 있는 줄)은 남겨 둬요
      for (var q0 = allr.length - 1; q0 >= 0; q0--) if (!allr[q0][15]) sh.deleteRow(q0 + 2);
    } else if (d.mode === 'replaceCat') {   // 한 분류(예: 동아리)의 항목만 지우고 새로 넣기
      var cc = String(d.cat || '');
      if (CATS.indexOf(cc) < 0) throw new Error('분류가 올바르지 않아요.');
      var all = rows_('items');
      for (var q = all.length - 1; q >= 0; q--) if (all[q][1] === cc && !all[q][15]) sh.deleteRow(q + 2);
    }
    var data = rows_('items'), order = maxOrder_(data), rows = [], ids = [];
    cleaned.forEach(function (c, i) {
      var want = String((list[i] && list[i].id) || ''), idx = -1;
      if (want) for (var k = 0; k < data.length; k++) if (data[k][0] === want) { idx = k; break; }
      if (idx >= 0) {   // 이미 있는 항목은 새 내용으로 고쳐요 (ID·순서·보임 여부 그대로 → 학생 기록이 이어져요)
        writeRow_(sh, idx + 2, [want, c.cat, c.title, c.note, c.rtype, c.url, c.prov, String(Number(data[idx][7]) || 0), data[idx][8] === 'N' ? 'N' : '', c.tnote, c.give, c.term, c.dest, c.chk, c.due]);
        ids.push(want); return;
      }
      var nid = 'I' + Date.now().toString(36) + i + Math.floor(Math.random() * 1296).toString(36);
      ids.push(nid);
      rows.push([nid, c.cat, c.title, c.note, c.rtype, c.url, c.prov, String(++order), c.on, c.tnote, c.give, c.term, c.dest, c.chk, c.due]);
    });
    if (rows.length) {
      ensureRows_(sh, sh.getLastRow() + rows.length);
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, 15).setNumberFormat('@').setValues(rows);
    }
    cdel_('c_items_s');
    return { ok: true, count: ids.length, ids: ids };
  });
}
/** 항목 여러 개를 한 번에 지우기 (불러온 파일 삭제용) */
function deleteItems_(p) {
  need_(p);
  var ids = JSON.parse(p.ids || '[]');
  if (!Array.isArray(ids) || ids.length > 200) throw new Error('한 번에 200개까지만 지울 수 있어요.');
  var set = {}; ids.forEach(function (x) { set[String(x)] = 1; });
  return lock_(function () {
    var sh = sh_('items'), data = rows_('items'), n = 0;
    for (var i = data.length - 1; i >= 0; i--) if (set[data[i][0]]) { sh.deleteRow(i + 2); n++; }
    cdel_('c_items_s');
    return { ok: true, count: n };
  });
}
function deleteItem_(p) {
  need_(p);
  var id = String(p.id || '');
  return lock_(function () {
    var data = rows_('items');
    for (var i = 0; i < data.length; i++) if (data[i][0] === id) { sh_('items').deleteRow(i + 2); cdel_('c_items_s'); return { ok: true }; }
    throw new Error('삭제할 항목을 찾지 못했어요.');
  });
}
/** 학생이 추가한 활동을 정식 항목으로 올리기 — ID가 그대로라 학생이 적어 둔 기록이 이어져요 */
function promoteSItem_(p) {
  need_(p);
  var id = String(p.id || ''), cat = String(p.cat || '');
  if (cat !== '자율' && cat !== '진로' && cat !== '교과') throw new Error('분류는 자율·진로·교과 중에서 골라 주세요.');
  return lock_(function () {
    var sh = sh_('items'), data = rows_('items'), idx = -1;
    for (var i = 0; i < data.length; i++) if (data[i][0] === id && data[i][15]) { idx = i; break; }
    if (idx < 0) throw new Error('학생이 추가한 활동을 찾지 못했어요. 화면을 새로고침해 주세요.');
    var title = cut_(p.title, 60) || data[idx][2];
    writeRow_(sh, idx + 2, [id, cat, title, '', '없음', '', '', String(maxOrder_(data) + 1), '', '', '', '', '', '', '', '']);
    cdel_('c_items_s'); cdel_('c_items_u');
    return { ok: true, id: id };
  });
}
/** 학생이 추가한 활동의 기록을 같은 이름의 정식 항목으로 옮기고 학생 항목은 지우기 */
function mergeSItem_(p) {
  need_(p);
  var id = String(p.id || ''), target = String(p.target || '');
  return lock_(function () {
    var ish = sh_('items'), items = rows_('items'), si = -1, tg = null;
    for (var i = 0; i < items.length; i++) {
      if (items[i][0] === id && items[i][15]) si = i;
      if (items[i][0] === target && !items[i][15]) tg = items[i];
    }
    if (si < 0) throw new Error('학생이 추가한 활동을 찾지 못했어요. 화면을 새로고침해 주세요.');
    if (!tg) throw new Error('합칠 정식 항목을 찾지 못했어요.');
    if (tg[1] === '동아리') throw new Error('동아리 항목에는 합칠 수 없어요.');
    var rsh = sh_('records'), recs = rows_('records'), moved = 0, skipped = 0, multi = tg[1] === '교과';
    for (var k = 0; k < recs.length; k++) {
      if (recs[k][1] !== id) continue;
      var sid = recs[k][0], n = 1;
      if (multi) {
        var used = {}; recs.forEach(function (r) { if (r[0] === sid && r[1] === target) used[Number(r[2]) || 1] = 1; });
        n = 0; for (var q = 1; q <= 9; q++) if (!used[q]) { n = q; break; }
      } else if (findRec_(recs, sid, target, 1) >= 0) n = 0;
      if (!n) { skipped++; continue; }
      rsh.getRange(k + 2, 2, 1, 2).setNumberFormat('@').setValues([[target, String(n)]]);
      recs[k][1] = target; recs[k][2] = String(n); moved++;
    }
    var left = recs.filter(function (r) { return r[1] === id; }).length;
    if (!left) ish.deleteRow(si + 2);   // 못 옮긴 기록이 남아 있으면 학생 항목을 남겨 둬요 (기록이 사라지지 않게)
    cdel_('c_items_u');
    return { ok: true, moved: moved, skipped: skipped, removed: !left };
  });
}
/** 학생이 추가한 활동 지우기 (그 학생의 기록도 함께) */
function delSItem_(p) {
  need_(p);
  var id = String(p.id || '');
  return lock_(function () {
    var sh = sh_('items'), data = rows_('items');
    for (var i = 0; i < data.length; i++) if (data[i][0] === id && data[i][15]) {
      sh.deleteRow(i + 2); var n = delRecsOf_(id); cdel_('c_items_u');
      return { ok: true, records: n };
    }
    throw new Error('학생이 추가한 활동을 찾지 못했어요.');
  });
}
function reorderItems_(p) {
  need_(p);
  var list = JSON.parse(p.data || '[]');
  if (!Array.isArray(list) || list.length > 100) throw new Error('순서 정보가 올바르지 않아요.');
  return lock_(function () {
    var sh = sh_('items'), data = rows_('items');
    list.forEach(function (o) {
      for (var i = 0; i < data.length; i++) if (data[i][0] === String(o.id)) { sh.getRange(i + 2, 8).setNumberFormat('@').setValue(String(parseInt(o.order, 10) || 0)); break; }
    });
    cdel_('c_items_s');
    return { ok: true };
  });
}
/** 명단 저장 — 이미 있는 학생의 PIN·진로희망은 그대로 지켜요 */
function saveStudents_(p) {
  need_(p);
  var list = JSON.parse(p.data || '[]');
  if (!Array.isArray(list) || list.length > 60) throw new Error('명단은 60명까지 저장할 수 있어요.');
  return lock_(function () {
    var sh = sh_('students'), old = {};
    rows_('students').forEach(function (r) { old[r[0]] = r; });
    var vals = list.map(function (s) {
      var sid = cut_(s.sid, 10), o = old[sid] || [];
      return [sid, cut_(s.name, 20), o[2] || '', o[3] || '', o[4] || '', o[5] || ''];
    });
    var last = sh.getLastRow();
    if (last >= 2) sh.getRange(2, 1, last - 1, 6).clearContent();
    if (vals.length) { ensureRows_(sh, vals.length + 1); sh.getRange(2, 1, vals.length, 6).setNumberFormat('@').setValues(vals); }
    cdel_('c_stu');
    return { ok: true, count: vals.length };
  });
}
function resetPin_(p) {
  need_(p);
  var sid = String(p.sid || '').trim();
  return lock_(function () {
    var st = stuRow_(sid, true);
    if (!st) throw new Error('학생을 찾지 못했어요.');
    sh_('students').getRange(st.row, 3, 1, 3).setNumberFormat('@').setValues([['', st.v[3], '']]);
    cdel_('c_stu');
    CacheService.getScriptCache().remove('bad_' + sid);
    return { ok: true };
  });
}
function saveSettings_(p) {
  need_(p);
  var s = JSON.parse(p.data || '{}');
  return lock_(function () {
    if (s.year !== undefined) setSetting_('학년도', cut_(s.year, 4));
    if (s.className !== undefined) setSetting_('학급명', cut_(s.className, 30));
    if (s.grade !== undefined) setSetting_('학년', /^[123]$/.test(String(s.grade)) ? String(s.grade) : '');
    if (s.newPw) {
      if (String(s.newPw) === DEFAULT_PW) throw new Error('초기 비밀번호와 다른 비밀번호로 정해 주세요.');
      if (String(s.newPw).length < 4) throw new Error('비밀번호는 4자 이상으로 정해 주세요.');
      setSetting_('관리자비번', s.newPw);
    }
    return { ok: true };
  });
}
/** 담임 점검: 확인 표시 + 한마디 */
function tReview_(p) {
  need_(p);
  var sid = String(p.sid || '').trim(), iid = String(p.iid || ''), n = String(parseInt(p.n, 10) || 1);
  var chk = p.chk === '1' ? 'Y' : '', memo = cut_(p.memo, 200);
  if (!stuRow_(sid)) throw new Error('학생을 찾지 못했어요.');
  return lock_(function () {
    var sh = sh_('records'), data = rows_('records'), idx = findRec_(data, sid, iid, n);
    if (idx < 0) writeRow_(sh, sh.getLastRow() + 1, [sid, iid, n, '', '', '', '', '', '', now_(), chk, memo]);
    else sh.getRange(idx + 2, 11, 1, 2).setNumberFormat('@').setValues([[chk, memo]]);
    return { ok: true };
  });
}
/** 학년도 초기화: 학생 기록만 지워요 (항목·명단은 그대로, 지우기 전에 백업 탭을 만들어요) */
function yearReset_(p) {
  need_(p);
  return lock_(function () {
    var sh = sh_('records'), last = sh.getLastRow(), bk = '';
    if (last >= 2) {
      // 지우기 전에 '기록백업_날짜' 탭으로 통째로 복사해 둬요 (되돌릴 때 쓰세요)
      bk = '기록백업_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmm');
      sh.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName(bk);
      sh.deleteRows(2, last - 1);
    }
    var ssh = sh_('students'), sl = ssh.getLastRow();
    if (sl >= 2) ssh.getRange(2, 6, sl - 1, 1).clearContent();   // 새 학년도엔 동아리를 다시 고르게 해요
    var ish = sh_('items'), idata = rows_('items');   // 학생이 직접 추가했던 활동도 정리해요 (기록과 함께 백업 탭·백업 파일에 남아 있어요)
    for (var qi = idata.length - 1; qi >= 0; qi--) if (idata[qi][15]) ish.deleteRow(qi + 2);
    cdel_('c_items_u'); cdel_('c_items_s');
    cdel_('c_stu');
    return { ok: true, count: Math.max(0, last - 1), backup: bk };
  });
}
