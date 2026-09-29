/**
 * 학생공지 GAS 백엔드 — Made by Cheon_R
 * 모든 요청은 JSONP(doGet + callback) 방식
 * 재배포 시 반드시 [배포 관리] → [수정] → 새 버전 (새 배포 금지)
 */

var TZ = 'Asia/Seoul';
var SHEETS = {
  settings: { name: '설정', headers: ['항목', '값'] },
  students: { name: '학생명단', headers: ['학번', '이름'] },
  notices: { name: '공지', headers: ['ID', '유형', '날짜', '제목', '내용', '대상메모', '중요', '팝업', '작성자', '작성시각', '수정시각', '종료일'] },
  targets: { name: '호출대상', headers: ['공지ID', '학번', '이름'] },
  confirms: { name: '확인기록', headers: ['공지ID', '학번', '이름', '확인시각'] },
  requests: { name: '수정요청', headers: ['ID', '공지ID', '학번', '이름', '내용', '작성시각', '처리', '처리시각'] },
  calldone: { name: '호출완료', headers: ['공지ID', '학번', '이름', '시각'] },
  files: { name: '첨부', headers: ['공지ID', '순서', '종류', '이름', '주소'] },
  links: { name: '바로가기', headers: ['이름', '주소'] },
  photos: { name: '사진', headers: ['사진ID', '순서', '조각'] },
  photoIdx: { name: '사진목록', headers: ['사진ID', '조각수', '글자수', '저장시각'] }
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
    case 'search': return search_(p);
    case 'who': return who_(p);
    case 'confirm': return confirm_(p);
    case 'request': return request_(p);
    case 'callDone': return callDone_(p);
    case 'resolveRequest': return resolveRequest_(p);
    case 'login': return login_(p);
    case 'adminLoad': return adminLoad_(p);
    case 'saveNotice': return saveNotice_(p);
    case 'saveNotices': return saveNotices_(p);
    case 'deleteNotice': return deleteNotice_(p);
    case 'saveStudents': return saveStudents_(p);
    case 'saveSettings': return saveSettings_(p);
    case 'saveLinks': return saveLinks_(p);
    case 'photo': return photo_(p);
    case 'photoStatus': return photoStatus_(p);
    case 'photoDrop': return photoDrop_(p);
    case 'photoCleanup': return photoCleanup_(p);
    default: return { ok: false, error: '알 수 없는 요청이에요.' };
  }
}

/* ---------- 시트 준비 (탭이 없으면 헤더 포함 자동 생성) ---------- */
/** 사진 올리기 — 용량이 커서 GET(주소)으로는 못 보내고 POST로 받아요 */
function doPost(e) {
  var p = {}, out;
  try {
    setup_();
    p = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (p.action === 'uploadPhoto') out = uploadPhoto_(p);
    else out = { ok: false, error: '알 수 없는 요청이에요.' };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
    try { if (okPhotoId_(String(p.id || ''))) CacheService.getScriptCache().put('photoerr_' + p.id, out.error, 600); } catch (e2) {}
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
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
      if (k === 'photos') { try { sh.hideSheet(); } catch (e) {} }   // 긴 글자로 가득한 탭이라 숨겨 둠
    } else if (k === 'notices') {
      var hc = sh.getRange(1, 12);   // 예전에 만든 시트: '종료일' 열 머리글 자동 추가
      if (String(hc.getValue()) !== '종료일') { hc.setValue('종료일').setFontWeight('bold'); sh.getRange(2, 12, Math.max(1, sh.getMaxRows() - 1), 1).setNumberFormat('@'); }
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
function isDate_(v) { return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime()); }
function str_(v) {
  if (isDate_(v)) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return v === null || v === undefined ? '' : String(v).trim();
}
/** 날짜를 항상 yyyy-MM-dd 로 정리 (시트가 날짜 형식으로 바꿔도 대응) */
function dateStr_(v) {
  if (isDate_(v)) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  var s = String(v === null || v === undefined ? '' : v).trim();
  var m = s.match(/^(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  var d = new Date(s);
  return isNaN(d.getTime()) ? s : Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}
/** 행을 글자 형식으로 고정해서 쓰기 (자동 날짜 변환 방지) */
/** 시트 행이 모자라면 늘림 (기본 1000행) */
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
function noticeObj_(r, targetMap, attMap) {
  return {
    id: r[0], type: r[1], date: dateStr_(r[2]), title: r[3], body: r[4], memo: r[5],
    important: r[6] === 'Y', popup: r[7] === 'Y', author: r[8], created: r[9], updated: r[10],
    end: r[11] ? dateStr_(r[11]) : '',
    targets: targetMap[r[0]] || [], attachments: (attMap && attMap[r[0]]) || []
  };
}
/** 첨부(사진·링크) — { 공지ID: [{kind, name, url}] } */
function attMap_() {
  var m = {};
  rows_('files').filter(function (r) { return r[0]; })
    .sort(function (a, b) { return Number(a[1]) - Number(b[1]); })
    .forEach(function (r) { (m[r[0]] = m[r[0]] || []).push({ kind: r[2] === 'img' ? 'img' : (r[2] === 'photo' ? 'photo' : 'link'), name: r[3], url: r[4] }); });
  return m;
}
var MAX_ATT = 5, MAX_URL = 500, MAX_PHOTOS = 3;
function okUrl_(u) { return /^https?:\/\/\S+$/i.test(u) && u.length <= MAX_URL; }
function cleanAtts_(list) {
  if (!Array.isArray(list)) throw new Error('첨부 형식이 올바르지 않아요.');
  if (list.length > MAX_ATT) throw new Error('첨부는 최대 ' + MAX_ATT + '개까지예요.');
  var photos = 0;
  var out = list.map(function (a) {
    var url = String(a && a.url || '').trim();
    var kind = a && a.kind === 'photo' ? 'photo' : (a && a.kind === 'img' ? 'img' : 'link');
    if (kind === 'photo') {
      if (!/^photo:P[a-z0-9]{12,24}$/.test(url)) throw new Error('사진 정보가 올바르지 않아요.');
      if (!photoMeta_(url.slice(6))) throw new Error('사진 업로드가 아직 끝나지 않았어요. 잠시 뒤에 다시 올려 주세요.');
      photos++;
    } else if (!okUrl_(url)) throw new Error('첨부 주소는 http:// 또는 https://로 시작하는 ' + MAX_URL + '자 이내 주소여야 해요.');
    return { kind: kind, name: String(a.name || '').trim().slice(0, 40), url: url };
  });
  if (photos > MAX_PHOTOS) throw new Error('사진은 공지당 ' + MAX_PHOTOS + '장까지예요.');
  return out;
}
function replaceAttachments_(nid, list) {
  var sh = sh_('files'), data = rows_('files'), keep = {}, drop = [];
  list.forEach(function (a) { if (a.kind === 'photo') keep[a.url.slice(6)] = true; });
  for (var i = data.length - 1; i >= 0; i--) if (data[i][0] === nid) {
    if (data[i][2] === 'photo' && !keep[String(data[i][4]).slice(6)]) drop.push(String(data[i][4]).slice(6));
    sh.deleteRow(i + 2);
  }
  if (list.length) {
    var rows = list.map(function (a, i) { return [nid, i + 1, a.kind, a.name, a.url]; });
    ensureRows_(sh, sh.getLastRow() + rows.length);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, 5).setNumberFormat('@').setValues(rows);
  }
  drop.forEach(deletePhoto_);   // 공지에서 빠진 사진은 시트에서도 지움
}

/* ---------- 사진 (폰에서 올린 사진을 시트에 글자로 저장) ---------- */
var PHOTO_CHUNK = 40000, PHOTO_MAX_B64 = 850000, PHOTO_MAX_COUNT = 400;
function okPhotoId_(id) { return /^P[a-z0-9]{12,24}$/.test(id); }
function photoMeta_(id) {
  var r = rows_('photoIdx');
  for (var i = 0; i < r.length; i++) if (r[i][0] === id) return { row: i + 2, chunks: Number(r[i][1]), len: Number(r[i][2]), at: r[i][3] };
  return null;
}
/** '사진' 탭에서 이 사진의 조각이 시작하는 줄과 개수 (조각은 항상 연속) */
function photoSpan_(id) {
  var sh = sh_('photos'), last = sh.getLastRow();
  if (last < 2) return null;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues(), st = -1, n = 0;
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === id) { if (st < 0) st = i + 2; n++; } else if (st >= 0) break;
  }
  return st < 0 ? null : { start: st, count: n };
}
function deletePhoto_(id) {
  var sp = photoSpan_(id);
  if (sp) sh_('photos').deleteRows(sp.start, sp.count);
  var m = photoMeta_(id);
  if (m) sh_('photoIdx').deleteRow(m.row);
}
function uploadPhoto_(p) {
  need_(p, ['admin', 'president']);
  var id = String(p.id || ''), data = String(p.data || '');
  if (!okPhotoId_(id)) throw new Error('사진 번호가 올바르지 않아요.');
  if (!data || data.length > PHOTO_MAX_B64) throw new Error('사진이 너무 커요. 더 작게 줄여서 올려 주세요.');
  if (data.indexOf('/9j/') !== 0 || !/^[A-Za-z0-9+\/=]+$/.test(data)) throw new Error('JPG 사진만 올릴 수 있어요.');
  return lock_(function () {
    if (photoMeta_(id)) throw new Error('이미 저장된 사진 번호예요.');
    if (rows_('photoIdx').filter(function (r) { return r[0]; }).length >= PHOTO_MAX_COUNT) throw new Error('저장된 사진이 너무 많아요. ⚙️ 탭에서 사진을 정리해 주세요.');
    var rows = [];
    for (var i = 0; i < data.length; i += PHOTO_CHUNK) rows.push([id, rows.length + 1, data.substr(i, PHOTO_CHUNK)]);
    var sh = sh_('photos');
    ensureRows_(sh, sh.getLastRow() + rows.length);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, 3).setNumberFormat('@').setValues(rows);
    sh_('photoIdx').appendRow([id, rows.length, data.length, now_()]);   // 마지막에 기록 = 저장 완료 표시
    return { ok: true };
  });
}
function photoStatus_(p) {
  var id = String(p.id || '');
  if (!okPhotoId_(id)) return { ok: false, error: '사진 번호가 올바르지 않아요.' };
  if (photoMeta_(id)) return { ok: true, ready: true };
  var err = '';
  try { err = CacheService.getScriptCache().get('photoerr_' + id) || ''; } catch (e) {}
  return { ok: true, ready: false, error: err };
}
function photo_(p) {
  var id = String(p.id || ''), m = okPhotoId_(id) ? photoMeta_(id) : null, sp = m ? photoSpan_(id) : null;
  if (!m || !sp) return { ok: false, error: '사진을 찾지 못했어요.' };
  var v = sh_('photos').getRange(sp.start, 3, sp.count, 1).getValues();
  return { ok: true, data: 'data:image/jpeg;base64,' + v.map(function (r) { return r[0]; }).join('') };
}
/** 올렸다가 공지에 안 쓰게 된 사진 한 장 지우기 (공지에 연결된 사진은 지우지 않음) */
function photoDrop_(p) {
  need_(p, ['admin', 'president']);
  var id = String(p.id || '');
  if (!okPhotoId_(id)) return { ok: false, error: '사진 번호가 올바르지 않아요.' };
  return lock_(function () {
    var used = rows_('files').some(function (r) { return r[2] === 'photo' && String(r[4]) === 'photo:' + id; });
    if (!used) deletePhoto_(id);
    return { ok: true };
  });
}
function photoStats_() {
  var r = rows_('photoIdx').filter(function (x) { return x[0]; }), len = 0;
  r.forEach(function (x) { len += Number(x[2]) || 0; });
  return { count: r.length, kb: Math.round(len * 0.75 / 1024) };
}
/** 90일 지난 공지의 사진 + 어디에도 안 쓰는 사진 정리 (최근 24시간 안에 올린 사진은 작성 중일 수 있어 건드리지 않음) */
function photoCleanup_(p) {
  need_(p, ['admin']);
  return lock_(function () {
    var old = 0, orphan = 0, cutoff = daysAgo_(90);
    var recent = Utilities.formatDate(new Date(Date.now() - 86400000), TZ, 'yyyy-MM-dd HH:mm:ss');
    var nd = {};
    rows_('notices').forEach(function (r) { if (r[0]) nd[r[0]] = lastDay_(r); });
    var fsh = sh_('files'), fd = rows_('files');
    for (var i = fd.length - 1; i >= 0; i--) {
      if (fd[i][2] === 'photo' && nd[fd[i][0]] !== undefined && nd[fd[i][0]] < cutoff) {
        deletePhoto_(String(fd[i][4]).slice(6)); fsh.deleteRow(i + 2); old++;
      }
    }
    var used = {};
    rows_('files').forEach(function (r) { if (r[2] === 'photo') used[String(r[4]).slice(6)] = true; });
    var metaAt = {};
    rows_('photoIdx').forEach(function (r) { if (r[0]) metaAt[r[0]] = r[3]; });
    Object.keys(metaAt).forEach(function (id) {
      if (!used[id] && String(metaAt[id]) < recent) { deletePhoto_(id); orphan++; }
    });
    var sh = sh_('photos'), last = sh.getLastRow();      // 저장이 중간에 끊겨 조각만 남은 사진
    if (last >= 2) {
      var seen = {};
      sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) { var id = String(r[0]); if (id && !metaAt[id]) seen[id] = true; });
      Object.keys(seen).forEach(function (id) { deletePhoto_(id); orphan++; });
    }
    return { ok: true, old: old, orphan: orphan };
  });
}
/** 우리 반 앱 바로가기 */
function linksList_() {
  return rows_('links').filter(function (r) { return r[0] && r[1]; }).map(function (r) { return { name: r[0], url: r[1] }; });
}
function saveLinks_(p) {
  need_(p, ['admin']);
  var list = JSON.parse(p.data || '[]');
  if (!Array.isArray(list) || list.length > 10) throw new Error('바로가기는 최대 10개까지예요.');
  var rows = list.map(function (a) {
    var name = String(a && a.name || '').trim().slice(0, 20), url = String(a && a.url || '').trim();
    if (!name) throw new Error('이름이 빠진 바로가기가 있어요.');
    if (!okUrl_(url)) throw new Error('"' + name + '" 주소는 http:// 또는 https://로 시작해야 해요.');
    return [name, url];
  });
  return lock_(function () {
    var sh = sh_('links'), last = sh.getLastRow();
    if (last >= 2) sh.getRange(2, 1, last - 1, 2).clearContent();
    if (rows.length) sh.getRange(2, 1, rows.length, 2).setNumberFormat('@').setValues(rows);
    return { ok: true, count: rows.length };
  });
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
/** 공지가 끝나는 날 (기간 공지는 종료일, 아니면 날짜) */
function lastDay_(r) { var s = dateStr_(r[2]), e = r[11] ? dateStr_(r[11]) : ''; return e > s ? e : s; }
var MAX_PERIOD = 90;
function checkEnd_(n) {
  var e = String(n.end || '').trim();
  if (!e) return '';
  if (n.type === '호출') throw new Error('호출은 기간으로 올릴 수 없어요.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(e) || !/^\d{4}-\d{2}-\d{2}$/.test(String(n.date))) throw new Error('날짜 형식이 올바르지 않아요.');
  if (e < n.date) throw new Error('종료일이 시작일보다 빨라요.');
  if (e === n.date) return '';   // 같은 날이면 하루짜리
  var days = Math.round((new Date(e + 'T00:00:00') - new Date(n.date + 'T00:00:00')) / 86400000);
  if (days > MAX_PERIOD) throw new Error('기간은 최대 ' + MAX_PERIOD + '일까지예요.');
  return e;
}
/** 🔍 검색 — 전체 기간, 제목·내용, 띄어쓰기 무시, 여러 단어는 모두 포함 (호출 제외, 최신순 50개) */
function search_(p) {
  var q = String(p.q || '').trim().slice(0, 40);
  var norm = function (s) { return String(s || '').toLowerCase().replace(/\s+/g, ''); };
  var words = q.split(/\s+/).map(norm).filter(function (w) { return w; });
  if (!words.length) return { ok: true, q: q, results: [], more: false };
  var tm = targetMap_(), am = attMap_();
  var hits = rows_('notices').filter(function (r) {
    if (!r[0] || r[1] === '호출') return false;
    var text = norm(r[3] + ' ' + r[4]);
    return words.every(function (w) { return text.indexOf(w) >= 0; });
  }).sort(function (a, b) { return dateStr_(b[2]) < dateStr_(a[2]) ? -1 : dateStr_(b[2]) > dateStr_(a[2]) ? 1 : 0; });
  return { ok: true, q: q, results: hits.slice(0, 50).map(function (r) { return noticeObj_(r, tm, am); }), more: hits.length > 50 };
}
function load_(p) {
  var sid = String(p.sid || '');
  var from = daysAgo_(60);
  var tm = targetMap_(), am = attMap_();
  var notices = rows_('notices')
    .filter(function (r) { return r[0] && lastDay_(r) >= from; })
    .map(function (r) { return noticeObj_(r, tm, am); });
  var mine = sid ? rows_('confirms')
    .filter(function (r) { return r[1] === sid; })
    .map(function (r) { return r[0]; }) : [];
  return { ok: true, className: getSetting_('학급명'), year: year_(), notices: notices, confirmed: mine, callDone: callDoneMap_(), links: linksList_(), serverTime: now_() };
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
/** 호출 '다녀왔어요' — { 공지ID: [{sid, name, at}] } */
function callDoneMap_() {
  var m = {};
  rows_('calldone').forEach(function (r) { if (r[0]) (m[r[0]] = m[r[0]] || []).push({ sid: r[1], name: r[2], at: r[3] }); });
  return m;
}
function callDone_(p) {
  var sid = String(p.sid || '').trim(), nid = String(p.nid || ''), undo = p.undo === '1';
  var st = rows_('students').filter(function (x) { return x[0] === sid; })[0];
  if (!st) return { ok: false, error: '명단에 없는 학번이에요.' };
  var n = rows_('notices').filter(function (r) { return r[0] === nid; })[0];
  if (!n || n[1] !== '호출') return { ok: false, error: '호출 공지를 찾지 못했어요.' };
  return lock_(function () {
    var sh = sh_('calldone'), data = rows_('calldone');
    for (var i = data.length - 1; i >= 0; i--) {
      if (data[i][0] === nid && data[i][1] === sid) {
        if (undo) sh.deleteRow(i + 2);
        return { ok: true, name: st[1] };
      }
    }
    if (!undo) { ensureRows_(sh, sh.getLastRow() + 1); sh.getRange(sh.getLastRow() + 1, 1, 1, 4).setNumberFormat('@').setValues([[nid, sid, st[1], now_()]]); }
    return { ok: true, name: st[1] };
  });
}

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
  var tm = targetMap_(), am = attMap_();
  var notices = rows_('notices').filter(function (r) { return r[0]; })
    .map(function (r) { return noticeObj_(r, tm, am); });
  var res = { ok: true, role: role, className: getSetting_('학급명'), year: year_(), notices: notices };
  var presNids = {};
  notices.forEach(function (n) { if (n.author === '회장') presNids[n.id] = true; });
  res.callDone = callDoneMap_();
  res.requests = rows_('requests').filter(function (r) { return r[0] && (role === 'admin' || presNids[r[1]]); })
    .map(function (r) { return { id: r[0], nid: r[1], sid: r[2], name: r[3], text: r[4], at: r[5], done: r[6] === 'Y', doneAt: r[7] }; });
  if (role === 'admin') {
    res.students = rows_('students').filter(function (r) { return r[0]; })
      .map(function (r) { return { sid: r[0], name: r[1] }; });
    res.confirms = rows_('confirms').map(function (r) { return { nid: r[0], sid: r[1], name: r[2], at: r[3] }; });
    res.presidentCode = getSetting_('회장코드');
    res.links = linksList_();
    res.photos = photoStats_();
    res.installId = installId_();
  }
  return res;
}

function saveNotice_(p) {
  var role = need_(p, ['admin', 'president']);
  var n = JSON.parse(p.data || '{}');
  var author = role === 'admin' ? '담임' : '회장';
  if (role === 'president') {
    if (PRESIDENT_TYPES.indexOf(n.type) < 0) throw new Error('회장은 학급공지·교과공지·일반만 작성할 수 있어요.');
    n.popup = false; n.targets = [];
  }
  if (!n.title || !n.date || !n.type) throw new Error('유형, 날짜, 제목은 꼭 입력해 주세요.');
  var atts = (n.attachments === undefined) ? null : cleanAtts_(n.attachments);   // 저장 전에 미리 검사
  var endDate = checkEnd_(n);

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
      idx >= 0 ? now_() : '',
      (idx >= 0 && n.end === undefined) ? data[idx][11] : endDate];   // 종료일 항목을 안 보내면 기존 값 유지
    writeRow_(sh, idx >= 0 ? idx + 2 : sh.getLastRow() + 1, row);

    replaceTargets_(id, n.targets || []);
    if (atts) replaceAttachments_(id, atts);   // 첨부 항목을 보내지 않으면 기존 첨부는 그대로 둠
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
    var who = (role === 'admin' && n.author === '회장') ? '회장' : author;   // 담임이 회장 글을 나눌 때 작성자 유지
    return ['N' + Date.now() + '_' + i, n.type, n.date, n.title, n.body || '', '', n.important ? 'Y' : '', '', who, now_(), '', ''];
  });
  if (!rows.length) return { ok: true, count: 0 };
  return lock_(function () {
    var sh = sh_('notices');
    ensureRows_(sh, sh.getLastRow() + rows.length);
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
        replaceAttachments_(nid, []);
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
