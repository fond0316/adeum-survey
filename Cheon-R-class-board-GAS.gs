/**
 * 학생공지 GAS 백엔드 — Made by Cheon_R
 * 모든 요청은 JSONP(doGet + callback) 방식
 * 재배포 시 반드시 [배포 관리] → [수정] → 새 버전 (새 배포 금지)
 */

/** GAS 코드 버전 — 코드를 고칠 때마다 1씩 올려요. 관리 화면이 이 번호로 업데이트가 필요한지 알려줘요. */
var GAS_VER = 4;

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
SHEETS.timetable = { name: '시간표', headers: ['교시', '월', '화', '수', '목', '금'] };   // 우리 반 시간표 (나이스에 시간표가 없는 학교용 · N교시 제안)
SHEETS.scheduleOv = { name: '학사일정수정', headers: ['날짜', '원래일정', '상태', '내용', '수정시각'] };   // 담임이 나이스 학사일정을 고친 내용 (수정/숨김/추가)
var PRESIDENT_TYPES = ['확인', '교과', '일반'];
var DEFAULT_PW = 'classboard';   // 초기 비밀번호 — 첫 로그인 때 변경 강제
var DEFAULT_SETTINGS = [['학년도', String(new Date().getFullYear())], ['학급명', '우리 반'], ['관리자비번', DEFAULT_PW], ['회장코드', '']];

/* ---------- 진입점 ---------- */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var cb = p.callback || 'callback';
  var out;
  try {
    out = fastLoad_(p);                    // 학생 불러오기: 잠깐 기억해 둔 결과가 있으면 그걸로 (없거나 문제가 생기면 null → 아래 원래 방식)
    if (!out) { setup_(); out = route_(p); }
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService
    .createTextOutput(cb + '(' + JSON.stringify(out) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function route_(p) {
  switch (p.action) {
    case 'ping': return { ok: true, className: getSetting_('학급명'), year: year_(), gasVer: GAS_VER, cacheVer: CACHE_VER };
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
    case 'ttLoad': return ttLoad_(p);
    case 'ttSave': return ttSave_(p);
    case 'neisToday': return neisToday_(p);
    case 'neisSettings': return neisSettings_(p);
    case 'neisSave': return neisSave_(p);
    case 'neisSearch': return neisSearch_(p);
    case 'neisTest': return neisTest_(p);
    case 'neisPeriod': return neisPeriod_(p);
    case 'neisCalDots': return neisCalDots_(p);
    case 'neisMonth': return neisMonth_(p);
    case 'neisCalList': return neisCalList_(p);
    case 'neisOvSet': return neisOvSet_(p);
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
/** 수정 버전 — 1/1000초까지 (담임·회장이 같은 초에 고쳐도 구분) */
function verNow_() { return now_() + '.' + ('00' + (Date.now() % 1000)).slice(-3); }
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
  return { ok: true, role: role, gasVer: GAS_VER, installId: installId_(), className: getSetting_('학급명'),
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
function lock_(fn, scope) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { cbump_(scope); lock.releaseLock(); }   // 저장이 끝나면(실패해도) 기억해 둔 불러오기 결과를 비워서 바로 새 내용이 보이게
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
  return { ok: true, className: getSetting_('학급명'), year: year_(), notices: notices, confirmed: mine, callDone: callDoneMap_(), links: linksList_(), gasVer: GAS_VER, serverTime: now_() };
}

/* ---------- 빠른 불러오기 (학생 21~30명이 한꺼번에 열어도 가볍게) ----------
 * 학생 앱의 '불러오기'는 시트를 여러 개 읽어서 한꺼번에 몰리면 느려져요.
 * 읽은 결과를 CACHE_TTL초 동안 기억해 두고, 공지·설정·확인 등 무언가 저장되면 lock_ 가 바로 비워요.
 * 어떤 단계든 문제가 생기면 null 을 돌려줘서 원래 방식(load_)으로 처리돼요.
 * 시트를 직접 고친 경우에는 최대 CACHE_TTL초 뒤에 반영돼요. */
var CACHE_VER = 1, CACHE_TTL = 30, CACHE_CHUNK = 30000;
function fastLoad_(p) {
  if (p.action !== 'load') return null;
  try {
    var hit = cacheRead_();
    if (hit.base && hit.cf) return loadAssemble_(hit.base, hit.cf, p);
    var lock = null, got = false;
    try { lock = LockService.getDocumentLock(); got = !!lock && lock.tryLock(8000); } catch (e) { got = false; }   // 한 명만 시트를 읽고 나머지는 그 결과를 받아가요
    try {
      if (got) { hit = cacheRead_(); if (hit.base && hit.cf) return loadAssemble_(hit.base, hit.cf, p); }
      setup_();
      var gens = cacheGens_();                       // 읽기 전에 번호를 적어 둠 (읽는 동안 저장이 있었다면 이 결과는 버려짐)
      var base = hit.base, cf = hit.cf;
      if (!base) {                                   // 공지·설정·링크 등 (확인 체크 제외)
        base = load_({});
        delete base.confirmed; delete base.serverTime;
        cacheWrite_('base', gens.base, base);
      }
      if (!cf) {                                     // 확인 체크 — 학번별 목록
        cf = {};
        rows_('confirms').forEach(function (r) { var k = 's' + r[1]; (cf[k] = cf[k] || []).push(r[0]); });
        cacheWrite_('cf', gens.cf, cf);
      }
      return loadAssemble_(base, cf, p);
    } finally {
      if (got) { try { lock.releaseLock(); } catch (e2) {} }
    }
  } catch (err) {
    return null;
  }
}
function loadAssemble_(base, cf, p) {
  var sid = String(p.sid || ''), o = {};
  Object.keys(base).forEach(function (k) { o[k] = base[k]; });
  o.confirmed = (sid && Object.prototype.hasOwnProperty.call(cf, 's' + sid)) ? cf['s' + sid].slice() : [];
  o.serverTime = now_();
  return o;
}
function cacheGens_() {
  var g = {};
  try { g = CacheService.getScriptCache().getAll(['gen_base', 'gen_cf']) || {}; } catch (e) {}
  return { base: g.gen_base || '0', cf: g.gen_cf || '0' };
}
/** scope: 'cf' = 확인기록만 바뀜, 그 외(없음) = 전부 비움 */
function cbump_(scope) {
  try {
    var kv = { gen_cf: String(Date.now()) + Math.floor(Math.random() * 1000) };
    if (scope !== 'cf') kv.gen_base = String(Date.now()) + Math.floor(Math.random() * 1000);
    CacheService.getScriptCache().putAll(kv, 21600);
  } catch (e) {}
}
/** 긴 결과는 조각내서 저장 (한 칸 100KB 제한 + 글자 반쪽이 잘리지 않게) */
function cacheWrite_(name, gen, obj) {
  try {
    var s = JSON.stringify(obj), parts = [], i = 0;
    while (i < s.length) {
      var e = Math.min(i + CACHE_CHUNK, s.length);
      if (e < s.length) { var cc = s.charCodeAt(e - 1); if (cc >= 0xD800 && cc <= 0xDBFF) e--; }
      parts.push(s.substring(i, e)); i = e;
    }
    var pre = 'ld_' + name + '_' + gen + '_', kv = {};
    parts.forEach(function (x, n) { kv[pre + n] = x; });
    kv[pre + 'm'] = parts.length + ':' + s.length;
    CacheService.getScriptCache().putAll(kv, CACHE_TTL);
  } catch (e) {}
}
/** 기억해 둔 결과 읽기 — { base, cf } 각각 없거나 길이가 안 맞으면 null */
function cacheRead_() {
  var res = { base: null, cf: null };
  try {
    var c = CacheService.getScriptCache(), gens = cacheGens_();
    var pairs = [['base', 'ld_base_' + gens.base + '_'], ['cf', 'ld_cf_' + gens.cf + '_']];
    var m = c.getAll([pairs[0][1] + 'm', pairs[1][1] + 'm']) || {};
    var keys = [], info = [];
    pairs.forEach(function (pr) {
      var meta = m[pr[1] + 'm'];
      if (!meta) return;
      var mm = String(meta).split(':'), n = Number(mm[0]), len = Number(mm[1]);
      if (!(n >= 1) || !(len >= 0)) return;
      for (var i = 0; i < n; i++) keys.push(pr[1] + i);
      info.push({ name: pr[0], pre: pr[1], n: n, len: len });
    });
    if (!info.length) return res;
    var got = c.getAll(keys) || {};
    info.forEach(function (it) {
      var s = '', j;
      for (j = 0; j < it.n; j++) {
        var part = got[it.pre + j];
        if (part === undefined || part === null) return;
        s += part;
      }
      if (s.length !== it.len) return;
      try { res[it.name] = JSON.parse(s); } catch (e) {}
    });
  } catch (e) {}
  return res;
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
  }, 'cf');
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
  var rid = /^R[a-z0-9]{10,30}$/.test(String(p.rid || '')) ? String(p.rid) : '';
  return lock_(function () {
    if (rid && rows_('requests').some(function (r) { return r[0] === rid; })) return { ok: true, name: st[1], already: true };   // 응답이 늦어 다시 보낸 같은 요청
    sh_('requests').appendRow([rid || ('R' + Date.now()), nid, sid, st[1], text, now_(), '', '']);
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
  var res = { ok: true, role: role, gasVer: GAS_VER, className: getSetting_('학급명'), year: year_(), notices: notices };
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
      if (n.ifVer !== undefined) {                                   // 내가 보고 고친 버전이 지금도 그대로인지
        var cur = data[idx][10] || data[idx][9];
        if (String(n.ifVer) !== String(cur)) {
          var sameNow = data[idx][1] === n.type && dateStr_(data[idx][2]) === n.date && data[idx][3] === n.title && String(data[idx][4]) === String(n.body || '');
          if (sameNow) return { ok: true, id: String(n.id), already: true };   // 응답이 늦어 다시 보낸 같은 수정
          throw new Error('그사이 이 공지가 바뀌었어요. 목록을 새로 불러온 뒤 다시 해 주세요.');
        }
      }
    } else if (okCid_(n.cid)) {
      for (var j = 0; j < data.length; j++) if (data[j][0] === n.cid) return { ok: true, id: n.cid, already: true };   // 이미 저장된 새 공지
    }
    var id = n.id ? String(n.id) : (okCid_(n.cid) ? n.cid : 'N' + Date.now());
    var row = [id, n.type, n.date, n.title, n.body || '', n.memo || '',
      n.important ? 'Y' : '', n.popup ? 'Y' : '',
      idx >= 0 ? data[idx][8] : author,
      idx >= 0 ? data[idx][9] : now_(),
      idx >= 0 ? verNow_() : '',
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
  list.forEach(function (n) {
    if (allowed.indexOf(n.type) < 0) throw new Error('올릴 수 없는 유형이 있어요.');
    if (!n.title || !n.date) throw new Error('날짜와 제목이 빠진 줄이 있어요.');
  });
  if (!list.length) return { ok: true, count: 0, skipped: 0 };
  return lock_(function () {
    var sh = sh_('notices'), data = rows_('notices'), have = {}, recent = {}, skipped = 0;
    var since = Utilities.formatDate(new Date(Date.now() - 10 * 60000), TZ, 'yyyy-MM-dd HH:mm:ss');
    data.forEach(function (r) {
      if (!r[0]) return;
      have[r[0]] = true;
      if (r[9] >= since) recent[[r[1], dateStr_(r[2]), r[3], r[8]].join('|')] = true;
    });
    var rows = [];
    list.forEach(function (n, i) {
      var who = (role === 'admin' && n.author === '회장') ? '회장' : author;   // 담임이 회장 글을 나눌 때 작성자 유지
      var cid = okCid_(n.cid) ? n.cid : '';
      if (cid && have[cid]) { skipped++; return; }                              // 이미 저장된 요청 (응답이 늦어 다시 누른 경우)
      var key = [n.type, n.date, n.title, who].join('|');
      if (!cid && recent[key]) { skipped++; return; }                           // 옛 화면: 10분 안에 똑같은 공지
      have[cid] = true; recent[key] = true;
      rows.push([cid || ('N' + Date.now() + '_' + i), n.type, n.date, n.title, n.body || '', '', n.important ? 'Y' : '', '', who, now_(), '', '']);
    });
    if (rows.length) {
      ensureRows_(sh, sh.getLastRow() + rows.length);
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setNumberFormat('@').setValues(rows);
    }
    return { ok: true, count: rows.length, skipped: skipped };
  });
}
/** 앱이 만든 공지 번호 — 같은 요청을 두 번 받아도 한 번만 저장하기 위해 */
function okCid_(s) { return typeof s === 'string' && /^N[a-z0-9]{10,30}$/.test(s); }

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


/* =====================================================================
 * 나이스(NEIS) 교육정보 개방포털 연동 — v2에서 추가
 *  - 급식·학사일정은 학생 화면의 '학교 알림판' 카드에, 시간표는 공지 작성 때 'N교시' 제안에만 써요.
 *  - 인증키·학교코드는 이 시트의 '설정' 탭에만 저장돼요 (등록부·학생 화면으로 보내지 않아요).
 *  - 처음 한 번 GAS 편집기에서 '나이스권한승인' 함수를 실행해 외부 연결 권한을 승인해야 해요.
 * ===================================================================== */
var NEIS_BASE = 'https://open.neis.go.kr/hub/';
var NEIS_TT = { '고등학교': 'hisTimetable', '중학교': 'misTimetable', '초등학교': 'elsTimetable', '특수학교': 'spsTimetable' };
var NEIS_MEAL = { '조식': '1', '점심': '2', '석식': '3' };

/** 수동 실행용 — 외부 서버(나이스) 연결 권한 승인창을 띄워요 */
function 나이스권한승인() {
  var r = UrlFetchApp.fetch(NEIS_BASE + 'schoolInfo?Type=json&pSize=1&SCHUL_NM=' + encodeURIComponent('서울'), { muteHttpExceptions: true });
  Logger.log('나이스 연결 확인: HTTP ' + r.getResponseCode());
}

function neisHas_(s, w) { return (',' + String(s || '') + ',').indexOf(',' + w + ',') >= 0; }
function neisYmd_(d) { return String(d).replace(/-/g, ''); }
function neisDate_(v) {
  v = String(v || '');
  if (!v) return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(v)) throw new Error('날짜 형식이 올바르지 않아요.');
  return v;
}
/** 나이스 호출 + 잠깐 기억(캐시). 결과 없음(INFO-200)은 빈 목록 */
function neisFetch_(svc, params, ttl) {
  var key = getSetting_('NEIS키');
  if (!key) throw new Error('나이스 인증키가 아직 없어요.');
  var q = Object.keys(params).sort().map(function (k) { return k + '=' + encodeURIComponent(params[k]); }).join('&');
  var ck = 'neis_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, svc + '?' + q)).slice(0, 24);
  var cache = CacheService.getScriptCache();
  try { var hit = cache.get(ck); if (hit) return JSON.parse(hit); } catch (e) {}
  var url = NEIS_BASE + svc + '?KEY=' + encodeURIComponent(key) + '&Type=json&pIndex=1&pSize=100&' + q;
  var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('나이스 서버 응답이 이상해요 (HTTP ' + res.getResponseCode() + ').');
  var j = JSON.parse(res.getContentText());
  var rows = [];
  if (j.RESULT) {
    var code = String(j.RESULT.CODE || '');
    if (code !== 'INFO-200') throw new Error('나이스 응답: ' + code + ' ' + (j.RESULT.MESSAGE || ''));
  } else if (j[svc]) {
    j[svc].forEach(function (blk) {
      if (blk.row) rows = rows.concat(blk.row);
      if (blk.head) blk.head.forEach(function (h) {
        if (h.RESULT && String(h.RESULT.CODE) !== 'INFO-000' && String(h.RESULT.CODE) !== 'INFO-200') throw new Error('나이스 응답: ' + h.RESULT.CODE + ' ' + (h.RESULT.MESSAGE || ''));
      });
    });
  }
  try { cache.put(ck, JSON.stringify(rows), ttl || 3600); } catch (e) {}
  return rows;
}
/** 메뉴 이름만 남기기 — 알레르기 번호·괄호·별표 제거 */
function neisMenu_(raw) {
  return String(raw || '').split(/<br\s*\/?>/i).map(function (m) {
    return m.replace(/\s*\(\s*[\d.,\s]+\)/g, '').replace(/[\d.]+\s*$/, '').replace(/[*＊]/g, '').replace(/\s+/g, ' ').trim();
  }).filter(function (m) { return m; });
}
function neisSchool_() {
  var a = getSetting_('NEIS교육청코드'), c = getSetting_('NEIS학교코드');
  if (!a || !c) throw new Error('학교가 아직 선택되지 않았어요.');
  return { ATPT_OFCDC_SC_CODE: a, SD_SCHUL_CODE: c };
}
function neisEnabled_() { return !!(getSetting_('NEIS키') && getSetting_('NEIS교육청코드') && getSetting_('NEIS학교코드')); }
function neisStrip_() { var v = getSetting_('NEIS띠'); return v === '' ? '학사일정,점심' : v; }
function neisOpen_() { var v = getSetting_('NEIS펼침'); return v === '' ? '점심' : v; }
function neisMeals_(date, names) {
  var out = {};
  names.forEach(function (nm) {
    var sc = neisSchool_(); sc.MLSV_YMD = neisYmd_(date); sc.MMEAL_SC_CODE = NEIS_MEAL[nm];
    var rows = neisFetch_('mealServiceDietInfo', sc, 7200);
    out[nm] = rows.length ? neisMenu_(rows[0].DDISH_NM) : [];
  });
  return out;
}
function neisSchedule_(date) {
  var sc = neisSchool_(); sc.AA_YMD = neisYmd_(date);
  var g = Number(getSetting_('NEIS학년')) || 0;
  var gf = ['', 'ONE', 'TW', 'THREE'];
  var seen = {}, list = [];
  neisFetch_('SchoolSchedule', sc, 7200).forEach(function (r) {
    if (g >= 1 && g <= 3) {
      var mine = r[gf[g] + '_GRADE_EVENT_YN'];
      var any = r.ONE_GRADE_EVENT_YN === 'Y' || r.TW_GRADE_EVENT_YN === 'Y' || r.THREE_GRADE_EVENT_YN === 'Y';
      if (mine === 'N' && any) return;   // 다른 학년만의 행사는 제외
    }
    var nm = String(r.EVENT_NM || '').trim();
    if (nm && !seen[nm]) { seen[nm] = 1; list.push(nm); }
  });
  return list;
}

/** 학생 화면용 — 로그인 없이 호출 (키·학교코드는 돌려주지 않아요) */
function neisToday_(p) {
  if (!neisEnabled_()) return { ok: true, enabled: false };
  var date = neisDate_(p.date);
  var strip = neisStrip_(), open = neisOpen_();
  var need = [];
  ['조식', '점심', '석식'].forEach(function (m) { if (neisHas_(strip, m) || neisHas_(open, m)) need.push(m); });
  var res = { ok: true, enabled: true, date: date, strip: strip, open: open, schedule: [], meals: {} };
  res.schedule = neisSchedule_(date);
  res.meals = neisMeals_(date, need);
  return res;
}

/** 담임 설정 화면용 */
function neisSettings_(p) {
  need_(p, ['admin']);
  return { ok: true, hasKey: !!getSetting_('NEIS키'), atpt: getSetting_('NEIS교육청코드'), code: getSetting_('NEIS학교코드'),
    school: getSetting_('NEIS학교명'), level: getSetting_('NEIS학교급'), grade: getSetting_('NEIS학년'), cls: getSetting_('NEIS반'),
    strip: neisStrip_(), open: neisOpen_() };
}
function neisSave_(p) {
  need_(p, ['admin']);
  var s = JSON.parse(p.data || '{}');
  return lock_(function () {
    if (s.clearKey) setSetting_('NEIS키', '');
    else if (s.key !== undefined && String(s.key).trim()) {
      var k = String(s.key).trim();
      if (!/^[A-Za-z0-9]{10,64}$/.test(k)) throw new Error('인증키 형식이 올바르지 않아요. 복사한 키를 다시 확인해 주세요.');
      setSetting_('NEIS키', k);
    }
    if (s.atpt !== undefined) setSetting_('NEIS교육청코드', String(s.atpt).replace(/[^A-Za-z0-9]/g, ''));
    if (s.code !== undefined) setSetting_('NEIS학교코드', String(s.code).replace(/[^A-Za-z0-9]/g, ''));
    if (s.school !== undefined) setSetting_('NEIS학교명', String(s.school).slice(0, 60));
    if (s.level !== undefined) setSetting_('NEIS학교급', String(s.level).slice(0, 20));
    if (s.grade !== undefined) setSetting_('NEIS학년', String(Number(s.grade) || ''));
    if (s.cls !== undefined) setSetting_('NEIS반', String(Number(s.cls) || ''));
    var ok = ['학사일정', '조식', '점심', '석식'];
    if (s.strip !== undefined) setSetting_('NEIS띠', String(s.strip).split(',').filter(function (x) { return ok.indexOf(x) >= 0; }).join(','));
    if (s.open !== undefined) setSetting_('NEIS펼침', String(s.open).split(',').filter(function (x) { return ['조식', '점심', '석식'].indexOf(x) >= 0; }).join(','));
    return { ok: true };
  });
}
/** 학교 이름으로 검색 (저장된 인증키 사용) */
function neisSearch_(p) {
  need_(p, ['admin']);
  var q = String(p.q || '').trim();
  if (q.length < 2) throw new Error('학교 이름을 2글자 이상 입력해 주세요.');
  var rows = neisFetch_('schoolInfo', { SCHUL_NM: q }, 600);
  return { ok: true, list: rows.slice(0, 20).map(function (r) {
    return { atpt: r.ATPT_OFCDC_SC_CODE, code: r.SD_SCHUL_CODE, name: r.SCHUL_NM, level: r.SCHUL_KND_SC_NM, addr: r.ORG_RDNMA || '', office: r.ATPT_OFCDC_SC_NM || '' };
  }) };
}
/** 연결 확인 — 오늘(또는 지정일) 급식·학사일정·시간표가 가져와지는지 */
function neisTest_(p) {
  need_(p, ['admin']);
  if (!neisEnabled_()) throw new Error('인증키와 학교를 먼저 저장해 주세요.');
  var date = neisDate_(p.date);
  var res = { ok: true, date: date };
  try { res.schedule = neisSchedule_(date); } catch (e) { res.scheduleErr = String(e.message || e); }
  try { res.meals = neisMeals_(date, ['점심']); } catch (e) { res.mealErr = String(e.message || e); }
  try { var pl = neisPeriods_(date); res.periods = pl.length; res.periodList = pl.map(function (x) { return x.perio + '교시 ' + x.name; }).join(' · '); } catch (e) { res.periodErr = String(e.message || e); }
  try { res.ttList = ttDay_(date); } catch (e) {}
  if (!res.periods && !res.periodErr) { try { res.diag = neisDiag_(date); } catch (e) { res.diag = '진단 중 오류: ' + String(e.message || e); } }
  return res;
}
/** 시간표가 비어 있을 때 이유를 찾기 위한 진단 (날짜 없이·학급 없이 넓혀서 조회) */
function neisDiag_(date) {
  var level = getSetting_('NEIS학교급'), svc = NEIS_TT[level];
  if (!svc) return '학교급 정보가 없거나 시간표를 지원하지 않는 학교급이에요 (' + (level || '비어 있음') + '). 학교를 다시 검색해서 선택해 주세요.';
  var g = getSetting_('NEIS학년'), c = getSetting_('NEIS반');
  if (!g || !c) return '학년·반이 비어 있어요. 학년과 반을 입력하고 저장해 주세요.';
  var out = ['서비스 ' + svc + ' · 조회 학급 ' + g + '학년 ' + c + '반 · 조회일 ' + date];
  function span(rows) {
    var d = rows.map(function (r) { return String(r.ALL_TI_YMD || ''); }).filter(function (x) { return x; }).sort();
    return rows.length + '칸' + (d.length ? ' (' + d[0] + ' ~ ' + d[d.length - 1] + ')' : '');
  }
  var sc = neisSchool_(); sc.GRADE = g; sc.CLASS_NM = c;
  out.push('우리 반 전체 기간: ' + span(neisFetch_(svc, sc, 600)));
  var sc2 = neisSchool_(); sc2.GRADE = g;
  var r2 = neisFetch_(svc, sc2, 600);
  var cls = {}; r2.forEach(function (r) { cls[String(r.CLASS_NM)] = 1; });
  out.push(g + '학년 전체: ' + span(r2) + (Object.keys(cls).length ? ' · 반 이름 ' + Object.keys(cls).sort().slice(0, 15).join(',') : ''));
  out.push('학교 전체: ' + span(neisFetch_(svc, neisSchool_(), 600)));
  return out.join('\n');
}
function neisPeriods_(date) {
  var svc = NEIS_TT[getSetting_('NEIS학교급')];
  var g = getSetting_('NEIS학년'), c = getSetting_('NEIS반');
  if (!svc || !g || !c) return [];
  var sc = neisSchool_(); sc.ALL_TI_YMD = neisYmd_(date); sc.GRADE = g; sc.CLASS_NM = c;
  var seen = {}, out = [];
  neisFetch_(svc, sc, 3600).forEach(function (r) {
    var per = String(r.PERIO || '').trim(), nm = String(r.ITRT_CNTNT || '').trim();
    if (per && nm && !seen[per + '|' + nm]) { seen[per + '|' + nm] = 1; out.push({ perio: Number(per) || per, name: nm }); }
  });
  return out.sort(function (a, b) { return Number(a.perio) - Number(b.perio); });
}
/** 공지 작성용 — 그날 시간표에서 과목이 몇 교시인지 (담임·회장) */
function neisPeriod_(p) {
  need_(p, ['admin', 'president']);
  var sub = String(p.subject || '').replace(/\s+/g, '');
  var hasTt = ttHas_();
  if (!neisEnabled_() && !hasTt) return { ok: true, enabled: false, periods: [] };
  if (!sub) return { ok: true, enabled: true, periods: [] };
  var date = neisDate_(p.date);
  var list = [];
  if (neisEnabled_()) {   // 나이스 시간표가 있으면 그걸 먼저 (없거나 안 되면 아래 우리 반 시간표)
    try { list = neisPeriods_(date).filter(function (x) { return ttMatch_(x.name, sub); }); } catch (e) { list = []; }
  }
  if (!list.length && hasTt) list = ttPeriods_(date, sub);
  return { ok: true, enabled: true, date: date, periods: list };
}

/* ---------- 우리 반 시간표 (시트 '시간표' 탭 · 월~금 × 1~7교시) ---------- */
var TT_ALIAS = { '과학': '과', '사회': '사', '체육': '체', '한국사': '국사', '일본어': '일', '일어': '일', '한문': '한문', '진로': '진로', '미술': '미', '음악': '음' };
function ttGrid_() {
  var g = [];
  for (var i = 0; i < 7; i++) g.push(['', '', '', '', '']);
  rows_('timetable').forEach(function (r) {
    var m = String(r[0]).match(/^(\d)/);
    var n = m ? Number(m[1]) : 0;
    if (n >= 1 && n <= 7) for (var c = 0; c < 5; c++) g[n - 1][c] = r[c + 1] || '';
  });
  return g;
}
function ttHas_() {
  return ttGrid_().some(function (row) { return row.some(function (v) { return v && v !== '공강' && v !== '자율'; }); });
}
function ttLoad_(p) {
  need_(p, ['admin']);
  return { ok: true, grid: ttGrid_() };
}
function ttSave_(p) {
  need_(p, ['admin']);
  var d = JSON.parse(p.data || '{}');
  var grid = d.grid;
  if (!grid || grid.length !== 7) throw new Error('시간표 형식이 올바르지 않아요.');
  var days = ['월', '화', '수', '목', '금'];
  var out = grid.map(function (row, i) {
    if (!row || row.length !== 5) throw new Error('시간표 형식이 올바르지 않아요.');
    return [(i + 1) + '교시'].concat(row.map(function (v) { return String(v || '').replace(/\s+/g, ' ').trim().slice(0, 30); }));
  });
  return lock_(function () {
    var sh = sh_('timetable');
    sh.getRange(1, 1, 1, 6).setValues([SHEETS.timetable.headers]).setFontWeight('bold');
    sh.getRange(2, 1, 7, 6).setNumberFormat('@').setValues(out);
    return { ok: true };
  });
}
/** 과목 칸의 '기본 이름' — 공통수학2→수학, 통합과학2→과학, 국어A→국어, 과B→과, 한국사2→한국사 */
function ttBase_(cell) {
  var first = String(cell || '').trim().split(/\s+/)[0] || '';
  var cs = first.replace(/^[A-Da-d]_/, '').replace(/^(공통|통합|기초|심화)/, '');
  return cs.replace(/[0-9ⅠⅡⅢⅣⅤ]+$/, '').replace(/[A-Da-d]$/, '');
}
/** 공지 제목의 과목(sub)이 시간표 칸의 과목과 맞는지 — '수학숙제'(띄어쓰기 없이)도, '과학'→'과B'·'통합과학2'도 인정 */
function ttMatch_(cell, sub) {
  var first = String(cell || '').trim().split(/\s+/)[0] || '';
  if (!first || first === '공강' || first === '자율') return false;
  var b = ttBase_(cell);
  if (!b || !sub) return false;
  if (b.length >= 2 && sub.indexOf(b) === 0) return true;      // 제목이 과목으로 시작 (수학숙제 ← 수학)
  if (sub.length >= 2 && b.indexOf(sub) === 0) return true;    // 짧게 쓴 경우 (국어 ← 국어회화)
  for (var k in TT_ALIAS) {                                     // 약칭 (과학숙제 ← 과B, 체육 ← 체A)
    if (sub.indexOf(k) === 0 && b.indexOf(TT_ALIAS[k]) === 0) return true;
  }
  return false;
}
/** 그날(요일) 우리 반 시간표 시트의 교시별 과목 — 확인용 */
function ttDay_(date) {
  var m = String(date).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return '';
  var wd = new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]))).getUTCDay();
  if (wd < 1 || wd > 5) return '';
  var g = ttGrid_(), out = [];
  for (var i = 0; i < 7; i++) if (g[i][wd - 1]) out.push((i + 1) + '교시 ' + g[i][wd - 1]);
  return out.join(' · ');
}
function ttPeriods_(date, sub) {
  var m = String(date).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return [];
  var wd = new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]))).getUTCDay();   // 0=일 … 6=토 (서버 시간대와 상관없이 날짜 그대로)
  if (wd < 1 || wd > 5) return [];
  var g = ttGrid_(), out = [];
  for (var i = 0; i < 7; i++) { var cell = g[i][wd - 1]; if (ttMatch_(cell, sub)) out.push({ perio: i + 1, name: cell }); }
  return out;
}


/* ---------- 학사일정 달력 점 · 담임 보정 (v3) ---------- */
/** 기본 기간 = 학년도 전체 (3/1 ~ 다음 해 2월 말) */
function neisDefRange_() {
  var t = new Date(), y = t.getFullYear(), m = t.getMonth() + 1;
  var sy = m >= 3 ? y : y - 1;
  var last = new Date(sy + 1, 2, 0).getDate();
  return { from: sy + '-03-01', to: (sy + 1) + '-02-' + (last < 10 ? '0' + last : last) };
}
function neisRange_() {
  var d = neisDefRange_();
  var f = getSetting_('NEIS일정시작'), t = getSetting_('NEIS일정끝');
  var ok = /^\d{4}-\d{2}-\d{2}$/;
  if (!ok.test(f) || !ok.test(t) || f > t) return { from: d.from, to: d.to, custom: false };
  return { from: f, to: t, custom: true };
}
/** 기간 전체를 한 번에 받아서(1000건씩) 날짜별 이름 목록으로 정리 · 주말/공휴일/휴업일 제외 · 우리 학년 것만 */
function neisYearMap_() {
  var rg = neisRange_();
  var g = Number(getSetting_('NEIS학년')) || 0;
  var ck = 'neisYr_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,
    [getSetting_('NEIS학교코드'), rg.from, rg.to, g].join('|'))).slice(0, 20);
  var cache = CacheService.getScriptCache();
  try { var hit = cache.get(ck); if (hit) return JSON.parse(hit); } catch (e) {}
  var key = getSetting_('NEIS키');
  var sc = neisSchool_();
  var gf = ['', 'ONE', 'TW', 'THREE'];
  var map = {}, total = 0;
  for (var page = 1; page <= 5; page++) {
    var url = NEIS_BASE + 'SchoolSchedule?KEY=' + encodeURIComponent(key) + '&Type=json&pIndex=' + page + '&pSize=1000'
      + '&ATPT_OFCDC_SC_CODE=' + sc.ATPT_OFCDC_SC_CODE + '&SD_SCHUL_CODE=' + sc.SD_SCHUL_CODE
      + '&AA_FROM_YMD=' + neisYmd_(rg.from) + '&AA_TO_YMD=' + neisYmd_(rg.to);
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw new Error('나이스 서버 응답이 이상해요 (HTTP ' + res.getResponseCode() + ').');
    var j = JSON.parse(res.getContentText());
    var rows = [];
    if (j.RESULT) {
      if (String(j.RESULT.CODE || '') !== 'INFO-200') throw new Error('나이스 응답: ' + j.RESULT.CODE + ' ' + (j.RESULT.MESSAGE || ''));
    } else if (j.SchoolSchedule) {
      j.SchoolSchedule.forEach(function (blk) {
        if (blk.row) rows = rows.concat(blk.row);
        if (blk.head) blk.head.forEach(function (h) {
          if (h.RESULT && String(h.RESULT.CODE) !== 'INFO-000' && String(h.RESULT.CODE) !== 'INFO-200') throw new Error('나이스 응답: ' + h.RESULT.CODE + ' ' + (h.RESULT.MESSAGE || ''));
        });
      });
    }
    total += rows.length;
    rows.forEach(function (r) {
      var ymd = String(r.AA_YMD || ''); if (ymd.length !== 8) return;
      var d = ymd.slice(0, 4) + '-' + ymd.slice(4, 6) + '-' + ymd.slice(6, 8);
      var nm = String(r.EVENT_NM || '').trim(); if (!nm) return;
      var wd = new Date(Number(ymd.slice(0, 4)), Number(ymd.slice(4, 6)) - 1, Number(ymd.slice(6, 8))).getDay();
      if (wd === 0 || wd === 6) return;                                            // 주말
      var sb = String(r.SBTR_DD_SC_NM || '');
      if (sb === '공휴일' || sb === '휴업일' || /토요휴업일/.test(nm)) return;   // 공휴일·휴업일
      if (g >= 1 && g <= 3) {
        var mine = r[gf[g] + '_GRADE_EVENT_YN'];
        var any = r.ONE_GRADE_EVENT_YN === 'Y' || r.TW_GRADE_EVENT_YN === 'Y' || r.THREE_GRADE_EVENT_YN === 'Y';
        if (mine === 'N' && any) return;
      }
      var arr = map[d] = map[d] || [];
      if (arr.indexOf(nm) < 0) arr.push(nm);
    });
    if (rows.length < 1000) break;
  }
  var out = { map: map, total: total };
  try { cache.put(ck, JSON.stringify(out), 10800); } catch (e) {}
  return out;
}
/** 담임이 고친 내용 읽기 */
function neisOvRows_() {
  return rows_('scheduleOv').filter(function (r) { return r[0] && r[2]; }).map(function (r) {
    return { date: r[0], orig: r[1], status: r[2], text: r[3] };
  });
}
/** 나이스 이름 목록에 담임 보정(숨김/수정/추가)을 덮어씌움 */
function neisApplyOv_(date, list, ov) {
  var out = list.slice();
  (ov || neisOvRows_()).forEach(function (o) {
    if (o.date !== date) return;
    var i = out.indexOf(o.orig);
    if (o.status === '숨김') { if (i >= 0) out.splice(i, 1); }
    else if (o.status === '수정') { if (i >= 0) { if (o.text) out[i] = o.text; } }
    else if (o.status === '추가') { var t = o.text || o.orig; if (t && out.indexOf(t) < 0) out.push(t); }
  });
  return out;
}
/* 학교 알림판의 학사일정에도 담임 보정이 반영되게 — 기존 함수를 감싸요 */
(function () {
  var base = neisSchedule_;
  neisSchedule_ = function (date) {
    var list = base(date);
    try { return neisApplyOv_(neisDate_(date), list); } catch (e) { return list; }
  };
})();
/** 학생 화면용 — 달력에 점 찍을 날짜들 (로그인 없이) */
function neisCalDots_(p) {
  if (!neisEnabled_()) return { ok: true, enabled: false, dots: [] };
  var strip = neisStrip_();
  if (!neisHas_(strip, '학사일정')) return { ok: true, enabled: true, dots: [] };
  var yr = neisYearMap_(), ov = neisOvRows_(), dots = [];
  var dates = {};
  Object.keys(yr.map).forEach(function (d) { dates[d] = 1; });
  ov.forEach(function (o) { dates[o.date] = 1; });
  Object.keys(dates).sort().forEach(function (d) {
    if (neisApplyOv_(d, yr.map[d] || [], ov).length) dots.push(d);
  });
  var rg = neisRange_();
  return { ok: true, enabled: true, dots: dots, from: rg.from, to: rg.to };
}
/** 담임 화면용 — 기간 안 학사일정 목록(보정 상태 포함) */
function neisCalList_(p) {
  need_(p, ['admin']);
  if (!neisEnabled_()) throw new Error('인증키와 학교를 먼저 저장해 주세요.');
  var yr = neisYearMap_(), ov = neisOvRows_(), items = [];
  Object.keys(yr.map).sort().forEach(function (d) {
    yr.map[d].forEach(function (nm) {
      var o = ov.filter(function (x) { return x.date === d && x.orig === nm && x.status !== '추가'; })[0];
      items.push({ date: d, name: nm, status: o ? o.status : '', text: o ? o.text : '' });
    });
  });
  ov.filter(function (x) { return x.status === '추가'; }).forEach(function (o) {
    items.push({ date: o.date, name: o.text || o.orig, status: '추가', text: o.text || o.orig });
  });
  items.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
  var rg = neisRange_();
  return { ok: true, items: items, total: yr.total, from: rg.from, to: rg.to, custom: rg.custom };
}
/** 담임 보정 저장 — data: {date, orig, status: 수정|숨김|추가|해제, text} */
function neisOvSet_(p) {
  need_(p, ['admin']);
  var d = JSON.parse(p.data || '{}');
  var date = neisDate_(d.date), orig = String(d.orig || '').trim().slice(0, 80);
  var st = String(d.status || ''), text = String(d.text || '').replace(/\s+/g, ' ').trim().slice(0, 80);
  if (['수정', '숨김', '추가', '해제'].indexOf(st) < 0) throw new Error('상태 값이 올바르지 않아요.');
  if (st === '추가') { if (!text) throw new Error('추가할 일정 이름을 적어 주세요.'); orig = text; }
  if (st === '수정' && !text) throw new Error('바꿀 이름을 적어 주세요.');
  if (!orig) throw new Error('원래 일정 이름이 없어요.');
  return lock_(function () {
    var sh = sh_('scheduleOv');
    var last = sh.getLastRow(), rows = last < 2 ? [] : sh.getRange(2, 1, last - 1, 5).getValues().map(function (r) { return r.map(str_); });
    for (var i = rows.length - 1; i >= 0; i--) {
      if (rows[i][0] === date && rows[i][1] === orig) sh.deleteRow(i + 2);   // 같은 일정의 예전 보정은 지우고 새로 씀
    }
    if (st !== '해제') writeRow_(sh, sh.getLastRow() + 1, [date, orig, st, st === '숨김' ? '' : text, now_()]);
    return { ok: true };
  });
}
/* 설정 화면 — 기간 값도 같이 주고받기 (기존 함수를 감싸요) */
(function () {
  var bs = neisSettings_, bv = neisSave_;
  neisSettings_ = function (p) {
    var r = bs(p), rg = neisRange_();
    r.rangeFrom = rg.from; r.rangeTo = rg.to; r.rangeCustom = rg.custom;
    var df = neisDefRange_(); r.defFrom = df.from; r.defTo = df.to;
    return r;
  };
  neisSave_ = function (p) {
    var r = bv(p);
    var s = JSON.parse(p.data || '{}');
    if (s.rangeFrom !== undefined || s.rangeTo !== undefined || s.rangeReset) {
      lock_(function () {
        if (s.rangeReset) { setSetting_('NEIS일정시작', ''); setSetting_('NEIS일정끝', ''); return; }
        var ok = /^\d{4}-\d{2}-\d{2}$/;
        if (!ok.test(String(s.rangeFrom)) || !ok.test(String(s.rangeTo)) || s.rangeFrom > s.rangeTo) throw new Error('기간(시작~끝)이 올바르지 않아요.');
        setSetting_('NEIS일정시작', s.rangeFrom); setSetting_('NEIS일정끝', s.rangeTo);
      });
    }
    return r;
  };
})();


/* ---------- 학교 알림판 한 달치 한 번에 (v4) — 날짜를 누를 때마다 기다리지 않게 ---------- */
function neisMonth_(p) {
  if (!neisEnabled_()) return { ok: true, enabled: false };
  var m = String(p.month || '');
  if (!/^\d{4}-\d{2}$/.test(m)) m = neisDate_('').slice(0, 7);
  var y = Number(m.slice(0, 4)), mo = Number(m.slice(5, 7));
  if (mo < 1 || mo > 12) throw new Error('달 형식이 올바르지 않아요.');
  var last = new Date(y, mo, 0).getDate();
  var two = function (n) { return n < 10 ? '0' + n : String(n); };
  var from = m + '-01', to = m + '-' + two(last);
  var strip = neisStrip_(), open = neisOpen_();
  var need = [];
  ['조식', '점심', '석식'].forEach(function (x) { if (neisHas_(strip, x) || neisHas_(open, x)) need.push(x); });
  var days = {}, partial = false;
  for (var d = 1; d <= last; d++) days[m + '-' + two(d)] = { schedule: [], meals: {} };
  var toDate = function (ymd) { ymd = String(ymd || ''); return ymd.length === 8 ? ymd.slice(0, 4) + '-' + ymd.slice(4, 6) + '-' + ymd.slice(6, 8) : ''; };
  // 학사일정 (우리 학년 것만 · 공휴일도 알림판에는 보여줘요)
  var sc = neisSchool_(); sc.AA_FROM_YMD = neisYmd_(from); sc.AA_TO_YMD = neisYmd_(to);
  var g = Number(getSetting_('NEIS학년')) || 0, gf = ['', 'ONE', 'TW', 'THREE'];
  var srows = neisFetch_('SchoolSchedule', sc, 7200);
  if (srows.length >= 100) partial = true;
  srows.forEach(function (r) {
    var dt = toDate(r.AA_YMD); if (!days[dt]) return;
    if (g >= 1 && g <= 3) {
      var mine = r[gf[g] + '_GRADE_EVENT_YN'];
      var any = r.ONE_GRADE_EVENT_YN === 'Y' || r.TW_GRADE_EVENT_YN === 'Y' || r.THREE_GRADE_EVENT_YN === 'Y';
      if (mine === 'N' && any) return;
    }
    var nm = String(r.EVENT_NM || '').trim();
    if (nm && days[dt].schedule.indexOf(nm) < 0) days[dt].schedule.push(nm);
  });
  var ov = neisOvRows_();
  Object.keys(days).forEach(function (dt) { days[dt].schedule = neisApplyOv_(dt, days[dt].schedule, ov); });
  // 급식
  need.forEach(function (nm) {
    var mc = neisSchool_(); mc.MMEAL_SC_CODE = NEIS_MEAL[nm]; mc.MLSV_FROM_YMD = neisYmd_(from); mc.MLSV_TO_YMD = neisYmd_(to);
    var mrows = neisFetch_('mealServiceDietInfo', mc, 7200);
    if (mrows.length >= 100) partial = true;
    Object.keys(days).forEach(function (dt) { days[dt].meals[nm] = []; });
    mrows.forEach(function (r) { var dt = toDate(r.MLSV_YMD); if (days[dt]) days[dt].meals[nm] = neisMenu_(r.DDISH_NM); });
  });
  return { ok: true, enabled: true, month: m, strip: strip, open: open, days: days, partial: partial };
}
