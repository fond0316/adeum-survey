/**
 * 섬탈출원정대 GAS (Cheon-R-island-escape.gs)
 *
 * [설치]
 * 1) 구글 시트를 하나 만들고 확장 프로그램 > Apps Script 에 이 코드를 통째로 붙여넣기
 * 2) 함수 목록에서 setup 을 한 번 실행 (권한 허용) -> 시트 탭 6개가 자동으로 생겨요
 * 3) 배포 > 새 배포 > 웹 앱 (실행: 나 / 액세스: 모든 사용자)
 *    * 코드를 고친 뒤에는 반드시 [배포 관리 > 수정 > 새 버전]으로 재배포 (새 배포 금지)
 * 4) 학생 링크:  앱주소.html?gas=웹앱주소
 *
 * [시트 탭]
 *  참가자   : 입장한 사람 (조/이름/캐릭터/코인/열쇠/도움중 + 문항별 완료·오답·힌트받음)
 *  풀이기록 : 제출할 때마다 한 줄 (정답/오답)
 *  힌트     : 조원 힌트 + 담임 힌트 쪽지
 *  도움요청 : 도움 요청 기록
 *  상자     : 보석함 결과
 *  조설정   : 섬 이름 변경 + 탈출 시각
 *
 * 이 앱의 정답·힌트는 HTML 파일의 QUESTIONS 에 있어요 (GAS는 기록과 공유만 담당).
 * 관리자 입장코드: 아래 ADMIN_CODE (스크립트 속성 ADMIN_CODE 를 만들면 그 값이 우선)
 */
var SH_P = '참가자', SH_LOG = '풀이기록', SH_H = '힌트', SH_R = '도움요청', SH_B = '상자', SH_T = '조설정';
var NQ = 5;                    // 문제 수 (HTML의 QUESTIONS 개수와 같아야 해요)
var MAX_TEAM = 6;              // 조당 최대 인원
var TEAMS = [1, 2, 3, 4, 5];
var ISLAND_DEFAULT = { 1: '🐢 거북이섬', 2: '🌴 야자섬', 3: '🌋 화산섬', 4: '🪸 산호섬', 5: '🐚 조개섬' };
var ADMIN_CODE = '0413';
var CHARS = ['🦊', '🐯', '🐻', '🐼', '🐰', '🐸', '🦁', '🐧', '🐶', '🐱', '🐹', '🐨', '🐵', '🐔', '🦄', '🐙', '🦖', '🐳', '🦋', '🐢'];
var COIN_BASE = 10, COIN_FIRST = 5, COIN_HINT_PENALTY = 2, COIN_MIN = 5, COIN_HELP = 5, DOUBLE_CAP = 30;
var BOX_CODES = ['X', '5', '10', '20', 'D'];

var B_H = ['상자ID', '키', '문항', '상자1', '상자2', '상자3', '상자4', '상자5', '선택', '획득코인', '시각'];
var LOG_H = ['시각', '조', '이름', '문항', '결과', '획득코인'];
var H_H = ['ID', '시각', '조', '보낸이', '받는이', '문항', '종류', '내용', '감사'];
var R_H = ['ID', '시각', '조', '요청자', '문항', '종류', '상태'];
var T_H = ['조', '섬이름', '탈출시각ms', '탈출시각'];
function pHeaders_() {
  var h = ['키', '조', '이름', '캐릭터', '코인', '열쇠', '도움중', '입장시각', '갱신'];
  for (var i = 1; i <= NQ; i++) { h.push('Q' + i + '_완료'); h.push('Q' + i + '_오답'); h.push('Q' + i + '_힌트받음'); }
  return h;
}

/* ===== 공통 ===== */
function nowKST_() { return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'); }
function num_(v) { var n = parseInt(String(v).replace(/[^0-9]/g, ''), 10); return isNaN(n) ? 0 : n; }
function clean_(v, n) { return String(v == null ? '' : v).replace(/[<>]/g, '').replace(/\s+/g, ' ').trim().slice(0, n); }
function key_(t, n) { return 'T' + num_(t) + '_' + clean_(n, 10); }
function props_() { return PropertiesService.getScriptProperties(); }
function adminOk_(code) { return String(code || '') === (props_().getProperty('ADMIN_CODE') || ADMIN_CODE); }
function sheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    if (name === SH_P) sh.getRange(1, 1, sh.getMaxRows(), 3).setNumberFormat('@');   // 키·이름이 날짜로 바뀌지 않게
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#dff0ff');
    sh.setFrozenRows(1);
  }
  return sh;
}
function setup() {
  sheet_(SH_P, pHeaders_()); sheet_(SH_LOG, LOG_H); sheet_(SH_H, H_H); sheet_(SH_R, R_H); sheet_(SH_B, B_H); sheet_(SH_T, T_H);
  return 'OK';
}
function jsonp_(cb, obj) {
  var txt = JSON.stringify(obj);
  if (cb && /^[A-Za-z0-9_]+$/.test(cb)) return ContentService.createTextOutput(cb + '(' + txt + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(txt).setMimeType(ContentService.MimeType.JSON);
}
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return fn(); } finally { try { lock.releaseLock(); } catch (x) {} }
}
function bust_() { try { CacheService.getScriptCache().remove('snap'); } catch (e) {} }
function shuffle_(a) { for (var i = a.length - 1; i > 0; i--) { var j = Math.floor(Math.random() * (i + 1)); var t = a[i]; a[i] = a[j]; a[j] = t; } return a; }
function nextId_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return 1;
  return num_(sh.getRange(last, 1, 1, 1).getValues()[0][0]) + 1;
}

/* ===== 게임 시간·설정 ===== */
function game_() {
  var pr = props_(), st = Number(pr.getProperty('START')) || 0, dur = Number(pr.getProperty('DUR'));
  if (!(dur > 0)) dur = 15;
  var now = Date.now(), end = st ? st + dur * 60000 : 0;
  return { phase: !st ? 'ready' : (now > end ? 'over' : 'running'), startAt: st, endAt: end, durationMin: Math.round(dur * 10) / 10 };
}
function cfg_() {
  var pr = props_();
  return { always: pr.getProperty('ALWAYS') === '1', minTeam: Math.max(1, num_(pr.getProperty('MINTEAM')) || 2) };
}

/* ===== 시트 읽기·쓰기 ===== */
function readP_() {
  var H = pHeaders_(), sh = sheet_(SH_P, H), last = sh.getLastRow(), list = [];
  if (last >= 2) {
    var v = sh.getRange(2, 1, last - 1, H.length).getValues();
    for (var i = 0; i < v.length; i++) {
      var r = v[i];
      if (!r[0]) continue;
      var p = { row: i + 2, key: String(r[0]), team: num_(r[1]), name: String(r[2]), ch: String(r[3] || ''), coins: Number(r[4]) || 0, hasKey: String(r[5]) === 'O', helping: num_(r[6]), joined: String(r[7]), done: [], wrong: [], hinted: [] };
      for (var q = 0; q < NQ; q++) { p.done.push(String(r[9 + q * 3]) === 'O'); p.wrong.push(num_(r[10 + q * 3])); p.hinted.push(num_(r[11 + q * 3])); }
      list.push(p);
    }
  }
  return { sh: sh, list: list };
}
function pRow_(p) {
  var r = [p.key, p.team, p.name, p.ch, p.coins, p.hasKey ? 'O' : '', p.helping || '', p.joined, nowKST_()];
  for (var q = 0; q < NQ; q++) { r.push(p.done[q] ? 'O' : ''); r.push(p.wrong[q] || 0); r.push(p.hinted[q] || 0); }
  return r;
}
function saveP_(sh, p) { var r = pRow_(p); sh.getRange(p.row, 1, 1, r.length).setValues([r]); }
function findP_(list, key) { for (var i = 0; i < list.length; i++) if (list[i].key === key) return list[i]; return null; }
function newP_(team, name) {
  var p = { key: key_(team, name), team: team, name: clean_(name, 10), ch: '', coins: 0, hasKey: false, helping: 0, joined: nowKST_(), done: [], wrong: [], hinted: [] };
  for (var q = 0; q < NQ; q++) { p.done.push(false); p.wrong.push(0); p.hinted.push(0); }
  return p;
}
function readH_() {
  var sh = sheet_(SH_H, H_H), last = sh.getLastRow(), list = [];
  if (last >= 2) {
    var v = sh.getRange(2, 1, last - 1, H_H.length).getValues();
    for (var i = 0; i < v.length; i++) {
      var r = v[i];
      list.push({ row: i + 2, id: num_(r[0]), t: String(r[1]), team: num_(r[2]), from: String(r[3]), to: String(r[4]), q: num_(r[5]), kind: String(r[6]), text: String(r[7]), thanks: String(r[8]) === 'O' });
    }
  }
  return { sh: sh, list: list };
}
function readR_() {
  var sh = sheet_(SH_R, R_H), last = sh.getLastRow(), list = [];
  if (last >= 2) {
    var v = sh.getRange(2, 1, last - 1, R_H.length).getValues();
    for (var i = 0; i < v.length; i++) { var r = v[i]; list.push({ row: i + 2, id: num_(r[0]), t: String(r[1]), team: num_(r[2]), name: String(r[3]), q: num_(r[4]), kind: String(r[5]), status: String(r[6]) }); }
  }
  return { sh: sh, list: list };
}
function readB_() {
  var sh = sheet_(SH_B, B_H), last = sh.getLastRow(), list = [];
  if (last >= 2) {
    var v = sh.getRange(2, 1, last - 1, B_H.length).getValues();
    for (var i = 0; i < v.length; i++) {
      var r = v[i];
      list.push({ row: i + 2, id: String(r[0]), key: String(r[1]), q: num_(r[2]), codes: [r[3], r[4], r[5], r[6], r[7]].map(String), pick: String(r[8]), gained: Number(r[9]) || 0 });
    }
  }
  return { sh: sh, list: list };
}
function readT_() {
  var sh = sheet_(SH_T, T_H), last = sh.getLastRow(), map = {};
  if (last >= 2) {
    var v = sh.getRange(2, 1, last - 1, 4).getValues();
    for (var i = 0; i < v.length; i++) { var t = num_(v[i][0]); if (t) map[t] = { row: i + 2, label: String(v[i][1] || ''), esc: Number(v[i][2]) || 0 }; }
  }
  return { sh: sh, map: map };
}
function upsertT_(T, t, label, esc) {
  var cur = T.map[t];
  var lab = label != null ? label : (cur ? cur.label : '');
  var e = esc != null ? esc : (cur ? cur.esc : 0);
  var row = [t, lab, e || '', e ? Utilities.formatDate(new Date(e), 'Asia/Seoul', 'HH:mm:ss') : ''];
  if (cur) { T.sh.getRange(cur.row, 1, 1, 4).setValues([row]); cur.label = lab; cur.esc = e; }
  else { T.sh.appendRow(row); T.map[t] = { row: T.sh.getLastRow(), label: lab, esc: e }; }
}
function labelOf_(T, t) { return (T[t] && T[t].label) || ISLAND_DEFAULT[t] || (t + '조'); }

/* 읽기용 스냅샷 (4초 캐시: 학생 20명이 5초마다 불러도 시트를 덜 읽게) */
function snap_() {
  var c = CacheService.getScriptCache(), hit = c.get('snap');
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  var T = readT_().map, tt = {};
  Object.keys(T).forEach(function (k) { tt[k] = { label: T[k].label, esc: T[k].esc }; });
  var pend = readB_().list.filter(function (b) { return b.pick === ''; }).map(function (b) { return { id: b.id, key: b.key, q: b.q }; });
  var s = { P: readP_().list, H: readH_().list, R: readR_().list.slice(-60), T: tt, B: pend };
  try { c.put('snap', JSON.stringify(s), 4); } catch (e) {}
  return s;
}

/* ===== 보드 ===== */
function board_(s) {
  var tm = {};
  s.P.forEach(function (pl) {
    var t = tm[pl.team] || (tm[pl.team] = { team: pl.team, n: 0, coins: 0, solved: 0 });
    t.n++; t.coins += pl.coins;
    pl.done.forEach(function (d) { if (d) t.solved++; });
  });
  return Object.keys(tm).map(function (k) {
    var t = tm[k], info = s.T[t.team] || {};
    return { team: t.team, label: labelOf_(s.T, t.team), n: t.n, avg: Math.round(t.coins / t.n * 10) / 10, progress: Math.round(t.solved / (t.n * NQ) * 100), escapedAt: info.esc || 0 };
  });
}
function pubP_(p) { return { name: p.name, ch: p.ch, coins: p.coins, hasKey: p.hasKey, helping: p.helping, done: p.done }; }

/* ===== 조회 ===== */
function lobby_() {
  var s = snap_(), cnt = {};
  s.P.forEach(function (p) { cnt[p.team] = (cnt[p.team] || 0) + 1; });
  return { ok: true, now: Date.now(), game: game_(), teams: TEAMS.map(function (t) { return { team: t, label: labelOf_(s.T, t), n: cnt[t] || 0, max: MAX_TEAM }; }) };
}
function state_(p) {
  var s = snap_(), team = num_(p.team), key = key_(team, p.name), me = null, members = [];
  s.P.forEach(function (x) { if (x.key === key) me = x; if (x.team === team) members.push(x); });
  if (!me) return { ok: false, gone: true, error: '명단에 없어요. 다시 입장해 주세요' };
  var info = s.T[team] || {};
  var recv = [], sent = [];
  s.H.forEach(function (h) {
    if (h.team !== team) return;
    if (h.to === me.name) recv.push({ id: h.id, t: h.t, from: h.from, q: h.q, kind: h.kind, text: h.text, thanks: h.thanks });
    if (h.from === me.name && h.kind === '조원') sent.push({ id: h.id, to: h.to, q: h.q, thanks: h.thanks });
  });
  return {
    ok: true, now: Date.now(), game: game_(), cfg: cfg_(),
    me: { key: me.key, name: me.name, team: team, ch: me.ch, coins: me.coins, hasKey: me.hasKey, helping: me.helping, done: me.done, wrong: me.wrong, hinted: me.hinted },
    team: { team: team, label: labelOf_(s.T, team), leader: members[0] ? members[0].name : '', escapedAt: info.esc || 0, members: members.map(pubP_) },
    hints: { recv: recv, sent: sent },
    boxes: s.B.filter(function (b) { return b.key === key; }),
    board: board_(s)
  };
}

/* ===== 입장·캐릭터 ===== */
function join_(p) {
  var team = num_(p.team), name = clean_(p.name, 10);
  if (TEAMS.indexOf(team) < 0) return { ok: false, error: '조를 다시 골라 주세요' };
  if (!name) return { ok: false, error: '이름을 입력해 주세요' };
  var R = readP_(), key = key_(team, name), ex = findP_(R.list, key), cnt = 0;
  R.list.forEach(function (x) { if (x.team === team) cnt++; });
  if (ex) {
    if (String(p.confirm) !== '1') return { ok: true, exists: true };
    return { ok: true, rejoin: true, ch: ex.ch };
  }
  if (cnt >= MAX_TEAM) return { ok: false, error: '이 섬은 가득 찼어요 (최대 ' + MAX_TEAM + '명). 다른 조를 골라 주세요' };
  R.sh.appendRow(pRow_(newP_(team, name)));
  return { ok: true, created: true, ch: '' };
}
function cancelBoxes_(key) {
  var B = readB_();
  B.list.forEach(function (b) { if (b.key === key && b.pick === '') B.sh.getRange(b.row, 9).setValue('취소'); });
}
function leave_(p) {
  var R = readP_(), key = key_(p.team, p.name), x = findP_(R.list, key);
  if (x) { R.sh.deleteRow(x.row); cancelBoxes_(key); }
  return { ok: true };
}
function setchar_(p) {
  if (CHARS.indexOf(String(p.char)) < 0) return { ok: false, error: '캐릭터 오류' };
  var R = readP_(), x = findP_(R.list, key_(p.team, p.name));
  if (!x) return { ok: false, gone: true, error: '다시 입장해 주세요' };
  x.ch = String(p.char); saveP_(R.sh, x);
  return { ok: true };
}
function setisland_(p) {
  var R = readP_(), team = num_(p.team), mem = R.list.filter(function (x) { return x.team === team; });
  var isAdmin = adminOk_(p.code);
  if (!isAdmin && !(mem[0] && mem[0].name === clean_(p.name, 10))) return { ok: false, error: '섬 이름은 조장(처음 입장한 사람)만 바꿀 수 있어요' };
  var label = clean_(p.label, 12);
  var T = readT_();
  upsertT_(T, team, label || ISLAND_DEFAULT[team] || '', null);
  return { ok: true, label: label };
}

/* ===== 풀이 결과 ===== */
function resolveReq_(team, name, q) {
  var Rq = readR_();
  Rq.list.forEach(function (r) { if (r.team === team && r.name === name && r.q === q && r.status === '대기') Rq.sh.getRange(r.row, 7).setValue('해결'); });
}
function checkEscape_(team) {
  var R = readP_(), mem = R.list.filter(function (x) { return x.team === team; }), T = readT_();
  if (T.map[team] && T.map[team].esc) return 0;
  if (mem.length < cfg_().minTeam) return 0;
  for (var i = 0; i < mem.length; i++) if (!mem[i].hasKey) return 0;
  var now = Date.now();
  upsertT_(T, team, null, now);
  return now;
}
function result_(p) {
  var g = game_();
  if (g.phase !== 'running') return { ok: false, error: g.phase === 'ready' ? '아직 시작 전이에요. 아듬샘의 시작 신호를 기다려 주세요' : '시간이 끝났어요 🌊' };
  var R = readP_(), team = num_(p.team), pl = findP_(R.list, key_(team, p.name));
  if (!pl) return { ok: false, gone: true, error: '다시 입장해 주세요' };
  var q = num_(p.q);
  if (q < 1 || q > NQ) return { ok: false, error: '문항 오류' };
  if (pl.done[q - 1]) return { ok: true, already: true, coins: pl.coins };
  if (q > 1 && !pl.done[q - 2]) return { ok: false, error: '앞 문제를 먼저 풀어야 해요' };
  var correct = String(p.correct) === '1';
  var lg = sheet_(SH_LOG, LOG_H);
  if (!correct) {
    pl.wrong[q - 1]++;
    saveP_(R.sh, pl);
    lg.appendRow([nowKST_(), team, pl.name, q, '오답', 0]);
    return { ok: true, correct: false, wrong: pl.wrong[q - 1] };
  }
  var first = pl.wrong[q - 1] === 0 && pl.hinted[q - 1] === 0;
  var gained = Math.max(COIN_MIN, COIN_BASE + (first ? COIN_FIRST : 0) - COIN_HINT_PENALTY * pl.hinted[q - 1]);
  pl.done[q - 1] = true; pl.coins += gained;
  if (pl.helping === q) pl.helping = 0;
  if (q === NQ) pl.hasKey = true;
  saveP_(R.sh, pl);
  lg.appendRow([nowKST_(), team, pl.name, q, '정답', gained]);
  resolveReq_(team, pl.name, q);
  var boxId = pl.key + '_' + q, codes = shuffle_(BOX_CODES.slice()), B = readB_(), dup = false;
  B.list.forEach(function (b) { if (b.id === boxId) dup = true; });
  if (!dup) B.sh.appendRow([boxId, pl.key, q, codes[0], codes[1], codes[2], codes[3], codes[4], '', '', nowKST_()]);
  var esc = (q === NQ) ? checkEscape_(team) : 0;
  return { ok: true, correct: true, gained: gained, coins: pl.coins, boxId: dup ? '' : boxId, key: q === NQ, escaped: esc > 0 };
}

/* ===== 도움 요청 · 힌트 ===== */
function help_(p) {
  var R = readP_(), team = num_(p.team), me = findP_(R.list, key_(team, p.name));
  if (!me) return { ok: false, gone: true, error: '다시 입장해 주세요' };
  var q = num_(p.q);
  if (q < 1 || q > NQ) return { ok: false, error: '문항 오류' };
  if (me.done[q - 1]) return { ok: false, error: '이미 푼 문제예요' };
  var solvers = R.list.filter(function (x) { return x.team === team && x.name !== me.name && x.done[q - 1]; }).length;
  var Rq = readR_();
  if (String(p.mode) === 'team') {
    if (!solvers) return { ok: false, code: 'nosolver' };
    me.helping = q; saveP_(R.sh, me);
    Rq.sh.appendRow([nextId_(Rq.sh), nowKST_(), team, me.name, q, '조원', '대기']);
    return { ok: true };
  }
  // 담임 힌트
  var cf = cfg_();
  if (solvers && !cf.always) return { ok: false, code: 'blocked', solvers: solvers };
  if (solvers && String(p.force) !== '1') return { ok: false, code: 'confirm', solvers: solvers };
  var text = clean_(p.text, 200);
  if (!text) return { ok: false, error: '힌트 내용이 없어요' };
  var H = readH_();
  H.sh.appendRow([nextId_(H.sh), nowKST_(), team, '담임', me.name, q, '담임', text, '']);
  me.hinted[q - 1]++; saveP_(R.sh, me);
  Rq.sh.appendRow([nextId_(Rq.sh), nowKST_(), team, me.name, q, '담임', '처리']);
  return { ok: true };
}
function hint_(p) {
  var R = readP_(), team = num_(p.team), me = findP_(R.list, key_(team, p.name)), to = findP_(R.list, key_(team, p.to));
  if (!me || !to) return { ok: false, error: '조원을 찾을 수 없어요' };
  var q = num_(p.q);
  if (q < 1 || q > NQ) return { ok: false, error: '문항 오류' };
  if (!me.done[q - 1]) return { ok: false, error: '이 문제를 푼 사람만 힌트를 보낼 수 있어요' };
  var text = clean_(p.text, 200);
  if (text.length < 4) return { ok: false, error: '힌트를 조금 더 자세히 적어 주세요' };
  var H = readH_(), already = H.list.some(function (h) { return h.team === team && h.from === me.name && h.to === to.name && h.q === q && h.kind === '조원'; });
  H.sh.appendRow([nextId_(H.sh), nowKST_(), team, me.name, to.name, q, '조원', text, '']);
  to.hinted[q - 1]++; saveP_(R.sh, to);
  if (!already) { me.coins += COIN_HELP; saveP_(R.sh, me); }
  return { ok: true, coins: me.coins, bonus: !already, help: COIN_HELP };
}
function thanks_(p) {
  var H = readH_(), id = num_(p.id), team = num_(p.team);
  for (var i = 0; i < H.list.length; i++) {
    var h = H.list[i];
    if (h.id === id && h.team === team && h.to === clean_(p.name, 10)) { H.sh.getRange(h.row, 9).setValue('O'); return { ok: true }; }
  }
  return { ok: false, error: '힌트를 찾을 수 없어요' };
}

/* ===== 보석함 ===== */
function openbox_(p) {
  var id = String(p.id || ''), pick = parseInt(p.pick, 10);
  if (!id || !(pick >= 1 && pick <= 5)) return { ok: false, error: '상자 선택 오류' };
  var B = readB_(), R = readP_();
  for (var i = 0; i < B.list.length; i++) {
    var b = B.list[i];
    if (b.id !== id) continue;
    var me = findP_(R.list, b.key);
    if (b.pick !== '') return { ok: true, already: true, codes: b.codes, pick: num_(b.pick), gained: b.gained, coins: me ? me.coins : 0 };
    var code = b.codes[pick - 1], gained = 0;
    if (code === 'D') gained = Math.min(DOUBLE_CAP, me ? me.coins : 0);
    else if (code !== 'X') gained = parseInt(code, 10) || 0;
    if (me && gained) { me.coins += gained; saveP_(R.sh, me); }
    B.sh.getRange(b.row, 9, 1, 3).setValues([[pick, gained, nowKST_()]]);
    return { ok: true, already: false, codes: b.codes, pick: pick, gained: gained, coins: me ? me.coins : 0 };
  }
  return { ok: false, error: '상자를 찾을 수 없어요' };
}

/* ===== 관리자 ===== */
function needAdmin_(p) { return adminOk_(p.code) ? null : { ok: false, error: '입장코드가 달라요' }; }
function adminstate_(p) {
  var e = needAdmin_(p); if (e) return e;
  var s = snap_();
  return {
    ok: true, now: Date.now(), game: game_(), cfg: cfg_(),
    players: s.P.map(function (x) { return { team: x.team, name: x.name, ch: x.ch, coins: x.coins, hasKey: x.hasKey, helping: x.helping, done: x.done, wrong: x.wrong, hinted: x.hinted }; }),
    labels: TEAMS.map(function (t) { return { team: t, label: labelOf_(s.T, t), escapedAt: (s.T[t] && s.T[t].esc) || 0 }; }),
    board: board_(s),
    reqs: s.R.filter(function (r) { return r.kind === '담임'; }).slice(-20).reverse()
  };
}
function admincfg_(p) {
  var e = needAdmin_(p); if (e) return e;
  var pr = props_(), now = Date.now(), g = game_();
  if (String(p.start) === '1') { pr.setProperty('START', String(now)); }
  if (String(p.unstart) === '1') { pr.deleteProperty('START'); }
  if (String(p.stop) === '1' && g.startAt) { pr.setProperty('DUR', String(Math.max(0.1, (now - g.startAt) / 60000))); }
  if (p.dur !== undefined && p.dur !== '') { var d = Number(p.dur); if (d > 0 && d <= 600) pr.setProperty('DUR', String(d)); }
  if (p.delta !== undefined && p.delta !== '') {
    var dl = Number(p.delta) || 0;
    if (g.startAt) {
      var base = g.phase === 'over' ? (now - g.startAt) / 60000 : g.durationMin;
      pr.setProperty('DUR', String(Math.max(0.5, base + dl)));
    } else pr.setProperty('DUR', String(Math.max(1, g.durationMin + dl)));
  }
  if (p.always !== undefined && p.always !== '') pr.setProperty('ALWAYS', String(p.always) === '1' ? '1' : '0');
  if (p.minteam !== undefined && p.minteam !== '') pr.setProperty('MINTEAM', String(Math.max(1, Math.min(MAX_TEAM, num_(p.minteam) || 2))));
  return { ok: true, game: game_(), cfg: cfg_() };
}
function adminmove_(p) {
  var e = needAdmin_(p); if (e) return e;
  var R = readP_(), team = num_(p.team), to = num_(p.to), x = findP_(R.list, key_(team, p.name));
  if (!x) return { ok: false, error: '대상을 찾을 수 없어요' };
  if (TEAMS.indexOf(to) < 0) return { ok: false, error: '조 오류' };
  if (findP_(R.list, key_(to, x.name))) return { ok: false, error: '그 조에 같은 이름이 있어요' };
  var cnt = R.list.filter(function (y) { return y.team === to; }).length;
  if (cnt >= MAX_TEAM) return { ok: false, error: '그 섬은 가득 찼어요' };
  var oldKey = x.key;
  x.team = to; x.key = key_(to, x.name);
  saveP_(R.sh, x);
  var B = readB_();
  B.list.forEach(function (b) { if (b.key === oldKey) B.sh.getRange(b.row, 2).setValue(x.key); });
  return { ok: true };
}
function admindel_(p) {
  var e = needAdmin_(p); if (e) return e;
  return leave_(p);
}
function adminhint_(p) {
  var e = needAdmin_(p); if (e) return e;
  var text = clean_(p.text, 200);
  if (!text) return { ok: false, error: '내용을 적어 주세요' };
  var R = readP_(), q = num_(p.q), team = num_(p.team), targets = [];
  R.list.forEach(function (x) {
    if (String(p.scope) === 'all') targets.push(x);
    else if (String(p.scope) === 'team' && x.team === team) targets.push(x);
    else if (String(p.scope) === 'one' && x.team === team && x.name === clean_(p.name, 10)) targets.push(x);
  });
  if (!targets.length) return { ok: false, error: '받을 사람이 없어요' };
  var H = readH_(), id = nextId_(H.sh), now = nowKST_();
  var rows = targets.map(function (x, i) { return [id + i, now, x.team, '담임', x.name, q || 0, '담임', text, '']; });
  H.sh.getRange(H.sh.getLastRow() + 1, 1, rows.length, H_H.length).setValues(rows);
  return { ok: true, sent: targets.length };
}
function clearRows_(name, headers) {
  var sh = sheet_(name, headers), last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, 1, last - 1, headers.length).clearContent();
}
function adminreset_(p) {
  var e = needAdmin_(p); if (e) return e;
  clearRows_(SH_P, pHeaders_()); clearRows_(SH_LOG, LOG_H); clearRows_(SH_H, H_H); clearRows_(SH_R, R_H); clearRows_(SH_B, B_H); clearRows_(SH_T, T_H);
  props_().deleteProperty('START');
  return { ok: true };
}

/* ===== 진입점 ===== */
var WRITERS = { join: join_, leave: leave_, setchar: setchar_, setisland: setisland_, result: result_, help: help_, hint: hint_, thanks: thanks_, openbox: openbox_, admincfg: admincfg_, adminmove: adminmove_, admindel: admindel_, adminhint: adminhint_, adminreset: adminreset_ };
function route_(p) {
  var a = p.action;
  if (WRITERS[a]) { var r = withLock_(function () { return WRITERS[a](p); }); bust_(); return r; }
  if (a === 'state') return state_(p);
  if (a === 'lobby') return lobby_();
  if (a === 'adminlogin') return adminOk_(p.code) ? { ok: true } : { ok: false, error: '입장코드가 달라요' };
  if (a === 'adminstate') return adminstate_(p);
  return { ok: true, ping: 1 };
}
function doGet(e) {
  var p = (e && e.parameter) || {}, out;
  try { out = route_(p); } catch (err) { out = { ok: false, error: String(err) }; }
  return jsonp_(p.callback, out);
}
