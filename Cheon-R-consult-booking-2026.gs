// ======================================================
// 학부모 상담 예약 시스템 - Google Apps Script (보안·마감 강화판)
// Made by Cheon_R · GAS가 모든 데이터의 단일 진실 공급원
// ------------------------------------------------------
// [처음 설치] 코드 붙여넣기 → 저장 → setupTrigger() 1회 실행 → 웹앱 배포
// [재배포]    반드시 "배포 관리 → 수정(연필)" ← "새 배포" 하면 URL이 바뀌어요!
// [비번 분실] 편집기에서 resetPassword() 실행 → 초기 비밀번호(consultbooking)로 복구
// ======================================================

const GAS_VER = 8;                 // 이 파일의 버전 (HTML의 NEED_GAS와 비교)
const PW_MIN  = 6;                 // 새 비밀번호 최소 길이
const NOTICE_MAX = 500;            // 공지 최대 글자 수 (요청이 주소창으로 가므로 너무 길면 구글이 거절해요)
const SHEET_BOOKING  = '예약';
const SHEET_SLOTS    = '슬롯';
const SHEET_SETTINGS = '설정';
const TZ             = 'Asia/Seoul';
const INIT_PW        = 'consultbooking';  // 초기 비밀번호 (첫 로그인 때 변경 강제)
const DAY_KO         = ['일','월','화','수','목','금','토'];

// ── 출력 헬퍼 (JSONP 지원 · callback 이름 검증) ───────────
function corsOutput(data, callback) {
  const json = JSON.stringify(data);
  const okCb = callback && /^[A-Za-z0-9_.$]{1,80}$/.test(callback);
  const body = okCb ? callback + '(' + json + ')' : json;
  const mime = okCb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON;
  return ContentService.createTextOutput(body).setMimeType(mime);
}

// ── POST 진입점 (기존 유지) ────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    return routeAction(ss, data, null);
  } catch (err) {
    return corsOutput({ result: 'error', message: err.message }, null);
  }
}

// ── GET 진입점 (JSONP + jsonBody 지원) ─────────────────
function doGet(e) {
  const callback = (e.parameter && e.parameter.callback) || null;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let data;
    if (e.parameter.jsonBody) {
      // e.parameter는 이미 1회 디코딩된 값 → 그대로 JSON 파싱 (실패 시에만 구버전 호환 디코딩)
      try { data = JSON.parse(e.parameter.jsonBody); }
      catch (x) { data = JSON.parse(decodeURIComponent(e.parameter.jsonBody)); }
    } else {
      data = {};
      Object.keys(e.parameter).forEach(k => { data[k] = e.parameter[k]; });
    }
    return routeAction(ss, data, callback);
  } catch (err) {
    return corsOutput({ result: 'error', message: err.message }, callback);
  }
}

// ── 액션 라우터 ────────────────────────────────────────
// 학부모용(비밀번호 불필요): getAll / getSlots / getSettings / checkDuplicate / checkBooking / saveBooking
// 관리자용(비밀번호 필수):   그 외 전부
function routeAction(ss, data, callback) {
  let result;
  // 주소만 열었을 때(설치 확인용) — 관리자 인증을 거치지 않으므로 잠금 횟수도 올라가지 않음
  if (!data.action) return corsOutput({ result: 'ok', app: 'consult-booking', gasVer: GAS_VER, message: '✅ 상담 예약 웹앱이 정상 작동 중이에요.' }, callback);
  try {
    switch (data.action) {
      // ─ 학부모용 ─
      case 'getAll':         result = getAll(ss); break;
      case 'getSlots':       result = { result: 'success', data: loadSlots(ss), serverNow: nowKst() }; break;
      case 'getSettings':    result = { result: 'success', data: publicSettings(ss) }; break;
      case 'checkDuplicate': result = checkDuplicate(ss, data); break;
      case 'checkBooking':   result = checkBooking(ss, data); break;
      case 'saveBooking':    result = saveBooking(ss, data); break;
      // ─ 관리자용 ─
      case 'adminLogin':     result = adminLogin(ss, data); break;
      default:
        result = adminRoute(ss, data);
    }
  } catch (err) {
    result = { result: 'error', message: err.message };
  }
  return corsOutput(result, callback);
}

function adminRoute(ss, data) {
  const auth = verifyAdmin(ss, data.pw);
  if (auth !== 'ok') return { result: auth };           // 'denied' | 'locked'
  // 초기 비밀번호 상태(설치 직후 · resetPassword 직후)에는 비밀번호 변경 말고는 아무것도 못 하게 서버에서 막음
  if (data.action !== 'changePassword' && String(readSettings(ss).pwChanged).toLowerCase() === 'false') return { result: 'mustChange' };
  switch (data.action) {
    case 'getAdminAll':     return getAdminAll(ss);
    case 'getBookings':     return { result: 'success', data: readBookings(ss) };
    case 'saveSettings':    return saveSettings(ss, data);
    case 'changePassword':  return changePassword(ss, data);
    case 'cancelBooking':   return withLock(() => cancelBooking(ss, data));
    case 'confirmBooking':  return withLock(() => confirmBooking(ss, data));
    case 'confirmAll':      return withLock(() => confirmAll(ss));
    case 'unconfirmBooking': return withLock(() => unconfirmBooking(ss, data));
    case 'addDay':          return withLock(() => addDay(ss, data));
    case 'deleteDay':       return withLock(() => deleteDay(ss, data));
    case 'addSlot':         return withLock(() => addSlot(ss, data));
    case 'addSlots':        return withLock(() => addSlots(ss, data));
    case 'deleteSlot':      return withLock(() => deleteSlot(ss, data));
    case 'clearBookings':   return withLock(() => clearBookings(ss, false));
    case 'clearAll':        return withLock(() => clearBookings(ss, true));
    default:                return { result: 'error', message: 'unknown action: ' + data.action };
  }
}

// ══════════════════════════════════════════════════════
// ★ onChange 트리거 - 시트에서 행 삭제 시 슬롯 자동 동기화
// ══════════════════════════════════════════════════════
// [설치] 편집기에서 setupTrigger 함수를 선택 → ▶ 실행 (1회) — 수동 설정 불필요
// ──────────────────────────────────────────────────────
function setupTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onBookingEdit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onBookingEdit').forSpreadsheet(ss).onChange().create();
  ensureBookingSheet(ss); ensureSlotsSheet(ss); ensureSettingsSheet(ss);
  console.log('✅ 트리거 설치 완료');
}

function onBookingEdit(e) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(8000)) return;
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_BOOKING);
    if (!sheet) return;

    if (e && e.changeType) {
      const allowedTypes = ['REMOVE_ROW', 'INSERT_ROW', 'OTHER', 'EDIT'];
      if (!allowedTypes.includes(e.changeType)) return;
    }

    const rows = sheet.getDataRange().getValues();
    const activeSlotIds = new Set(
      rows.slice(1).map(r => String(r[6]).trim()).filter(id => id !== '')
    );

    const slots = loadSlots(ss);
    let changed = false;
    slots.forEach(day => {
      day.slots.forEach(slot => {
        const shouldBeBooked = activeSlotIds.has(String(slot.id));
        if (slot.booked !== shouldBeBooked) { slot.booked = shouldBeBooked; changed = true; }
      });
    });
    if (changed) storeSlots(ss, slots);
  } catch (err) {
    console.error('onBookingEdit error:', err.message);
  } finally {
    lock.releaseLock();
  }
}

// ══════════════════════════════════════════════════════
// 인증
// ══════════════════════════════════════════════════════
function sha256Hex(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

// 'ok' | 'denied' | 'locked'
function verifyAdmin(ss, pw) {
  const cache = CacheService.getScriptCache();
  const fails = parseInt(cache.get('pwFail') || '0', 10);
  if (fails >= 8) return 'locked';
  pw = String(pw === undefined || pw === null ? '' : pw);
  if (pw === '') return 'denied';                       // 비밀번호 없는 요청은 실패 횟수에 넣지 않음 (아무나 잠그지 못하게)
  const S = readSettings(ss);
  let ok = false;
  if (S.pwHash) {
    ok = pw !== '' && sha256Hex(String(S.pwSalt) + pw) === S.pwHash;
  } else if (S.password !== undefined && S.password !== '') {
    ok = pw !== '' && pw === String(S.password);        // 구버전 시트(평문) 호환
  }
  if (ok) { cache.remove('pwFail'); return 'ok'; }
  cache.put('pwFail', String(fails + 1), 600);          // 10분간 8회 실패 시 잠금
  return 'denied';
}

function adminLogin(ss, data) {
  const auth = verifyAdmin(ss, data.pw);
  if (auth !== 'ok') return { result: auth };
  const lock = LockService.getScriptLock();
  if (lock.tryLock(8000)) {
    try {
      const S = readSettings(ss);
      if (!S.pwHash) setPassword(ss, String(data.pw), S.pwChanged === undefined ? 'true' : null);  // 평문 → 해시 이전
      if (!String(S.installId || '').trim()) setSetting(ss, 'installId', Utilities.getUuid());       // 설치 ID (사용자 등록부용 · 시트마다 한 번만 만들어요)
    } finally { lock.releaseLock(); }
  }
  const S2 = readSettings(ss);
  return {
    result:     'success',
    mustChange: String(S2.pwChanged).toLowerCase() === 'false',
    gasVer:     GAS_VER,
    installId:  String(S2.installId || '').trim(),
    settings:   adminSettings(ss),
  };
}

function setPassword(ss, newPw, changedFlag) {
  const salt = Utilities.getUuid();
  setSetting(ss, 'pwSalt', salt);
  setSetting(ss, 'pwHash', sha256Hex(salt + newPw));
  removeSetting(ss, 'password');                        // 평문 키 제거
  if (changedFlag !== null && changedFlag !== undefined) setSetting(ss, 'pwChanged', changedFlag);
}

function changePassword(ss, data) {
  const nw = String(data.newPw || '');
  if (nw.length < PW_MIN)       return { result: 'invalid', message: '새 비밀번호는 ' + PW_MIN + '자 이상이어야 해요.' };
  if (nw === INIT_PW)           return { result: 'invalid', message: '초기 비밀번호와 다른 비밀번호로 정해 주세요.' };
  if (nw === String(data.pw))   return { result: 'invalid', message: '현재 비밀번호와 다른 비밀번호로 정해 주세요.' };
  return withLock(() => { setPassword(ss, nw, 'true'); return { result: 'success' }; });
}

// 비밀번호 분실 시 편집기에서 직접 실행 (초기 비밀번호로 복구 + 다음 로그인 때 변경 강제)
function resetPassword() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSettingsSheet(ss);
  setPassword(ss, INIT_PW, 'false');
  CacheService.getScriptCache().remove('pwFail');
  console.log('✅ 비밀번호가 초기값(' + INIT_PW + ')으로 복구됐어요. 로그인 후 바로 변경해 주세요.');
}

function withLock(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { result: 'busy' };
  try { return fn(); } finally { lock.releaseLock(); }
}

// ══════════════════════════════════════════════════════
// 시간 (KST)
// ══════════════════════════════════════════════════════
// 상담 시간이 이미 끝났는지 (종료 시각 기준). time 형식: "HH:MM ~ HH:MM"
function slotEnded_(dateKey, time) {
  const m = String(time).match(/~\s*(\d{2}:\d{2})/);
  return m ? (String(dateKey) + ' ' + m[1]) <= nowKst() : false;
}
function nowKst() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }

function cellStr(v, withTime) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, withTime ? 'yyyy-MM-dd HH:mm:ss' : 'yyyy-MM-dd');
  return v === undefined || v === null ? '' : String(v);
}

// ══════════════════════════════════════════════════════
// 설정
// ══════════════════════════════════════════════════════
function readSettings(ss) {
  ensureSettingsSheet(ss);
  const rows = ss.getSheetByName(SHEET_SETTINGS).getDataRange().getValues();
  const map = {};
  rows.forEach(r => { if (r[0] !== '' && r[0] !== undefined) map[String(r[0])] = r[1]; });
  return map;
}

function truthy(v, dflt) {
  if (v === undefined || v === null || v === '') return dflt;
  return String(v).trim().toLowerCase() !== 'false';    // 대소문자 무시 (시트가 FALSE로 바꿔도 OK)
}

// 마감 판단의 단일 기준: 수동 마감 + 자동 마감 시각
function isOpen(S) {
  if (!truthy(S.bookingEnabled, true)) return false;
  const openAt = cellStr(S.openAt, true).replace('T', ' ').trim();          // 자동 오픈 시각 전이면 아직 안 열림
  if (openAt && nowKst() < openAt.slice(0, 16)) return false;
  const closeAt = cellStr(S.closeAt, true).replace('T', ' ').trim();
  if (closeAt && nowKst() >= closeAt.slice(0, 16)) return false;
  return true;
}

function publicSettings(ss) {
  const S = readSettings(ss);
  return {
    teacherName:    cellStr(S.teacherName),
    teacherRole:    cellStr(S.teacherRole),
    phoneEnabled:   truthy(S.phoneEnabled, true),
    bookingEnabled: isOpen(S),                           // 수동+자동 마감 모두 반영된 최종 상태
    notice:         cellStr(S.notice),
    closeAt:        cellStr(S.closeAt, true).replace(' ', 'T').slice(0, 16),
    openAt:         cellStr(S.openAt, true).replace(' ', 'T').slice(0, 16),   // 자동 오픈 시각 (비어 있으면 바로 열림)
    confirmMode:    truthy(S.confirmMode, false),         // 선생님 확인 후 확정 모드 (기본 꺼짐)
    notYetOpen:     truthy(S.bookingEnabled, true) && !!cellStr(S.openAt, true).trim() && nowKst() < cellStr(S.openAt, true).replace('T', ' ').trim().slice(0, 16),   // 스위치는 켜져 있는데 오픈 시각 전
  };
}

// 관리자용: 수동 스위치 값을 그대로 + 자동 마감 상태 분리 제공
function adminSettings(ss) {
  const S = readSettings(ss);
  const pub = publicSettings(ss);
  pub.bookingEnabled = truthy(S.bookingEnabled, true);   // 수동 스위치 원본
  pub.effectiveOpen  = isOpen(S);
  return pub;
}

function saveSettings(ss, data) {
  return withLock(() => {
    const rules = {
      teacherName: v => String(v).slice(0, 40),
      teacherRole: v => String(v).slice(0, 60),
      notice:      v => String(v).slice(0, NOTICE_MAX),
      phoneEnabled:   v => (String(v) === 'true' || String(v) === 'false') ? String(v) : null,
      bookingEnabled: v => (String(v) === 'true' || String(v) === 'false') ? String(v) : null,
      closeAt:     v => { v = String(v); return v === '' ? '' : (/^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}$/.test(v) ? v.replace(' ', 'T') : null); },
      confirmMode: v => (String(v) === 'true' || String(v) === 'false') ? String(v) : null,
      openAt:      v => { v = String(v); return v === '' ? '' : (/^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}$/.test(v) ? v.replace(' ', 'T') : null); },
    };
    Object.keys(rules).forEach(key => {
      if (data[key] === undefined || data[key] === null) return;
      const val = rules[key](data[key]);
      if (val === null) return;
      setSetting(ss, key, val);
    });
    return { result: 'success', settings: adminSettings(ss) };
  });
}

function setSetting(ss, key, val) {
  ensureSettingsSheet(ss);
  const sheet = ss.getSheetByName(SHEET_SETTINGS);
  const rows  = sheet.getDataRange().getValues();
  let row = 0;
  rows.forEach((r, i) => { if (String(r[0]) === key) row = i + 1; });
  if (!row) { row = sheet.getLastRow() + 1; sheet.getRange(row, 1).setValue(key); }
  const cell = sheet.getRange(row, 2);
  cell.setNumberFormat('@STRING@');                      // 'false'/날짜가 자동 변환되지 않도록 텍스트 고정
  cell.setValue(String(val));
}

function removeSetting(ss, key) {
  const sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet) return;
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 0; i--) {
    if (String(rows[i][0]) === key) sheet.deleteRow(i + 1);
  }
}

// ══════════════════════════════════════════════════════
// 슬롯 (JSON 1셀) — 항상 LockService 안에서만 수정
// ══════════════════════════════════════════════════════
function loadSlots(ss) {
  ensureSlotsSheet(ss);
  const val = ss.getSheetByName(SHEET_SLOTS).getRange('A1').getValue();
  let data = [];
  try { if (val) data = JSON.parse(val); } catch (e) {}
  return Array.isArray(data) ? data : [];
}

function storeSlots(ss, slots) {
  ensureSlotsSheet(ss);
  ss.getSheetByName(SHEET_SLOTS).getRange('A1').setValue(JSON.stringify(slots));
}

function findSlot(slots, slotId) {
  for (let i = 0; i < slots.length; i++) {
    const si = slots[i].slots.findIndex(s => String(s.id) === String(slotId));
    if (si >= 0) return { day: slots[i], slot: slots[i].slots[si] };
  }
  return null;
}

function addDay(ss, data) {
  const key = String(data.dateKey || '');
  const m = key.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return { result: 'invalid' };
  const slots = loadSlots(ss);
  if (slots.some(d => d.dateKey === key)) return { result: 'exists' };
  const dt = new Date(+m[1], +m[2] - 1, +m[3]);
  slots.push({ dateKey: key, label: (+m[2]) + '월 ' + (+m[3]) + '일 (' + DAY_KO[dt.getDay()] + ')', slots: [] });
  slots.sort((a, b) => a.dateKey.localeCompare(b.dateKey));
  storeSlots(ss, slots);
  return { result: 'success', slots };
}

function deleteDay(ss, data) {
  const slots = loadSlots(ss);
  const i = slots.findIndex(d => d.dateKey === String(data.dateKey));
  if (i < 0) return { result: 'notfound', slots };
  if (slots[i].slots.some(s => s.booked && !slotEnded_(slots[i].dateKey, s.time))) return { result: 'hasBooking', slots };   // 아직 안 지난 예약이 있는 날짜만 삭제 금지 (지난 예약 기록은 시트에 남음)
  slots.splice(i, 1);
  storeSlots(ss, slots);
  return { result: 'success', slots };
}

function addSlot(ss, data) {
  const start = String(data.start || ''), end = String(data.end || '');
  if (!/^\d{2}:\d{2}$/.test(start) || !/^\d{2}:\d{2}$/.test(end) || start >= end) return { result: 'invalid' };
  const slots = loadSlots(ss);
  const day = slots.find(d => d.dateKey === String(data.dateKey));
  if (!day) return { result: 'notfound', slots };
  const time = start + ' ~ ' + end;
  if (day.slots.some(s => s.time === time)) return { result: 'exists', slots };
  day.slots.push({ id: String(Date.now()) + String(Math.floor(Math.random() * 1000)), time, booked: false });
  day.slots.sort((a, b) => a.time.localeCompare(b.time));
  storeSlots(ss, slots);
  return { result: 'success', slots };
}

// 시간 여러 개를 한 번에 추가 (시간표 공강 → 상담 시간 만들기). 항목 형식: "YYYY-MM-DD HH:MM-HH:MM"
// 이미 있는 시간은 건너뛰므로 같은 요청을 다시 보내도 중복되지 않음
function addSlots(ss, data) {
  const items = Array.isArray(data.items) ? data.items : [];
  if (!items.length || items.length > 60) return { result: 'invalid' };
  const slots = loadSlots(ss);
  let added = 0, skipped = 0;
  items.forEach(it => {
    const m = String(it).match(/^(\d{4})-(\d{2})-(\d{2}) (\d{2}:\d{2})-(\d{2}:\d{2})$/);
    if (!m) { skipped++; return; }
    const y = +m[1], mo = +m[2], d = +m[3], start = m[4], end = m[5];
    const dt = new Date(y, mo - 1, d);
    if (dt.getMonth() !== mo - 1 || dt.getDate() !== d || start >= end) { skipped++; return; }
    const key = m[1] + '-' + m[2] + '-' + m[3];
    let day = slots.find(x => x.dateKey === key);
    if (!day) {
      day = { dateKey: key, label: mo + '월 ' + d + '일 (' + DAY_KO[dt.getDay()] + ')', slots: [] };
      slots.push(day);
    }
    const time = start + ' ~ ' + end;
    if (day.slots.some(s => s.time === time)) { skipped++; return; }
    day.slots.push({ id: Utilities.getUuid().replace(/-/g, '').slice(0, 12), time, booked: false });
    added++;
  });
  slots.sort((a, b) => a.dateKey.localeCompare(b.dateKey));
  slots.forEach(d => d.slots.sort((a, b) => a.time.localeCompare(b.time)));
  // 구글 시트 한 칸에는 5만 자까지만 들어가므로 넘기 전에 막음
  if (JSON.stringify(slots).length > 46000) return { result: 'tooMany', slots: loadSlots(ss) };
  storeSlots(ss, slots);
  return { result: 'success', added, skipped, slots };
}

function deleteSlot(ss, data) {
  const slots = loadSlots(ss);
  const f = findSlot(slots, data.slotId);
  if (!f) return { result: 'notfound', slots };
  if (f.slot.booked && !slotEnded_(f.day.dateKey, f.slot.time)) return { result: 'hasBooking', slots };   // 아직 안 지난 예약만 삭제 금지
  f.day.slots = f.day.slots.filter(s => String(s.id) !== String(data.slotId));
  storeSlots(ss, slots);
  return { result: 'success', slots };
}

function clearBookings(ss, alsoSlots) {
  ensureBookingSheet(ss);
  const sheet = ss.getSheetByName(SHEET_BOOKING);
  const last = sheet.getLastRow();
  if (last > 1) sheet.deleteRows(2, last - 1);
  const rl = ss.getSheetByName(SHEET_RELEASED); if (rl && rl.getLastRow() > 1) rl.deleteRows(2, rl.getLastRow() - 1);
  let slots = alsoSlots ? [] : loadSlots(ss);
  slots.forEach(d => d.slots.forEach(s => { s.booked = false; }));
  storeSlots(ss, slots);
  return { result: 'success', slots };
}

// ══════════════════════════════════════════════════════
// 조회
// ══════════════════════════════════════════════════════
// 학부모용: 설정(공개분) + 슬롯 + 서버시각. 예약자 정보는 절대 포함하지 않음
function getAll(ss) {
  return { result: 'success', settings: publicSettings(ss), slots: loadSlots(ss), serverNow: nowKst(), gasVer: GAS_VER };
}

function getAdminAll(ss) {
  return { result: 'success', settings: adminSettings(ss), slots: loadSlots(ss), bookings: readBookings(ss), serverNow: nowKst(), gasVer: GAS_VER };
}

function readBookings(ss) {
  const sheet = ss.getSheetByName(SHEET_BOOKING);
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];
  const headers = rows[0];
  return rows.slice(1).filter(r => r.some(c => c !== '')).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = (h === '신청일시') ? cellStr(row[i], true) : cellStr(row[i]); });
    return obj;
  });
}

function normName(s) { return String(s || '').replace(/\s+/g, ''); }
function last4(s)    { const d = String(s || '').replace(/\D/g, ''); return d.length >= 4 ? d.slice(-4) : ''; }

function findBookingRow(rows, student, l4) {
  const n = normName(student);
  if (!n || !l4) return null;
  for (let i = 1; i < rows.length; i++) {
    if (normName(rows[i][1]) === n && last4(rows[i][3]) === l4) return rows[i];
  }
  return null;
}

// 중복 확인(예약 제출 전 사전 확인): 학생 이름 + 연락처 뒤 4자리
function checkDuplicate(ss, data) {
  const sheet = ss.getSheetByName(SHEET_BOOKING);
  if (!sheet) return { result: 'success', exists: false };
  const r = findBookingRow(sheet.getDataRange().getValues(), data.student, last4(data.contact));
  if (!r) return { result: 'success', exists: false };
  return { result: 'success', exists: true, booking: { type: cellStr(r[0]), date: cellStr(r[4]), time: cellStr(r[5]) } };
}

// 내 예약 확인: 학생 이름 + 연락처 뒤 4자리 → 유형/날짜/시간만 반환 (학부모·연락처 등은 돌려주지 않음)
function checkBooking(ss, data) {
  const sheet = ss.getSheetByName(SHEET_BOOKING);
  const l4 = String(data.last4 || '').replace(/\D/g, '');
  if (!sheet || l4.length !== 4) return { result: 'success', found: false };
  const r = findBookingRow(sheet.getDataRange().getValues(), data.student, l4);
  if (!r) return { result: 'success', found: false, released: wasReleased(ss, data.student, l4) };
  return { result: 'success', found: true, status: String(r[10] || '').trim() === '신청' ? '신청' : '확정', booking: { type: cellStr(r[0]), date: cellStr(r[4]), time: cellStr(r[5]) } };
}

// ══════════════════════════════════════════════════════
// 예약 저장 / 취소
// ══════════════════════════════════════════════════════
function saveBooking(ss, data) {
  return withLock(() => {
    ensureBookingSheet(ss);
    const S = readSettings(ss);

    // 1) 마감 (수동 + 자동)
    if (!isOpen(S)) return { result: 'closed' };

    // 2) 입력 검증
    const type    = data.type === 'phone' ? 'phone' : 'face';
    if (type === 'phone' && !truthy(S.phoneEnabled, true)) return { result: 'phoneOff' };
    const student = String(data.student || '').trim().slice(0, 30);
    const parent  = String(data.parent  || '').trim().slice(0, 30);
    const memo    = String(data.memo    || '').trim().slice(0, 300);
    const digits  = String(data.contact || '').replace(/\D/g, '');
    if (!student || !parent || !memo || digits.length < 9 || digits.length > 11) return { result: 'invalid' };
    const contact = digits.length === 11 ? digits.replace(/(\d{3})(\d{4})(\d{4})/, '$1-$2-$3')
                  : digits.length === 10 ? digits.replace(/(\d{2,3})(\d{3,4})(\d{4})/, '$1-$2-$3')
                  : digits.replace(/(\d{2,3})(\d{3})(\d{4})/, '$1-$2-$3');

    // 3) 슬롯 확인 — 날짜/시간은 클라이언트가 아니라 서버의 슬롯 데이터에서 가져옴
    const slots = loadSlots(ss);
    const f = findSlot(slots, data.slotId);
    if (!f) return { result: 'gone' };                    // 그 사이 선생님이 시간을 삭제함
    const startM = String(f.slot.time).match(/^(\d{2}:\d{2})/);
    if (startM && (f.day.dateKey + ' ' + startM[1]) <= nowKst()) return { result: 'past' };   // 이미 지난 시간

    const sheet = ss.getSheetByName(SHEET_BOOKING);
    const rows  = sheet.getDataRange().getValues();

    // 4) 중복: 학생 이름 + 연락처 뒤 4자리
    if (findBookingRow(rows, student, last4(contact))) return { result: 'duplicate' };

    // 5) 슬롯 선점 (시트·JSON 양쪽 확인)
    const takenInSheet = rows.slice(1).some(r => String(r[6]).trim() === String(data.slotId).trim());
    if (takenInSheet || f.slot.booked) return { result: 'slotBooked' };

    // 6) 저장 (텍스트 서식으로 날짜·전화번호 자동변환 방지, 시각은 서버 KST)
    const row = sheet.getLastRow() + 1;
    ensureStatusHeader(sheet);
    const rng = sheet.getRange(row, 1, 1, 11);
    rng.setNumberFormat('@STRING@');
    rng.setValues([[
      type === 'phone' ? '전화' : '대면',
      student, parent, contact,
      f.day.label, f.slot.time,
      String(f.slot.id), f.day.dateKey,
      memo,
      Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
      truthy(S.confirmMode, false) ? '신청' : '확정',
    ]]);

    f.slot.booked = true;
    storeSlots(ss, slots);
    return { result: 'success', pending: truthy(S.confirmMode, false), booking: { type: type === 'phone' ? '전화' : '대면', date: f.day.label, time: f.slot.time } };
  });
}

function cancelBooking(ss, data) {
  const sheet = ss.getSheetByName(SHEET_BOOKING);
  if (!sheet) return { result: 'error', message: '예약 시트 없음' };
  const rows = sheet.getDataRange().getValues();
  let deleted = false;
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][6]).trim() === String(data.slotId).trim()) {
      logReleased(ss, rows[i][1], last4(rows[i][3]));
      sheet.deleteRow(i + 1);
      deleted = true;
      break;
    }
  }
  // 슬롯은 dateKey가 아니라 slotId만으로 찾아 해제 (날짜 서식 문제로 안 풀리던 버그 방지)
  const slots = loadSlots(ss);
  const f = findSlot(slots, data.slotId);
  if (f && f.slot.booked) { f.slot.booked = false; storeSlots(ss, slots); }
  return { result: deleted ? 'success' : 'notfound', slots };
}

// ══════════════════════════════════════════════════════
// v7: 신청 → 예약 확정 (상태 열 K) / 확정 안 된 학부모 안내용 기록
// 상태 열이 비어 있으면 '확정'으로 취급 (기존 예약 호환)
// ══════════════════════════════════════════════════════
const SHEET_RELEASED = '미확정기록';

function ensureStatusHeader(sheet) {
  if (String(sheet.getRange(1, 11).getValue()).trim() === '') {
    sheet.getRange(1, 11).setValue('상태').setFontWeight('bold').setBackground('#EDE9FE');
  }
}

function confirmBooking(ss, data) {
  const sheet = ss.getSheetByName(SHEET_BOOKING);
  if (!sheet) return { result: 'error', message: '예약 시트 없음' };
  ensureStatusHeader(sheet);
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][6]).trim() === String(data.slotId).trim()) {
      const c = sheet.getRange(i + 1, 11); c.setNumberFormat('@STRING@'); c.setValue('확정');
      return { result: 'success', bookings: readBookings(ss) };
    }
  }
  return { result: 'notfound' };
}

// 실수로 확정했을 때 '신청' 상태로 되돌리기
function unconfirmBooking(ss, data) {
  const sheet = ss.getSheetByName(SHEET_BOOKING);
  if (!sheet) return { result: 'error', message: '예약 시트 없음' };
  ensureStatusHeader(sheet);
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][6]).trim() === String(data.slotId).trim()) {
      const c = sheet.getRange(i + 1, 11); c.setNumberFormat('@STRING@'); c.setValue('신청');
      return { result: 'success', bookings: readBookings(ss) };
    }
  }
  return { result: 'notfound' };
}

function confirmAll(ss) {
  const sheet = ss.getSheetByName(SHEET_BOOKING);
  if (!sheet) return { result: 'success', count: 0, bookings: [] };
  ensureStatusHeader(sheet);
  const rows = sheet.getDataRange().getValues();
  let n = 0;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][10] || '').trim() === '신청') {
      const c = sheet.getRange(i + 1, 11); c.setNumberFormat('@STRING@'); c.setValue('확정'); n++;
    }
  }
  return { result: 'success', count: n, bookings: readBookings(ss) };
}

function logReleased(ss, student, l4) {
  try {
    let sh = ss.getSheetByName(SHEET_RELEASED);
    if (!sh) {
      sh = ss.insertSheet(SHEET_RELEASED);
      sh.appendRow(['학생이름', '연락처뒤4자리', '처리일시']);
      sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#EDE9FE');
    }
    const r = sh.getLastRow() + 1;
    const rng = sh.getRange(r, 1, 1, 3);
    rng.setNumberFormat('@STRING@');
    rng.setValues([[String(student), String(l4), Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss')]]);
  } catch (err) { console.error('logReleased:', err.message); }
}

function wasReleased(ss, student, l4) {
  const sh = ss.getSheetByName(SHEET_RELEASED);
  if (!sh) return false;
  const n = normName(student);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (normName(rows[i][0]) === n && String(rows[i][1]).trim() === String(l4)) return true;
  }
  return false;
}

// ══════════════════════════════════════════════════════
// 시트 준비
// ══════════════════════════════════════════════════════
function ensureBookingSheet(ss) {
  if (ss.getSheetByName(SHEET_BOOKING)) return;
  const sheet = ss.insertSheet(SHEET_BOOKING);
  sheet.appendRow(['유형','학생이름','학부모이름','연락처','날짜','시간','슬롯ID','날짜키','상담내용','신청일시','상태']);
  sheet.getRange(1,1,1,11).setFontWeight('bold').setBackground('#EDE9FE');
}

function ensureSlotsSheet(ss) {
  if (ss.getSheetByName(SHEET_SLOTS)) return;
  const sheet = ss.insertSheet(SHEET_SLOTS);
  sheet.getRange('A1').setValue('[]');
}

function ensureSettingsSheet(ss) {
  if (ss.getSheetByName(SHEET_SETTINGS)) return;
  const sheet = ss.insertSheet(SHEET_SETTINGS);
  const salt = Utilities.getUuid();
  const initData = [
    ['teacherName',    ''],
    ['teacherRole',    ''],
    ['phoneEnabled',   'true'],
    ['bookingEnabled', 'true'],
    ['notice',         ''],
    ['closeAt',        ''],
    ['pwSalt',         salt],
    ['pwHash',         sha256Hex(salt + INIT_PW)],
    ['pwChanged',      'false'],
  ];
  sheet.getRange(1, 2, initData.length, 1).setNumberFormat('@STRING@');
  sheet.getRange(1, 1, initData.length, 2).setValues(initData);
  sheet.getRange(1, 1, initData.length, 1).setFontWeight('bold').setBackground('#EDE9FE');
}
