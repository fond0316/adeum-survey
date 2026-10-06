/**
 * ─────────────────────────────────────────────────────────────
 *  경력 노트 · GAS (내 구글 시트 저장 서버)        Made by Cheon_R
 * ─────────────────────────────────────────────────────────────
 *  [처음 설치]
 *   1) 내 구글 시트 → 확장 프로그램 → Apps Script → 기본 코드 지우고 이 코드 전체 붙여넣기 → 저장
 *   2) 위쪽 함수 선택에서 setup 선택 → 실행 → 권한 허용 (처음 한 번)
 *   3) 배포 → 새 배포 → 유형: 웹 앱 / 실행: 나 / 액세스 권한: 모든 사용자 → 배포
 *   4) 웹 앱 URL(…/exec)을 복사해서 앱 첫 화면에 붙여넣기
 *
 *  [코드를 고친 뒤에는 꼭 "수정 배포"]
 *   배포 → 배포 관리 → 연필(수정) → 버전: 새 버전 → 배포
 *   ※ '새 배포'를 하면 주소가 바뀌어서 모든 기기에서 다시 연결해야 해요.
 *
 *  [비밀번호]
 *   처음 비밀번호는 아래 DEFAULT_PASSWORD 값이에요. 앱 설정에서 바꿀 수 있어요.
 *   잊어버렸으면 함수 선택에서 resetPassword 를 실행하면 처음 값으로 돌아가요.
 *
 *  [시트 탭]  기록 / 경력학력 / 설정
 *   탭을 통째로 지우지 마세요. 실수로 지웠다면 setup 을 다시 실행하면 빈 탭과 머리글이 복구돼요.
 *   시트에서 직접 줄을 추가해도 돼요 (ID 칸을 비워 두면 앱이 불러올 때 자동으로 채워요).
 *
 *  [AI 키]  앱 설정 > AI 에서 넣으면 이 스크립트의 '스크립트 속성'에만 저장돼요.
 *   (시트·브라우저에는 남지 않아요. 프로젝트 설정 > 스크립트 속성에서 직접 확인·삭제 가능)
 */

const APP_NAME = '경력 노트';
const GAS_VERSION = 1;
const DEFAULT_PASSWORD = '0413';
const TZ = 'Asia/Seoul';
const CHUNK_TTL = 600;          // 분할 전송 조각 보관 시간(초)
const HEADER_BG = '#EDEAF7';
const SETTINGS_SHEET = '설정';
const DEFAULT_MODELS = { claude: 'claude-sonnet-5-5', gpt: 'gpt-5-mini' };

// 시트 열 정의 [앱 키, 시트 머리글] — HTML(career.html)의 필드 키와 항상 일치시킬 것
const TABLES = {
  records: {
    sheet: '기록',
    cols: [
      ['id', 'ID'], ['cat', '카테고리'], ['title', '제목'],
      ['start', '시작일'], ['end', '종료일'], ['ongoing', '진행중'], ['year', '학년도'],
      ['affId', '소속ID'], ['affName', '소속'], ['role', '역할'],
      ['did', '한 일'], ['result', '성과'],
      ['m1k', '지표1'], ['m1v', '지표1 값'], ['m2k', '지표2'], ['m2v', '지표2 값'], ['m3k', '지표3'], ['m3v', '지표3 값'],
      ['f_org', '기관'], ['f_kind', '구분'], ['f_hours', '시간'], ['f_result', '등급·결과'],
      ['f_target', '대상'], ['f_scale', '규모·사용자'], ['f_subject', '과목·도구'], ['f_place', '위치·학급'],
      ['f_link', '링크'], ['f_no', '번호'],
      ['tags', '키워드'], ['star', '핵심'], ['area', '평가영역'],
      ['proof', '증빙'], ['proofType', '증빙 종류'], ['proofRef', '증빙 번호·링크'],
      ['memo', '메모'], ['draft', '정리필요'],
      ['created', '생성시각'], ['updated', '수정시각']
    ]
  },
  affs: {
    sheet: '경력학력',
    cols: [
      ['id', 'ID'], ['type', '구분'], ['name', '기관명'], ['dept', '부서·전공'],
      ['position', '직위·학위'], ['status', '형태·상태'],
      ['start', '시작일'], ['end', '종료일'], ['ongoing', '진행중'], ['memo', '메모'],
      ['created', '생성시각'], ['updated', '수정시각']
    ]
  }
};
Object.keys(TABLES).forEach(function (k) {
  const t = TABLES[k];
  t.k2h = {}; t.h2k = {};
  t.cols.forEach(function (c) { t.k2h[c[0]] = c[1]; t.h2k[c[1]] = c[0]; });
});

/* ───────── 입구 ───────── */

function doGet(e) {
  const p = (e && e.parameter) || {};
  const cb = safeCallback(p.cb);
  let out;
  try {
    out = handle(p);
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  const body = JSON.stringify(out);
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + body + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

function safeCallback(cb) {
  cb = String(cb || '');
  return /^[A-Za-z_$][\w$]{0,80}$/.test(cb) ? cb : '';
}

function handle(p) {
  if (!checkPassword(p.pw)) {
    Utilities.sleep(700);
    return { ok: false, code: 'PW', error: '비밀번호가 맞지 않아요.' };
  }
  if (p.a === 'chunk') {
    putChunk(p.tx, p.i, p.d);
    return { ok: true };
  }
  if (p.a !== 'run') return { ok: false, error: '알 수 없는 요청이에요.' };
  const b64 = p.tx ? takeChunks(p.tx, p.n) : String(p.d || '');
  const req = b64 ? JSON.parse(b64Decode(b64)) : {};
  return route(String(req.action || ''), req.payload || {});
}

function route(action, p) {
  switch (action) {
    case 'ping':
      return { ok: true, app: APP_NAME, version: GAS_VERSION, sheet: SpreadsheetApp.getActiveSpreadsheet().getName(), ai: aiStatus(), time: nowStr() };
    case 'loadAll':
      return withLock(function () {
        return { ok: true, records: readTable('records'), affs: readTable('affs'), settings: readSettings(), ai: aiStatus(), sheet: SpreadsheetApp.getActiveSpreadsheet().getName(), time: nowStr() };
      });
    case 'upsert':
      return withLock(function () { return { ok: true, row: upsertRow(tableKey(p.table), p.row || {}) }; });
    case 'remove':
      return withLock(function () { return { ok: true, removed: deleteById(tableKey(p.table), p.id) }; });
    case 'saveSettings':
      return withLock(function () { writeSettings(p.settings || {}, false); return { ok: true, settings: readSettings() }; });
    case 'importAll':
      return withLock(function () { return { ok: true, result: importAll(p) }; });
    case 'setPassword':
      return setPassword(p.newPw);
    case 'setAi':
      return setAi(p);
    case 'ai':
      return callAi(String(p.system || ''), String(p.prompt || ''), Number(p.maxTokens) || 2000);
    case 'aiTest':
      return callAi('짧게 답하세요.', '연결 확인입니다. "연결 성공"이라고만 답해 주세요.', 60);
    case 'aiVision':
      return callAiVision(String(p.system || ''), String(p.prompt || ''), p.images, Number(p.maxTokens) || 2000);
    default:
      return { ok: false, error: '알 수 없는 요청이에요: ' + action };
  }
}

/* ───────── 비밀번호 ───────── */

function getPassword() {
  return PropertiesService.getScriptProperties().getProperty('APP_PASSWORD') || DEFAULT_PASSWORD;
}
function checkPassword(pw) {
  return String(pw || '') === String(getPassword());
}
function setPassword(newPw) {
  newPw = String(newPw || '').trim();
  if (newPw.length < 4 || newPw.length > 40) return { ok: false, error: '비밀번호는 4~40자로 정해 주세요.' };
  PropertiesService.getScriptProperties().setProperty('APP_PASSWORD', newPw);
  return { ok: true };
}
/** 비밀번호를 잊었을 때 이 함수를 직접 실행 → 처음 값으로 돌아가요 */
function resetPassword() {
  PropertiesService.getScriptProperties().deleteProperty('APP_PASSWORD');
  Logger.log('비밀번호가 처음 값으로 돌아갔어요.');
}

/* ───────── 분할 전송 (긴 내용은 여러 조각으로 나눠 받기) ───────── */

function putChunk(tx, i, d) {
  tx = String(tx || '');
  const n = Number(i);
  if (!/^[A-Za-z0-9]{6,40}$/.test(tx)) throw new Error('전송 번호가 이상해요.');
  if (!(n >= 0 && n < 2000)) throw new Error('조각 번호가 이상해요.');
  d = String(d || '');
  if (!d || d.length > 20000) throw new Error('조각 크기가 이상해요.');
  CacheService.getScriptCache().put('cx_' + tx + '_' + n, d, CHUNK_TTL);
}

function takeChunks(tx, n) {
  tx = String(tx || '');
  n = Number(n);
  if (!/^[A-Za-z0-9]{6,40}$/.test(tx) || !(n >= 1 && n <= 2000)) throw new Error('전송 정보가 이상해요.');
  const cache = CacheService.getScriptCache();
  const keys = [];
  for (let i = 0; i < n; i++) keys.push('cx_' + tx + '_' + i);
  let s = '';
  for (let i = 0; i < keys.length; i += 100) {
    const part = keys.slice(i, i + 100);
    const got = cache.getAll(part);
    part.forEach(function (k) {
      if (got[k] === undefined || got[k] === null) throw new Error('전송 중 일부가 빠졌어요. 다시 시도해 주세요.');
      s += got[k];
    });
  }
  for (let i = 0; i < keys.length; i += 100) cache.removeAll(keys.slice(i, i + 100));
  return s;
}

function b64Decode(s) {
  s = String(s).replace(/\s/g, '');
  while (s.length % 4) s += '=';
  return Utilities.newBlob(Utilities.base64DecodeWebSafe(s)).getDataAsString('UTF-8');
}

/* ───────── 시트 읽기/쓰기 ───────── */

function tableKey(k) {
  if (k === 'records' || k === 'affs') return k;
  throw new Error('잘못된 표 이름이에요.');
}

function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function nowStr() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
}

function ensureSize(sh, rows, cols) {
  if (sh.getMaxColumns() < cols) sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());
  if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows() + 50);
}

function writeHeader(sh, startCol, headers) {
  ensureSize(sh, 2, startCol + headers.length - 1);
  sh.getRange(1, startCol, sh.getMaxRows(), headers.length).setNumberFormat('@');
  sh.getRange(1, startCol, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground(HEADER_BG);
}

function ensureTable(key) {
  const t = TABLES[key];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headers = t.cols.map(function (c) { return c[1]; });
  let sh = ss.getSheetByName(t.sheet);
  if (!sh) {
    sh = ss.insertSheet(t.sheet);
    writeHeader(sh, 1, headers);
    sh.setFrozenRows(1);
    return sh;
  }
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const cur = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function (v) { return String(v).trim(); });
  if (cur.every(function (v) { return !v; })) {
    writeHeader(sh, 1, headers);
    sh.setFrozenRows(1);
    return sh;
  }
  // 빠진 머리글만 맨 뒤에 추가 (기존 열 순서·직접 추가한 열은 그대로 둠)
  const missing = headers.filter(function (h) { return cur.indexOf(h) < 0; });
  if (missing.length) writeHeader(sh, lastCol + 1, missing);
  return sh;
}

function getHeaders(sh) {
  const lastCol = Math.max(sh.getLastColumn(), 1);
  return sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function (v) { return String(v).trim(); });
}

// 시트 → 앱 : 수식 방지용 작은따옴표가 남아 있으면 떼기
function cellIn(v) {
  let s = String(v === null || v === undefined ? '' : v);
  if (/^'[=+\-@]/.test(s)) s = s.slice(1);
  return s;
}
// 앱 → 시트 : '='로 시작하면 수식으로 바뀌지 않게 작은따옴표 붙이기
function cellOut(v) {
  if (v === null || v === undefined || v === false) return '';
  if (v === true) return 'O';
  let s = String(v);
  if (s.length > 45000) s = s.slice(0, 45000);
  if (/^=/.test(s)) s = "'" + s;
  return s;
}

function readTable(key) {
  const t = TABLES[key];
  const sh = ensureTable(key);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const headers = getHeaders(sh);
  const vals = sh.getRange(2, 1, last - 1, headers.length).getDisplayValues();
  const idCol = headers.indexOf('ID');
  const out = [];
  const fixes = [];
  vals.forEach(function (row, r) {
    const obj = {};
    let any = false;
    headers.forEach(function (h, i) {
      const k = t.h2k[h];
      if (!k) return;
      const v = cellIn(row[i]);
      obj[k] = v;
      if (v !== '' && k !== 'id') any = true;
    });
    if (!any) return;                                   // 빈 줄 건너뛰기
    let id = String(obj.id || '').trim();
    if (!id) {                                          // 시트에서 직접 쓴 줄 → ID 자동 부여
      id = (key === 'records' ? 'r' : 'a') + 's' + Date.now().toString(36) + r;
      fixes.push([r + 2, id]);
    }
    obj.id = id;
    out.push(obj);
  });
  if (fixes.length && idCol >= 0) {
    fixes.forEach(function (f) {
      sh.getRange(f[0], idCol + 1).setNumberFormat('@').setValue(f[1]);
    });
  }
  return out;
}

function findRow(sh, headers, id) {
  const idCol = headers.indexOf('ID');
  const last = sh.getLastRow();
  if (idCol < 0 || last < 2) return -1;
  const ids = sh.getRange(2, idCol + 1, last - 1, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(id)) return i + 2;
  }
  return -1;
}

function upsertRow(key, row) {
  const t = TABLES[key];
  const sh = ensureTable(key);
  const headers = getHeaders(sh);
  const id = String(row.id || '').trim();
  if (!id) throw new Error('ID가 없어요.');
  const now = nowStr();
  const rowIdx = findRow(sh, headers, id);
  let existing = null;
  if (rowIdx > 0) existing = sh.getRange(rowIdx, 1, 1, headers.length).getDisplayValues()[0];
  const createdCol = headers.indexOf('생성시각');
  const created = existing && createdCol >= 0 && existing[createdCol] ? existing[createdCol] : (row.created || now);
  const obj = {};
  Object.keys(row).forEach(function (k) { obj[k] = row[k]; });
  obj.id = id; obj.created = created; obj.updated = now;
  const values = headers.map(function (h, i) {
    const k = t.h2k[h];
    if (!k) return existing ? existing[i] : '';        // 직접 추가한 열은 그대로 유지
    return cellOut(obj[k]);
  });
  const target = rowIdx > 0 ? rowIdx : Math.max(sh.getLastRow(), 1) + 1;
  ensureSize(sh, target, headers.length);
  const rng = sh.getRange(target, 1, 1, headers.length);
  rng.setNumberFormat('@');
  rng.setValues([values]);
  const back = {};
  headers.forEach(function (h, i) { const k = t.h2k[h]; if (k) back[k] = cellIn(values[i]); });
  return back;
}

function deleteById(key, id) {
  const sh = ensureTable(key);
  const headers = getHeaders(sh);
  const rowIdx = findRow(sh, headers, String(id || '').trim());
  if (rowIdx < 0) return false;
  sh.deleteRow(rowIdx);
  return true;
}

function writeBulk(key, rows, mode) {
  const t = TABLES[key];
  const sh = ensureTable(key);
  const headers = getHeaders(sh);
  const idIdx = headers.indexOf('ID');
  const now = nowStr();
  const last = sh.getLastRow();
  let data = [];
  if (mode === 'merge' && last >= 2) data = sh.getRange(2, 1, last - 1, headers.length).getDisplayValues();
  const pos = {};
  data.forEach(function (r, i) { const id = String(r[idIdx]).trim(); if (id) pos[id] = i; });
  (rows || []).forEach(function (src) {
    if (!src) return;
    const id = String(src.id || '').trim();
    if (!id) return;
    const o = {};
    Object.keys(src).forEach(function (k) { o[k] = src[k]; });
    o.id = id; o.created = src.created || now; o.updated = src.updated || now;
    const old = pos[id] !== undefined ? data[pos[id]] : null;
    const arr = headers.map(function (h, i) {
      const k = t.h2k[h];
      if (!k) return old ? old[i] : '';
      return cellOut(o[k]);
    });
    if (old) data[pos[id]] = arr; else { pos[id] = data.length; data.push(arr); }
  });
  if (last >= 2) sh.getRange(2, 1, last - 1, headers.length).clearContent();
  if (data.length) {
    ensureSize(sh, data.length + 1, headers.length);
    const rng = sh.getRange(2, 1, data.length, headers.length);
    rng.setNumberFormat('@');
    rng.setValues(data);
  }
  return data.length;
}

function importAll(p) {
  const mode = p.mode === 'replace' ? 'replace' : 'merge';
  const res = {
    records: writeBulk('records', Array.isArray(p.records) ? p.records : [], mode),
    affs: writeBulk('affs', Array.isArray(p.affs) ? p.affs : [], mode)
  };
  if (p.settings && typeof p.settings === 'object') writeSettings(p.settings, mode === 'replace');
  return res;
}

/* ───────── 설정 탭 (항목 | 값) ───────── */

function ensureSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SETTINGS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SETTINGS_SHEET);
    writeHeader(sh, 1, ['항목', '값']);
    sh.setFrozenRows(1);
  } else if (!String(sh.getRange(1, 1).getDisplayValue()).trim()) {
    writeHeader(sh, 1, ['항목', '값']);
  }
  return sh;
}

function readSettings() {
  const sh = ensureSettings();
  const last = sh.getLastRow();
  const o = {};
  if (last < 2) return o;
  sh.getRange(2, 1, last - 1, 2).getDisplayValues().forEach(function (r) {
    const k = String(r[0]).trim();
    if (k) o[k] = cellIn(r[1]);
  });
  return o;
}

function writeSettings(obj, replaceAll) {
  const sh = ensureSettings();
  if (replaceAll && sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
  const last = sh.getLastRow();
  const vals = last >= 2 ? sh.getRange(2, 1, last - 1, 2).getDisplayValues() : [];
  const idx = {};
  vals.forEach(function (r, i) { const k = String(r[0]).trim(); if (k) idx[k] = i; });
  Object.keys(obj || {}).forEach(function (k) {
    const key = String(k).trim();
    if (!key) return;
    const v = cellOut(obj[k]);
    if (idx[key] !== undefined) vals[idx[key]][1] = v;
    else { idx[key] = vals.length; vals.push([key, v]); }
  });
  if (vals.length) {
    ensureSize(sh, vals.length + 1, 2);
    const rng = sh.getRange(2, 1, vals.length, 2);
    rng.setNumberFormat('@');
    rng.setValues(vals);
  }
}

/* ───────── AI (키는 스크립트 속성에만 보관) ───────── */

function aiStatus() {
  const pr = PropertiesService.getScriptProperties();
  const provider = pr.getProperty('AI_PROVIDER') || 'claude';
  const key = pr.getProperty('AI_KEY') || '';
  return {
    provider: provider,
    model: pr.getProperty('AI_MODEL') || DEFAULT_MODELS[provider] || '',
    hasKey: !!key,
    keyHint: key ? '…' + key.slice(-4) : ''
  };
}

function setAi(p) {
  const pr = PropertiesService.getScriptProperties();
  const provider = p.provider === 'gpt' ? 'gpt' : 'claude';
  pr.setProperty('AI_PROVIDER', provider);
  const model = String(p.model || '').trim().slice(0, 100);
  if (model) pr.setProperty('AI_MODEL', model); else pr.deleteProperty('AI_MODEL');
  if (p.clearKey) pr.deleteProperty('AI_KEY');
  const key = String(p.key || '').trim();
  if (key) pr.setProperty('AI_KEY', key);
  return { ok: true, ai: aiStatus() };
}

function apiError(code, body) {
  const msg = body && body.error && (body.error.message || body.error.type) ? String(body.error.message || body.error.type) : '';
  if (code === 401 || code === 403) return '키가 올바르지 않거나 권한이 없어요. (' + code + ') ' + msg;
  if (code === 404) return '모델 이름을 확인해 주세요. (' + code + ') ' + msg;
  if (code === 429) return '사용 한도나 잔액을 확인해 주세요. (' + code + ') ' + msg;
  if (code >= 500) return 'AI 서버가 바빠요. 잠시 후 다시 시도해 주세요. (' + code + ')';
  return 'AI 요청이 실패했어요. (' + code + ') ' + msg;
}

function callAi(system, prompt, maxTokens) {
  const st = aiStatus();
  const key = PropertiesService.getScriptProperties().getProperty('AI_KEY');
  if (!key) return { ok: false, code: 'NOKEY', error: 'AI 키가 아직 없어요. 설정 > AI에서 키를 넣어 주세요.' };
  if (!prompt) return { ok: false, error: '보낼 내용이 없어요.' };
  maxTokens = Math.max(50, Math.min(8000, Math.round(maxTokens)));
  let res, body = null, text = '';
  if (st.provider === 'gpt') {
    res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + key },
      payload: JSON.stringify({
        model: st.model,
        messages: [{ role: 'system', content: system }, { role: 'user', content: prompt }],
        max_completion_tokens: 12000
      })
    });
    try { body = JSON.parse(res.getContentText()); } catch (e) { body = null; }
    if (res.getResponseCode() >= 300) return { ok: false, error: apiError(res.getResponseCode(), body) };
    text = body && body.choices && body.choices[0] && body.choices[0].message ? String(body.choices[0].message.content || '') : '';
  } else {
    res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({
        model: st.model,
        max_tokens: maxTokens,
        system: system,
        messages: [{ role: 'user', content: prompt }]
      })
    });
    try { body = JSON.parse(res.getContentText()); } catch (e) { body = null; }
    if (res.getResponseCode() >= 300) return { ok: false, error: apiError(res.getResponseCode(), body) };
    text = body && body.content ? body.content.filter(function (c) { return c.type === 'text'; }).map(function (c) { return c.text; }).join('') : '';
  }
  text = String(text || '').trim();
  if (!text) return { ok: false, error: 'AI 응답이 비어 있어요. 잠시 후 다시 시도해 주세요.' };
  return { ok: true, text: text, provider: st.provider, model: st.model };
}

/* ───────── AI 사진 읽기 (Claude 전용 · 사진은 저장하지 않고 한 번 전달만 해요) ───────── */

function callAiVision(system, prompt, images, maxTokens) {
  const st = aiStatus();
  const key = PropertiesService.getScriptProperties().getProperty('AI_KEY');
  if (!key) return { ok: false, code: 'NOKEY', error: 'AI 키가 아직 없어요. 설정 > AI에서 키를 넣어 주세요.' };
  if (st.provider === 'gpt') return { ok: false, error: '사진 읽기는 Claude로 설정했을 때만 쓸 수 있어요. 설정 > AI에서 Claude를 골라 주세요.' };
  if (!prompt) return { ok: false, error: '보낼 내용이 없어요.' };
  if (!Array.isArray(images) || !images.length) return { ok: false, error: '사진이 없어요.' };
  if (images.length > 6) return { ok: false, error: '사진은 한 번에 6장까지 읽을 수 있어요.' };
  maxTokens = Math.max(50, Math.min(8000, Math.round(maxTokens)));
  const okType = { 'image/jpeg': 1, 'image/png': 1, 'image/webp': 1 };
  const content = [];
  for (let i = 0; i < images.length; i++) {
    const im = images[i] || {};
    const mt = String(im.mt || 'image/jpeg');
    const data = String(im.data || '');
    if (!okType[mt] || !data || data.length > 4000000) return { ok: false, error: '사진 형식이나 크기가 맞지 않아요.' };
    content.push({ type: 'image', source: { type: 'base64', media_type: mt, data: data } });
  }
  content.push({ type: 'text', text: prompt });
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({ model: st.model, max_tokens: maxTokens, system: system, messages: [{ role: 'user', content: content }] })
  });
  let body = null;
  try { body = JSON.parse(res.getContentText()); } catch (e) { body = null; }
  if (res.getResponseCode() >= 300) return { ok: false, error: apiError(res.getResponseCode(), body) };
  const text = body && body.content ? body.content.filter(function (c) { return c.type === 'text'; }).map(function (c) { return c.text; }).join('').trim() : '';
  if (!text) return { ok: false, error: 'AI 응답이 비어 있어요. 잠시 후 다시 시도해 주세요.' };
  return { ok: true, text: text, provider: st.provider, model: st.model };
}

/* ───────── 직접 실행용 ───────── */

/** 처음 설치 때 한 번 실행 (권한 허용) · 탭을 실수로 지웠을 때 다시 실행하면 복구 */
function setup() {
  ensureTable('records');
  ensureTable('affs');
  ensureSettings();
  Logger.log('준비 완료! 이제 배포 → 새 배포(웹 앱)를 진행해 주세요.');
}
