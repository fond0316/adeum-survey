/**
 * ACTIVITY CHECK → 생기부AI 연동 추가 코드 (아듬샘 시트 전용) — Made by Cheon_R
 *
 * [붙이는 법] 메인 코드(Cheon-R-activity-check-GAS.gs)가 최신인지 먼저 확인 →
 *   Apps Script 왼쪽 [파일 +] → [스크립트] → 이름 '생기부연동' → 이 코드 전체 붙여넣기 → 저장 →
 *   [배포] → [배포 관리] → ✏️ 수정 → 새 버전 → 배포   (새 배포 금지)
 * 이 파일이 없는 시트(다른 선생님)에서는 연동 요청이 와도 "설치되어 있지 않아요"만 나가고 데이터는 나가지 않아요.
 * 읽기 전용이에요. 학생 PIN·교사 메모·학생 질문은 내보내지 않아요. 관리자 비밀번호가 있어야 해요.
 */
var SGB_DESTS = ['자율', '진로', '행발', '안보냄'];   // 항목마다 생기부AI의 어느 탭으로 보낼지 ('안보냄'이면 내보내지 않아요)

/** 기본값: 자율→자율, 진로→진로, 그 밖(교과·동아리)은 보내지 않아요 */
function sgbDefaultDest_(cat) { return cat === '자율' ? '자율' : cat === '진로' ? '진로' : '안보냄'; }

/** 아듬샘이 기본값과 다르게 정한 것만 설정 탭('생기부연동_목적지')에 JSON으로 저장돼 있어요 */
function sgbMap_() {
  try {
    var o = JSON.parse(getSetting_('생기부연동_목적지') || '{}');
    return (o && typeof o === 'object' && !Array.isArray(o)) ? o : {};
  } catch (e) { return {}; }
}

/** 메인 코드의 route_ 가 부르는 입구 */
function sgbRoute_(p) {
  need_(p);
  if (p.action === 'saengibuSaveMap') return sgbSaveMap_(p);
  return sgbExport_(p);
}

/** 정식 항목만 (학생이 직접 추가한 활동은 담임이 정식으로 올리기 전까지 제외) */
function sgbItems_() {
  var map = sgbMap_();
  return sortItems_(rows_('items').filter(function (r) { return r[0] && !r[15]; }).map(itemObj_)).map(function (it) {
    var d = map[it.id];
    return { id: it.id, cat: it.cat, title: it.title, on: it.on, dest: SGB_DESTS.indexOf(d) >= 0 ? d : sgbDefaultDest_(it.cat), custom: SGB_DESTS.indexOf(d) >= 0 };
  });
}

/** 내보내기: meta=1 이면 항목(보낼 곳 설정)만 가볍게, 아니면 학생별 기록까지 */
function sgbExport_(p) {
  var items = sgbItems_();
  var base = { ok: true, app: 'activity', gasVer: GAS_VER, className: getSetting_('학급명'), year: year_(), exportedAt: now_() };
  if (String(p.meta) === '1') { base.items = items; return base; }
  var by = {}, order = {};
  items.forEach(function (it, i) { by[it.id] = it; order[it.id] = i; });
  var recsBy = {};
  rows_('records').filter(function (r) { return r[0]; }).forEach(function (r) { (recsBy[r[0]] = recsBy[r[0]] || []).push(recObj_(r)); });
  var students = rows_('students').filter(function (r) { return r[0]; }).map(function (r) {
    var sid = r[0], entries = [];
    (recsBy[sid] || []).forEach(function (rc) {
      var it = by[rc.iid];
      if (!it || !it.on || it.dest === '안보냄') return;   // 정식 항목이 아니거나, 숨김·보내지 않음 항목은 제외
      if (rc.part !== 'yes' && !rc.summary) return;       // 참여했거나 요약을 적은 기록만
      entries.push({ iid: rc.iid, cat: it.cat, dest: it.dest, title: it.title, n: rc.n, part: rc.part, report: rc.report, topic: rc.topic,
        career: rc.career, summary: rc.summary, link: rc.link, tchk: rc.tchk, star: rc.star, deep: rc.deep, atype: rc.atype, rterm: rc.rterm });
    });
    entries.sort(function (a, b) { return (order[a.iid] - order[b.iid]) || (a.n - b.n); });
    return { sid: sid, num: parseInt(String(sid).slice(-2), 10) || 0, name: r[1], career: r[3], entries: entries };
  });
  base.items = items; base.students = students;
  return base;
}

/** 항목별 보낼 곳 저장 — data = {항목ID: '자율'|'진로'|'행발'|'안보냄'} (기본값과 같은 것은 저장하지 않아요) */
function sgbSaveMap_(p) {
  var d = JSON.parse(p.data || '{}');
  if (!d || typeof d !== 'object' || Array.isArray(d)) throw new Error('설정 형식이 올바르지 않아요.');
  var cat = {};
  sgbItems_().forEach(function (it) { cat[it.id] = it.cat; });
  var out = {}, n = 0;
  Object.keys(d).forEach(function (id) {
    if (!(id in cat)) return;   // 없는 항목은 무시
    if (SGB_DESTS.indexOf(d[id]) < 0) throw new Error('보낼 곳은 자율·진로·행발·안보냄 중에서 골라 주세요.');
    if (d[id] !== sgbDefaultDest_(cat[id])) { out[id] = d[id]; n++; }
  });
  return lock_(function () { setSetting_('생기부연동_목적지', JSON.stringify(out)); return { ok: true, count: n }; });
}
