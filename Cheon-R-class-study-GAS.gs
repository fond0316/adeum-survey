// ============================================================
// 협력학습 자율스터디 - Google Apps Script v10
// v5: 사진 구글폼 제출 방식으로 변경 (드라이브 연동 제거)
// v6: 학급자율활동 홈 항목 저장/불러오기 추가 ('활동항목' 시트) — 기존 기능은 변경 없음
// v7: 제출형 항목 — 제출 질문('제출폼항목' 시트), 학생 제출 저장, 항목별 제출 마감
// v8: 영역(자율/진로/기타) 구조 — 제출은 '제출_자율/제출_진로/제출_기타' 탭에 학생당 1행(열: 주제 · 질문),
//     '제출로그' 탭, 제출 마감일, 관리자 '주제별 보기'(관리자 암호 ADMIN_PW 필요)
//     v10: 제출 질문에 '객관식(choice)' 유형 추가 — '제출폼항목' 시트에 '선택지' 열(한 줄에 하나)
//     v11: 학생 핀(4자리 이상) — 내 제출 불러오기('getMySubmission'), 제출·조회 시 핀 확인, '학생핀' 시트(해시 저장),
//          관리자 핀 초기화('adminResetPin'), 주제별 보기에 핀 설정 여부 표시. 처음 제출 학생은 반+번호+이름으로 매칭해 핀을 새로 만들어요.
// ============================================================

const SHEET_MEMBER = '일반스터디';
const SHEET_KING = '왕멘토';
const SHEET_PLAN = '스터디계획표';
const SHEET_NOTICE = '스터디공지';
const SHEET_FINAL = '최종느낀점';
const SHEET_SETTINGS = '앱설정';
const SHEET_HUB = '활동항목';   // v6 추가: 학급자율활동 홈 항목
const SHEET_FORMS = '제출폼항목'; // v7 추가: 제출형 항목의 질문 목록
const SHEET_LOG = '제출로그';     // v8 추가: 제출할 때마다 한 줄씩 기록
const SHEET_FILES = '첨부파일';   // v9 추가: 항목에 붙인 PDF·사진 (파일 자체는 내 구글드라이브에 저장)
const FILE_HEADERS = ['항목ID','파일ID','파일이름','종류','순서'];
const FILES_FOLDER = '학급자율활동_첨부'; // 드라이브에 만들어지는 폴더 이름
const FILE_MAX_B64 = 14000000;            // 업로드 한도(글자 수 기준, 약 10MB)
const SHEET_PIN = '학생핀';               // v11 추가: 학생 핀(해시) — 시트에 핀 원문은 저장하지 않아요
const PIN_HEADERS = ['반','번호','핀(해시)','설정시각'];
const PIN_MAX_FAIL = 5;                   // 핀을 연속으로 틀리면 10분 잠금
const ADMIN_MAX_FAIL = 20;                // 관리자 암호를 10분 안에 20번 틀리면 10분 잠금 (무작위 대입 방지, 선생님 오타 몇 번은 괜찮아요)

// ★ v9 처음 한 번만: Apps Script 편집기 위쪽 함수 선택에서 "승인용_드라이브권한"을 고르고 ▶ 실행 → 권한 승인
//    (PDF·사진 올리기에 드라이브 권한이 필요해요) 그다음 "수정 배포"
function 승인용_드라이브권한() {
  const f = DriveApp.getRootFolder();
  Logger.log('드라이브 권한 승인 완료: ' + f.getName());
}

// ★★ 관리자 암호 (v8) — 따옴표 사이에 암호를 적고 "수정 배포"하세요. (예: 'mypass1234')
//    · 앱 관리자 화면의 "관리자 암호"에 같은 값을 입력하면 '주제별 보기'와 '저장하기'를 쓸 수 있어요.
//    · 비워두면 '주제별 보기'는 막혀요. 학생 답변을 읽는 기능이라 꼭 설정해 주세요.
//    · 학생이 볼 수 있는 HTML에는 이 암호가 들어 있지 않아요. (기본값 0413은 쓰지 마세요)
const ADMIN_PW = '0413';

function getSettingsSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SETTINGS);
    sheet.getRange(1,1,1,2).setValues([['키','값']]);
    sheet.getRange(1,1,1,2).setBackground('#6B7280').setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.appendRow(['isClosed', 'false']);
    sheet.appendRow(['closedAt', '']);
  }
  return sheet;
}

function getSetting(ss, key) {
  const sheet = getSettingsSheet(ss);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === key) return String(data[i][1]);
  }
  return '';
}

function setSetting(ss, key, value) {
  const sheet = getSettingsSheet(ss);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === key) { sheet.getRange(i+1,2).setValue(value); return; }
  }
  sheet.appendRow([key, value]);
}

function toKST(isoString) {
  if (!isoString) return '';
  const d = new Date(isoString);
  return Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
}

function doOptions(e) {
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // v11: 관리자 암호를 너무 많이 틀렸으면 잠깐 막아요
    if (data.pw !== undefined && adminLockedOut()) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'auth', message: '암호를 여러 번 틀려서 10분 동안 잠겼어요. 잠시 후 다시 시도해 주세요.' })).setMimeType(ContentService.MimeType.JSON);
    }

    // (전체 마감의 영향을 받지 않는 요청들. submitActivity는 항목별 마감으로 따로 확인해요)
    const adminTypes = ['setClosed', 'clearAllData', 'saveHubItems', 'submitActivity', 'adminTopicView', 'adminTopicSummary', 'uploadHubFile', 'getMySubmission', 'pinStatus', 'adminResetPin', 'adminLogin'];
    if (!adminTypes.includes(data.type)) {
      if (getSetting(ss, 'isClosed') === 'true') {
        return ContentService
          .createTextOutput(JSON.stringify({ status: 'closed', message: '현재 제출이 마감되었어요. 선생님께 문의해 주세요.' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // photoSubmitted 값 정규화 (없으면 'X')
    if (!data.photoSubmitted) data.photoSubmitted = 'X';

    if (data.type === 'king') {
      writeKingRecord(ss, data);
    } else if (data.type === 'final') {
      writeFinalRecord(ss, data);
    } else if (data.type === 'updateMember') {
      updateMemberRecord(ss, data);
    } else if (data.type === 'updateKing') {
      updateKingRecord(ss, data);
    } else if (data.type === 'setClosed') {
      setSetting(ss, 'isClosed', data.isClosed ? 'true' : 'false');
      setSetting(ss, 'closedAt', data.isClosed ? new Date().toISOString() : '');
    } else if (data.type === 'plan') {
      writePlanRecord(ss, data);
    } else if (data.type === 'saveHubItems') {
      const sr = saveHubItems(ss, data);
      if (sr && sr.status === 'auth') return ContentService.createTextOutput(JSON.stringify(sr)).setMimeType(ContentService.MimeType.JSON);
    } else if (data.type === 'adminTopicView') {
      return ContentService.createTextOutput(JSON.stringify(adminTopicView(ss, data))).setMimeType(ContentService.MimeType.JSON);
    } else if (data.type === 'uploadHubFile') {
      return ContentService.createTextOutput(JSON.stringify(uploadHubFile(ss, data))).setMimeType(ContentService.MimeType.JSON);
    } else if (data.type === 'adminTopicSummary') {
      return ContentService.createTextOutput(JSON.stringify(adminTopicSummary(ss, data))).setMimeType(ContentService.MimeType.JSON);
    } else if (data.type === 'submitActivity') {
      return ContentService.createTextOutput(JSON.stringify(submitActivity(ss, data))).setMimeType(ContentService.MimeType.JSON);
    } else if (data.type === 'adminLogin') {
      return ContentService.createTextOutput(JSON.stringify(adminLogin(data))).setMimeType(ContentService.MimeType.JSON);
    } else if (data.type === 'getMySubmission') {
      return ContentService.createTextOutput(JSON.stringify(getMySubmission(ss, data))).setMimeType(ContentService.MimeType.JSON);
    } else if (data.type === 'pinStatus') {
      return ContentService.createTextOutput(JSON.stringify(pinStatus(ss, data))).setMimeType(ContentService.MimeType.JSON);
    } else if (data.type === 'adminResetPin') {
      return ContentService.createTextOutput(JSON.stringify(adminResetPin(ss, data))).setMimeType(ContentService.MimeType.JSON);
    } else if (data.type === 'clearAllData') {
      clearAllData(ss);
    } else {
      writeMemberRecord(ss, data);
      if (data.memberRole === '멘토' && data.notice) writeNotice(ss, data);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

function writeMemberRecord(ss, d) {
  let sheet = ss.getSheetByName(SHEET_MEMBER);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_MEMBER);
    sheet.getRange(1,1,1,19).setValues([[
      '제출시각','분기','스터디명','회차','이름','반','번호','멘토/멘티',
      '활동날짜','활동시간','활동장소','활동제목','활동내용','느낀점','다음회차예고',
      '출석','결석','누적시간(분)','사진첨부'
    ]]);
    sheet.getRange(1,1,1,19).setBackground('#3B82F6').setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([
    toKST(d.timestamp),d.quarter,d.studyName,d.round,d.name,d.classNum,d.studentNum,d.memberRole||'',
    d.date,d.time,d.place||'',d.title,d.content,d.reflection||'',d.nextPlan||'',
    d.present||'',d.absent||'',d.totalMinutes||0,d.photoSubmitted||'X'
  ]);
  const lr=sheet.getLastRow();
  if(lr%2===0)sheet.getRange(lr,1,1,19).setBackground('#EFF6FF');
  if(lr===2)sheet.autoResizeColumns(1,19);
}

function updateMemberRecord(ss, d) {
  const sheet = ss.getSheetByName(SHEET_MEMBER);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const h = data[0];
  for (let i = 1; i < data.length; i++) {
    if (
      String(data[i][h.indexOf('이름')]) === String(d.name) &&
      String(data[i][h.indexOf('반')]) === String(d.classNum) &&
      String(data[i][h.indexOf('번호')]) === String(d.studentNum) &&
      String(data[i][h.indexOf('분기')]) === String(d.quarter) &&
      String(data[i][h.indexOf('스터디명')]) === String(d.studyName) &&
      String(data[i][h.indexOf('회차')]) === String(d.round)
    ) {
      const rn = i+1;
      sheet.getRange(rn,h.indexOf('제출시각')+1).setValue(toKST(d.timestamp));
      sheet.getRange(rn,h.indexOf('멘토/멘티')+1).setValue(d.memberRole||'');
      sheet.getRange(rn,h.indexOf('활동날짜')+1).setValue(d.date);
      sheet.getRange(rn,h.indexOf('활동시간')+1).setValue(d.time);
      sheet.getRange(rn,h.indexOf('활동장소')+1).setValue(d.place||'');
      sheet.getRange(rn,h.indexOf('활동제목')+1).setValue(d.title);
      sheet.getRange(rn,h.indexOf('활동내용')+1).setValue(d.content);
      sheet.getRange(rn,h.indexOf('느낀점')+1).setValue(d.reflection||'');
      sheet.getRange(rn,h.indexOf('다음회차예고')+1).setValue(d.nextPlan||'');
      sheet.getRange(rn,h.indexOf('출석')+1).setValue(d.present||'');
      sheet.getRange(rn,h.indexOf('결석')+1).setValue(d.absent||'');
      sheet.getRange(rn,h.indexOf('누적시간(분)')+1).setValue(d.totalMinutes||0);
      sheet.getRange(rn,h.indexOf('사진첨부')+1).setValue(d.photoSubmitted||'X');
      return;
    }
  }
}

function writeKingRecord(ss, d) {
  let sheet = ss.getSheetByName(SHEET_KING);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_KING);
    sheet.getRange(1,1,1,15).setValues([[
      '제출시각','분기','담당과목','회차','이름','반','번호',
      '활동날짜','활동시간','학습내용','느낀점','다음활동예고',
      '질문처리내역(JSON)','누적시간(분)','사진첨부'
    ]]);
    sheet.getRange(1,1,1,15).setBackground('#10B981').setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([
    toKST(d.timestamp),d.quarter,d.subject,d.round,d.name,d.classNum,d.studentNum,
    d.date,d.time,d.selfStudy,d.reflection||'',d.nextPlan||'',
    d.questions||'[]',d.totalMinutes||0,d.photoSubmitted||'X'
  ]);
  const lr=sheet.getLastRow();
  if(lr%2===0)sheet.getRange(lr,1,1,15).setBackground('#ECFDF5');
  if(lr===2)sheet.autoResizeColumns(1,15);
}

function updateKingRecord(ss, d) {
  const sheet = ss.getSheetByName(SHEET_KING);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const h = data[0];
  for (let i = 1; i < data.length; i++) {
    if (
      String(data[i][h.indexOf('이름')]) === String(d.name) &&
      String(data[i][h.indexOf('반')]) === String(d.classNum) &&
      String(data[i][h.indexOf('번호')]) === String(d.studentNum) &&
      String(data[i][h.indexOf('분기')]) === String(d.quarter) &&
      String(data[i][h.indexOf('담당과목')]) === String(d.subject) &&
      String(data[i][h.indexOf('회차')]) === String(d.round)
    ) {
      const rn = i+1;
      sheet.getRange(rn,h.indexOf('제출시각')+1).setValue(toKST(d.timestamp));
      sheet.getRange(rn,h.indexOf('활동날짜')+1).setValue(d.date);
      sheet.getRange(rn,h.indexOf('활동시간')+1).setValue(d.time);
      sheet.getRange(rn,h.indexOf('학습내용')+1).setValue(d.selfStudy);
      sheet.getRange(rn,h.indexOf('느낀점')+1).setValue(d.reflection||'');
      sheet.getRange(rn,h.indexOf('다음활동예고')+1).setValue(d.nextPlan||'');
      sheet.getRange(rn,h.indexOf('질문처리내역(JSON)')+1).setValue(d.questions||'[]');
      sheet.getRange(rn,h.indexOf('누적시간(분)')+1).setValue(d.totalMinutes||0);
      sheet.getRange(rn,h.indexOf('사진첨부')+1).setValue(d.photoSubmitted||'X');
      return;
    }
  }
}

function writeFinalRecord(ss, d) {
  let sheet = ss.getSheetByName(SHEET_FINAL);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_FINAL);
    sheet.getRange(1,1,1,11).setValues([[
      '제출시각','분기','스터디명','이름','반','번호','역할','배운점/성장한점','아쉬운점/다음목표','기억에남는교과내용','최종소감'
    ]]);
    sheet.getRange(1,1,1,11).setBackground('#8B5CF6').setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([toKST(d.timestamp),d.quarter,d.studyName,d.name,d.classNum,d.studentNum,d.role||'',d.growth||'',d.improve||'',d.memory||'',d.oneLine||'']);
  const lr=sheet.getLastRow();
  if(lr%2===0)sheet.getRange(lr,1,1,11).setBackground('#F5F3FF');
  if(lr===2)sheet.autoResizeColumns(1,11);
}

function writeNotice(ss, d) {
  let sheet = ss.getSheetByName(SHEET_NOTICE);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NOTICE);
    sheet.getRange(1,1,1,4).setValues([['분기','스터디명','공지내용','최종수정']]);
    sheet.getRange(1,1,1,4).setBackground('#F59E0B').setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  const data = sheet.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(String(data[i][0])===String(d.quarter)&&String(data[i][1])===String(d.studyName)){
      sheet.getRange(i+1,3).setValue(d.notice);sheet.getRange(i+1,4).setValue(toKST(d.timestamp));return;
    }
  }
  sheet.appendRow([d.quarter,d.studyName,d.notice,toKST(d.timestamp)]);
}

function writePlanRecord(ss, d) {
  let sheet = ss.getSheetByName(SHEET_PLAN);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_PLAN);
    sheet.getRange(1,1,1,11).setValues([[
      '제출시각','분기','멘토이름','반','번호','과목','팀명','팀원명단','스터디시간대','목표시간','계획내용'
    ]]);
    sheet.getRange(1,1,1,11).setBackground('#F59E0B').setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([
    toKST(d.timestamp), d.quarter, d.mentorName, d.classNum, d.studentNum,
    d.subject, d.teamName, d.members, d.timeSlots, d.goalTime||'', d.content
  ]);
  const lr = sheet.getLastRow();
  if (lr%2===0) sheet.getRange(lr,1,1,11).setBackground('#FFFBEB');
  if (lr===2) sheet.autoResizeColumns(1,11);
}

function clearAllData(ss) {
  [SHEET_MEMBER,SHEET_KING,SHEET_PLAN,SHEET_NOTICE,SHEET_FINAL].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet && sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow()-1);
  });
  setSetting(ss,'isClosed','false');
  setSetting(ss,'closedAt','');
}

// ============================================================
// v6 추가: 학급자율활동 홈 항목 ('활동항목' 시트)
// - 시트 탭은 절대 삭제하지 말고, 행만 지우거나 앱에서 수정하세요.
// - 헤더/값 모두 별도 셀. 전체 셀을 텍스트 서식으로 저장해 수식 오해석 방지.
// ============================================================
const HUB_HEADERS = ['ID','구분','유형','아이콘','이름','한줄설명','안내내용','링크','링크이름','표시','순서','제출마감','영역','마감일'];
const FORM_HEADERS = ['항목ID','질문ID','질문','유형','필수','안내문구','순서','선택지'];   // v10: 선택지(객관식, 한 줄에 하나)
const SUB_FIXED = ['이름','제출시각','학년','반','번호','시트명','구분'];   // 설문폼 시트와 같은 고정 열
const LOG_HEADERS = ['제출번호','제출시각','영역','주제ID','주제','반','번호','이름','구분'];
const AREAS = {
  auto:   { label: '자율', sheet: '제출_자율' },
  career: { label: '진로', sheet: '제출_진로' },
  etc:    { label: '기타', sheet: '제출_기타' }
};

function getHubSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_HUB);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_HUB);
    sheet.getRange(1,1,1,HUB_HEADERS.length).setValues([HUB_HEADERS]);
    sheet.getRange(1,1,1,HUB_HEADERS.length).setBackground('#9B8EC4').setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else if (sheet.getLastColumn() < HUB_HEADERS.length) {
    // v6 시트 → 새 열(제출마감) 헤더만 보강 (기존 데이터는 그대로)
    sheet.getRange(1,1,1,HUB_HEADERS.length).setValues([HUB_HEADERS]);
    sheet.getRange(1,1,1,HUB_HEADERS.length).setBackground('#9B8EC4').setFontColor('#FFFFFF').setFontWeight('bold');
  }
  return sheet;
}

function getFormsSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_FORMS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_FORMS);
    sheet.getRange(1,1,1,FORM_HEADERS.length).setValues([FORM_HEADERS]);
    sheet.getRange(1,1,1,FORM_HEADERS.length).setBackground('#9B8EC4').setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else if (sheet.getLastColumn() < FORM_HEADERS.length) {
    // v10: 기존 시트에 '선택지' 헤더만 보강 (기존 데이터는 그대로)
    sheet.getRange(1,1,1,FORM_HEADERS.length).setValues([FORM_HEADERS]);
    sheet.getRange(1,1,1,FORM_HEADERS.length).setBackground('#9B8EC4').setFontColor('#FFFFFF').setFontWeight('bold');
  }
  return sheet;
}

function hubRow(it, kind, order) {
  return [
    String(it.id || ''), kind, String(it.type || ''), String(it.icon || ''),
    String(it.name || ''), String(it.desc || ''), String(it.body || ''),
    String(it.link || ''), String(it.linkLabel || ''),
    it.visible === false ? 'X' : 'O', String(order),
    it.closed === true ? 'O' : 'X',
    String(it.area || ''), String(it.due || '')
  ];
}

function saveHubItems(ss, d) {
  // 관리자 암호가 설정되어 있으면 맞아야 저장돼요 (설정 안 했으면 이전처럼 저장)
  if (ADMIN_PW && String(d.pw || '') !== ADMIN_PW) {
    adminFail();   // v11: 틀린 암호 기록(+지연)
    return { status: 'auth', message: '관리자 암호가 달라요. 관리자 화면의 "관리자 암호"를 확인해 주세요.' };
  }
  const sheet = getHubSheet(ss);
  const rows = [];
  (d.items || []).forEach((it, i) => rows.push(hubRow(it, 'hub', i + 1)));
  (d.study || []).forEach((it, i) => rows.push(hubRow(it, 'study', i + 1)));
  const last = sheet.getLastRow();
  if (last > 1) sheet.getRange(2, 1, last - 1, HUB_HEADERS.length).clearContent();
  if (rows.length) {
    const range = sheet.getRange(2, 1, rows.length, HUB_HEADERS.length);
    range.setNumberFormat('@');
    range.setValues(rows);
  }
  if (Array.isArray(d.forms)) {
    const fs = getFormsSheet(ss);
    const frows = d.forms.map((q, i) => [
      String(q.itemId || ''), String(q.id || ''), String(q.label || ''), String(q.type || 'short'),
      q.required === true ? 'O' : 'X', String(q.hint || ''), String(i + 1),
      String(q.type || '') === 'choice' ? normOptions(q.options).join('\n') : ''
    ]);
    const flast = fs.getLastRow();
    if (flast > 1) fs.getRange(2, 1, flast - 1, FORM_HEADERS.length).clearContent();
    if (frows.length) {
      const fr = fs.getRange(2, 1, frows.length, FORM_HEADERS.length);
      fr.setNumberFormat('@');
      fr.setValues(frows);
    }
  }
  if (Array.isArray(d.files)) {
    const fsheet = getFilesSheet(ss);
    const per = {};
    const frows = [];
    d.files.forEach(f => {
      const fid = String(f.id || ''), iid = String(f.itemId || '');
      const kind = String(f.kind || '') === 'pdf' ? 'pdf' : 'image';
      if (!iid || !/^[A-Za-z0-9_-]{10,}$/.test(fid)) return;
      per[iid] = (per[iid] || 0) + 1;
      if (per[iid] > 12) return;
      frows.push([iid, fid, String(f.name || '').slice(0, 100), kind, String(frows.length + 1)]);
    });
    const lastF = fsheet.getLastRow();
    if (lastF > 1) fsheet.getRange(2, 1, lastF - 1, FILE_HEADERS.length).clearContent();
    if (frows.length) {
      const rg = fsheet.getRange(2, 1, frows.length, FILE_HEADERS.length);
      rg.setNumberFormat('@');
      rg.setValues(frows);
    }
  }
  setSetting(ss, 'hubVersion', String(d.saveId || ''));
  setSetting(ss, 'hubSavedAt', Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'));
  return { status: 'ok' };
}

function getHubItems(ss) {
  const out = { items: [], study: [], forms: [], files: [], choiceOk: true, version: getSetting(ss, 'hubVersion'), savedAt: getSetting(ss, 'hubSavedAt') };
  const flsheet = ss.getSheetByName(SHEET_FILES);
  if (flsheet && flsheet.getLastRow() > 1) {
    const ld = flsheet.getDataRange().getValues();
    const lh = ld[0];
    const lc = n => lh.indexOf(n);
    out.files = ld.slice(1)
      .filter(r => String(r[lc('파일ID')]) !== '')
      .map(r => ({
        itemId: String(r[lc('항목ID')]), id: String(r[lc('파일ID')]), name: String(r[lc('파일이름')]),
        kind: String(r[lc('종류')]) === 'pdf' ? 'pdf' : 'image', order: Number(r[lc('순서')]) || 0
      }))
      .sort((a, b) => a.order - b.order);
  }
  const fsheet = ss.getSheetByName(SHEET_FORMS);
  if (fsheet && fsheet.getLastRow() > 1) {
    const fd = fsheet.getDataRange().getValues();
    const fh = fd[0];
    const fc = n => fh.indexOf(n);
    out.forms = fd.slice(1)
      .filter(r => String(r[fc('질문')]) !== '')
      .map(r => ({
        itemId: String(r[fc('항목ID')]), id: String(r[fc('질문ID')]), label: String(r[fc('질문')]),
        type: String(r[fc('유형')]), required: String(r[fc('필수')]) === 'O',
        hint: String(r[fc('안내문구')]), order: Number(r[fc('순서')]) || 0,
        options: fc('선택지') >= 0 ? normOptions(r[fc('선택지')]) : []
      }))
      .sort((a, b) => a.order - b.order);
  }
  const sheet = ss.getSheetByName(SHEET_HUB);
  if (!sheet || sheet.getLastRow() < 2) return out;
  const data = sheet.getDataRange().getValues();
  const h = data[0];
  const c = n => h.indexOf(n);
  const list = data.slice(1)
    .filter(r => String(r[c('이름')]) !== '')
    .map(r => ({
      id: String(r[c('ID')]),
      kind: String(r[c('구분')]),
      type: String(r[c('유형')]),
      icon: String(r[c('아이콘')]),
      name: String(r[c('이름')]),
      desc: String(r[c('한줄설명')]),
      body: String(r[c('안내내용')]),
      link: String(r[c('링크')]),
      linkLabel: String(r[c('링크이름')]),
      visible: String(r[c('표시')]) !== 'X',
      closed: c('제출마감') >= 0 && String(r[c('제출마감')]) === 'O',
      area: inferArea(String(r[c('유형')]), String(r[c('이름')]), c('영역') >= 0 ? String(r[c('영역')]) : ''),
      due: c('마감일') >= 0 ? dueStr(r[c('마감일')]) : '',
      order: Number(r[c('순서')]) || 0
    }))
    .sort((a, b) => a.order - b.order);
  list.forEach(x => { if (x.kind === 'study') out.study.push(x); else out.items.push(x); });
  return out;
}


// ============================================================
// v7 추가: 제출형 항목 — 학생 제출 저장
// - 항목마다 '제출_항목명' 시트가 자동 생성돼요. (탭 이름 변경/삭제 금지)
// - 같은 반+번호가 다시 제출하면 그 행을 덮어써요 (제출시각은 유지, 수정시각만 갱신).
// - 학생 입력은 전부 텍스트 서식으로 저장해서 수식(=...)으로 해석되지 않아요.
// ============================================================
// v10: 객관식 선택지 정리 (배열 또는 줄바꿈 문자열 → 최대 12개, 항목당 80자)
function normOptions(v) {
  const a = Array.isArray(v) ? v : String(v == null ? '' : v).split('\n');
  return a.map(o => String(o).trim().slice(0, 80)).filter(o => o).slice(0, 12);
}

function normNum(x) {
  const n = parseInt(String(x == null ? '' : x).replace(/[^0-9]/g, ''), 10);
  return isNaN(n) ? String(x == null ? '' : x).trim() : String(n);
}

function nowKST() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
}

function inferArea(type, name, area) {
  if (area === 'auto' || area === 'career' || area === 'etc') return area;
  if (type === 'submit') return String(name || '').indexOf('진로') >= 0 ? 'career' : 'auto';
  return 'etc';
}

function dueStr(v) {
  const s = (v instanceof Date) ? Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd') : String(v == null ? '' : v).trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}

function todayKST() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
}

function findHubItem(ss, itemId) {
  const sheet = ss.getSheetByName(SHEET_HUB);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const data = sheet.getDataRange().getValues();
  const h = data[0];
  const c = n => h.indexOf(n);
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (String(r[c('ID')]) === String(itemId) && String(r[c('구분')]) === 'hub') {
      const type = String(r[c('유형')]), name = String(r[c('이름')]);
      const due = c('마감일') >= 0 ? dueStr(r[c('마감일')]) : '';
      const flag = c('제출마감') >= 0 && String(r[c('제출마감')]) === 'O';
      return {
        id: String(itemId), name: name, type: type,
        area: inferArea(type, name, c('영역') >= 0 ? String(r[c('영역')]) : ''),
        visible: String(r[c('표시')]) !== 'X',
        due: due,
        closed: flag || (due !== '' && todayKST() > due)   // 마감일은 그날 밤 12시까지
      };
    }
  }
  return null;
}

function styleHeader(sheet, n) {
  const r = sheet.getRange(1, 1, 1, n);
  r.setBackground('#9B8EC4'); r.setFontColor('#FFFFFF'); r.setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function getAreaSheet(ss, areaKey) {
  const info = AREAS[areaKey] || AREAS.etc;
  let sheet = ss.getSheetByName(info.sheet);
  if (!sheet) {
    sheet = ss.insertSheet(info.sheet);
    sheet.getRange(1, 1, 1, SUB_FIXED.length).setValues([SUB_FIXED]);
    styleHeader(sheet, SUB_FIXED.length);
  }
  return sheet;
}

function getLogSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_LOG);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_LOG);
    sheet.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
    styleHeader(sheet, LOG_HEADERS.length);
  }
  return sheet;
}

function appendText(sheet, values) {
  const rg = sheet.getRange(sheet.getLastRow() + 1, 1, 1, values.length);
  rg.setNumberFormat('@');
  rg.setValues([values]);
}

function submitActivity(ss, d) {
  const itemId = String(d.itemId || '');
  const cls = normNum(d.cls), num = normNum(d.num);
  const name = String(d.name || '').trim().slice(0, 30);
  if (!itemId || !cls || !num || !name) return { status: 'error', message: '반·번호·이름을 확인해 주세요.' };
  const answers = Array.isArray(d.answers) ? d.answers.slice(0, 30) : [];
  if (!answers.length) return { status: 'error', message: '제출할 내용이 없어요.' };

  const item = findHubItem(ss, itemId);
  if (!item || item.type !== 'submit' || !item.visible) return { status: 'error', message: '지금은 제출할 수 없는 항목이에요. 새로고침 후 다시 시도해 주세요.' };
  if (item.closed) return { status: 'closed', message: '이 활동은 제출이 마감되었어요. 선생님께 문의해 주세요.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    // v11: 핀 확인 (핀이 있으면 맞아야 하고, 없으면 새로 만들어요. 예전에 낸 학생은 이름이 같아야 해요)
    const gate = pinGate(ss, cls, num, name, d.pin, false);
    if (!gate.ok) return gate.res;
    const areaInfo = AREAS[item.area] || AREAS.etc;
    const sheet = getAreaSheet(ss, item.area);
    const lastCol = Math.max(sheet.getLastColumn(), SUB_FIXED.length);
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);

    // 질문 → "주제 · 질문" 열 (처음 보는 열은 오른쪽에 새로 추가)
    const used = {}, cells = {};
    answers.forEach(a => {
      const q = String((a && a.label) || '').trim() || '질문';
      let label = (item.name + ' · ' + q).slice(0, 400);
      while (used[label]) label = label + '(2)';
      used[label] = true;
      let idx = headers.indexOf(label);
      if (idx < 0) {
        headers.push(label);
        idx = headers.length - 1;
        if (idx + 1 > sheet.getMaxColumns()) sheet.insertColumnsAfter(sheet.getMaxColumns(), idx + 1 - sheet.getMaxColumns());
        sheet.getRange(1, idx + 1).setValue(label);
      }
      let v;
      if (a.type === 'check') v = (a.value === true || a.value === 'true') ? 'O' : 'X';
      else v = String(a.value == null ? '' : a.value).slice(0, a.type === 'long' ? 3000 : 200);
      cells[idx] = v;
    });

    // 같은 반+번호 학생의 행 찾기 (없으면 새 행)
    const iBan = headers.indexOf('반'), iNum = headers.indexOf('번호');
    const lastRow = sheet.getLastRow();
    let rowNum = -1, row = new Array(headers.length).fill('');
    if (lastRow > 1 && iBan >= 0 && iNum >= 0) {
      const vals = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
      for (let i = 0; i < vals.length; i++) {
        if (normNum(vals[i][iBan]) === cls && normNum(vals[i][iNum]) === num) {
          rowNum = i + 2;
          row = vals[i].map(v => String(v == null ? '' : v));
          while (row.length < headers.length) row.push('');
          break;
        }
      }
    }
    const now = nowKST();
    const updated = rowNum > 0;
    if (!updated) rowNum = lastRow + 1;
    const put = (n, v) => { const k = headers.indexOf(n); if (k >= 0) row[k] = v; };
    put('이름', name); put('제출시각', now); put('반', cls); put('번호', num);
    put('시트명', areaInfo.sheet); put('구분', areaInfo.label);
    Object.keys(cells).forEach(k => { row[Number(k)] = cells[k]; });   // 이 주제의 칸만 갱신, 다른 주제 칸은 그대로

    const range = sheet.getRange(rowNum, 1, 1, headers.length);
    range.setNumberFormat('@');
    range.setValues([row]);

    appendText(getLogSheet(ss), [String(d.subId || ''), now, areaInfo.label, itemId, item.name, cls, num, name, updated ? '수정' : '신규']);
    return { status: 'ok', updated: updated };
  } finally {
    lock.releaseLock();
  }
}

// 제출번호로 실제 기록됐는지 확인 (응답을 못 읽었을 때 앱이 다시 확인해요)
function checkSubmit(ss, p) {
  const sheet = ss.getSheetByName(SHEET_LOG);
  if (!sheet || sheet.getLastRow() < 2) return { found: false };
  const vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  const token = String(p.token || ''), itemId = String(p.itemId || '');
  for (let i = 0; i < vals.length; i++) {
    if (token && String(vals[i][0]) === token && String(vals[i][3]) === itemId) return { found: true };
  }
  return { found: false };
}

// ============================================================
// v8 추가: 관리자 '주제별 보기' — 선택한 주제의 제출 내용만 읽어 보내요 (관리자 암호 필수)
// ============================================================
function adminTopicView(ss, d) {
  if (!ADMIN_PW) return { status: 'auth', message: 'GAS 파일 맨 위 ADMIN_PW에 암호를 먼저 적고 "수정 배포"해 주세요.' };
  if (String(d.pw || '') !== ADMIN_PW) {
    adminFail();   // v11: 틀린 암호 기록(+지연)
    return { status: 'auth', message: '관리자 암호가 달라요.' };
  }
  const item = findHubItem(ss, String(d.itemId || ''));
  if (!item || item.type !== 'submit') return { status: 'error', message: '제출형 주제가 아니에요. 저장하기를 먼저 눌러 보세요.' };
  const out = { status: 'ok', item: { id: item.id, name: item.name, area: item.area, areaLabel: (AREAS[item.area] || AREAS.etc).label }, questions: [], rows: [] };

  const sheet = ss.getSheetByName((AREAS[item.area] || AREAS.etc).sheet);
  if (!sheet || sheet.getLastRow() < 2) return out;
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(String);
  const prefix = item.name + ' · ';
  const cols = [];
  headers.forEach((h, i) => { if (h.indexOf(prefix) === 0) { cols.push(i); out.questions.push(h.slice(prefix.length)); } });

  // 이 주제를 낸 학생(제출로그 기준)과 마지막 제출 시각
  const logged = {};
  const log = ss.getSheetByName(SHEET_LOG);
  if (log && log.getLastRow() > 1) {
    log.getRange(2, 1, log.getLastRow() - 1, 8).getValues().forEach(r => {
      if (String(r[3]) === item.id) logged[normNum(r[5]) + '-' + normNum(r[6])] = String(r[1]);
    });
  }
  const pins = pinKeyMap(ss);   // v11: 핀 설정 여부
  const iBan = headers.indexOf('반'), iNum = headers.indexOf('번호'), iName = headers.indexOf('이름');
  data.slice(1).forEach(r => {
    const key = normNum(r[iBan]) + '-' + normNum(r[iNum]);
    const answers = cols.map(i => String(r[i] == null ? '' : r[i]));
    if (!(key in logged) && !answers.some(a => a !== '')) return;
    out.rows.push({ ban: normNum(r[iBan]), num: normNum(r[iNum]), name: String(r[iName] || ''), at: logged[key] || '', answers: answers, pin: !!pins[key] });
  });
  out.rows.sort((a, b) => (Number(a.ban) - Number(b.ban)) || (Number(a.num) - Number(b.num)));
  return out;
}

function getFilesSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_FILES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_FILES);
    sheet.getRange(1, 1, 1, FILE_HEADERS.length).setValues([FILE_HEADERS]);
    sheet.getRange(1, 1, 1, FILE_HEADERS.length).setBackground('#9B8EC4').setFontColor('#FFFFFF').setFontWeight('bold');
  }
  return sheet;
}

// v9 추가: 관리자가 올린 PDF·사진을 내 구글드라이브 폴더에 저장하고 "링크가 있는 사람은 보기"로 공유해요
function uploadHubFile(ss, d) {
  if (!ADMIN_PW) return { status: 'auth', message: 'GAS 파일 맨 위 ADMIN_PW에 암호를 먼저 적고 "수정 배포"해 주세요.' };
  if (String(d.pw || '') !== ADMIN_PW) {
    adminFail();   // v11: 틀린 암호 기록(+지연)
    return { status: 'auth', message: '관리자 암호가 달라요.' };
  }
  const mime = String(d.mime || '');
  const allowed = { 'application/pdf': 'pdf', 'image/jpeg': 'image', 'image/png': 'image', 'image/webp': 'image' };
  if (!allowed[mime]) return { status: 'error', message: 'PDF 또는 사진(JPG·PNG·WEBP)만 올릴 수 있어요.' };
  const b64 = String(d.data || '');
  if (!b64) return { status: 'error', message: '파일 내용이 비어 있어요.' };
  if (b64.length > FILE_MAX_B64) return { status: 'error', message: '파일이 너무 커요. (PDF는 약 10MB까지)' };
  let bytes;
  try { bytes = Utilities.base64Decode(b64); } catch (e) { return { status: 'error', message: '파일을 읽지 못했어요.' }; }
  // 확장자만 속인 파일을 막기 위해 파일 첫 부분(매직 바이트)을 확인해요
  const u = i => (bytes[i] & 255);
  const magicOk = mime === 'application/pdf' ? (u(0) === 0x25 && u(1) === 0x50 && u(2) === 0x44 && u(3) === 0x46)
    : mime === 'image/jpeg' ? (u(0) === 0xFF && u(1) === 0xD8)
    : mime === 'image/png' ? (u(0) === 0x89 && u(1) === 0x50 && u(2) === 0x4E && u(3) === 0x47)
    : (u(0) === 0x52 && u(1) === 0x49 && u(2) === 0x46 && u(3) === 0x46);
  if (!magicOk) return { status: 'error', message: '파일 형식이 올바르지 않아요.' };
  const name = String(d.name || '첨부').replace(/[\\/:*?"<>|]/g, '_').slice(0, 100) || '첨부';
  const blob = Utilities.newBlob(bytes, mime, name);
  const it = DriveApp.getFoldersByName(FILES_FOLDER);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(FILES_FOLDER);
  const file = folder.createFile(blob);
  let shared = true;
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { shared = false; }
  return { status: 'ok', fileId: file.getId(), name: name, kind: allowed[mime], shared: shared };
}

// 주제별 제출 현황판: 주제마다 제출한 학생 수와 마지막 제출 시각 (제출로그 기준, 같은 학생의 수정 제출은 1명으로 계산)
function adminTopicSummary(ss, d) {
  if (!ADMIN_PW) return { status: 'auth', message: 'GAS 파일 맨 위 ADMIN_PW에 암호를 먼저 적고 "수정 배포"해 주세요.' };
  if (String(d.pw || '') !== ADMIN_PW) {
    adminFail();   // v11: 틀린 암호 기록(+지연)
    return { status: 'auth', message: '관리자 암호가 달라요.' };
  }
  const per = {};
  const log = ss.getSheetByName(SHEET_LOG);
  if (log && log.getLastRow() > 1) {
    log.getRange(2, 1, log.getLastRow() - 1, 8).getValues().forEach(r => {
      const id = String(r[3]);
      if (!id) return;
      if (!per[id]) per[id] = { students: {}, last: '' };
      per[id].students[normNum(r[5]) + '-' + normNum(r[6])] = true;
      const at = String(r[1]);
      if (at > per[id].last) per[id].last = at;
    });
  }
  const hubSheet = ss.getSheetByName(SHEET_HUB);
  const ids = [];
  if (hubSheet && hubSheet.getLastRow() > 1) {
    const hd = hubSheet.getDataRange().getValues();
    const ci = hd[0].indexOf('ID'), cg = hd[0].indexOf('구분');
    hd.slice(1).forEach(r => { if (String(r[cg]) === 'hub' && r[ci] !== '') ids.push(String(r[ci])); });
  }
  const items = ids.map(id => findHubItem(ss, id)).filter(x => x && x.type === 'submit').map(x => ({
    id: x.id, name: x.name, area: x.area, closed: !!x.closed,
    count: per[x.id] ? Object.keys(per[x.id].students).length : 0,
    last: per[x.id] ? per[x.id].last : ''
  }));
  return { status: 'ok', items: items };
}

// ============================================================
// v11 추가: 학생 핀 — 내 제출 불러오기 / 핀 초기화
// - '학생핀' 시트에는 핀 원문이 아니라 해시만 저장돼요. 시트 탭은 삭제하지 마세요. (초기화는 앱의 🔑 버튼으로)
// - 핀은 숫자 4~12자리. 5번 연속 틀리면 그 학생은 10분간 잠겨요.
// ============================================================
function getPinSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_PIN);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_PIN);
    sheet.getRange(1, 1, 1, PIN_HEADERS.length).setValues([PIN_HEADERS]);
    styleHeader(sheet, PIN_HEADERS.length);
  }
  return sheet;
}

function pinSalt(ss) {
  let s = getSetting(ss, 'pinSalt');
  if (!s) { s = Utilities.getUuid(); setSetting(ss, 'pinSalt', s); }
  return s;
}

function pinHash(ss, cls, num, pin) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pinSalt(ss) + '|' + cls + '|' + num + '|' + pin, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}

function normName(x) {
  return String(x == null ? '' : x).replace(/\s+/g, '');
}

function findPinRow(sheet, cls, num) {
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const vals = sheet.getRange(2, 1, last - 1, 3).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (normNum(vals[i][0]) === cls && normNum(vals[i][1]) === num) return { row: i + 2, hash: String(vals[i][2]) };
  }
  return null;
}

// 핀이 설정된 학생 목록 {'반-번호': true}
function pinKeyMap(ss) {
  const out = {};
  const sheet = ss.getSheetByName(SHEET_PIN);
  if (!sheet || sheet.getLastRow() < 2) return out;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(r => { out[normNum(r[0]) + '-' + normNum(r[1])] = true; });
  return out;
}

// 예전(핀 도입 전)에 제출한 기록이 있으면 그때 적은 이름을, 없으면 null
function findLegacyName(ss, cls, num) {
  let name = null;
  const log = ss.getSheetByName(SHEET_LOG);
  if (log && log.getLastRow() > 1) {
    log.getRange(2, 1, log.getLastRow() - 1, 8).getValues().forEach(r => {
      if (normNum(r[5]) === cls && normNum(r[6]) === num) name = String(r[7] == null ? '' : r[7]);
    });
  }
  if (name !== null) return name;
  Object.keys(AREAS).forEach(k => {
    const sheet = ss.getSheetByName(AREAS[k].sheet);
    if (!sheet || sheet.getLastRow() < 2) return;
    const data = sheet.getDataRange().getValues();
    const h = data[0].map(String);
    const iB = h.indexOf('반'), iN = h.indexOf('번호'), iNm = h.indexOf('이름');
    if (iB < 0 || iN < 0) return;
    for (let i = 1; i < data.length; i++) {
      if (normNum(data[i][iB]) === cls && normNum(data[i][iN]) === num) { name = iNm >= 0 ? String(data[i][iNm] == null ? '' : data[i][iNm]) : ''; }
    }
  });
  return name;
}

// 핀 확인·생성 (호출하는 쪽에서 LockService를 잡고 부르세요)
//  · 핀이 이미 있으면 맞아야 통과
//  · 핀이 없으면: 예전에 낸 기록이 있는 학생은 '이름이 같을 때만' 새 핀 생성, 기록이 없으면 새 핀 생성
//  · requireExisting=true(조회용)이면 제출 기록이 없는 학생은 핀을 만들지 않고 안내만 해요
function pinGate(ss, cls, num, name, pin, requireExisting) {
  pin = String(pin == null ? '' : pin).trim();
  const sheet = getPinSheet(ss);
  const found = findPinRow(sheet, cls, num);
  const cache = CacheService.getScriptCache();
  const key = 'pf_' + cls + '_' + num;
  if (found) {
    const fails = Number(cache.get(key) || 0);
    if (fails >= PIN_MAX_FAIL) return { ok: false, res: { status: 'pin', code: 'locked', message: '핀을 여러 번 틀려서 10분 동안 잠겼어요. 잠시 후 다시 하거나 선생님께 핀 초기화를 요청해 주세요.' } };
    if (!pin) return { ok: false, res: { status: 'pin', code: 'need', message: '핀을 입력해 주세요.' } };
    if (found.hash !== pinHash(ss, cls, num, pin)) {
      const n = fails + 1;
      cache.put(key, String(n), 600);
      Utilities.sleep(1000);
      return { ok: false, res: { status: 'pin', code: 'wrong', message: '핀이 달라요. (' + n + '/' + PIN_MAX_FAIL + ') 잊었다면 선생님께 초기화를 요청해 주세요.' } };
    }
    cache.remove(key);
    return { ok: true, created: false };
  }
  const legacy = findLegacyName(ss, cls, num);
  if (requireExisting && legacy === null) return { ok: false, res: { status: 'none', message: '아직 제출한 기록이 없어요. 처음이라면 활동 화면에서 핀을 만들고 제출해 주세요.' } };
  if (!/^\d{4,12}$/.test(pin)) return { ok: false, res: { status: 'pin', code: 'format', message: '내 핀을 새로 만들어 주세요. (숫자 4자리 이상, 최대 12자리)' } };
  if (legacy !== null && normName(legacy) !== '' && normName(legacy) !== normName(name)) {
    return { ok: false, res: { status: 'pin', code: 'legacy_name', message: '예전에 제출한 기록의 이름과 달라요. 처음 제출할 때 쓴 이름으로 입력해 주세요.' } };
  }
  appendText(sheet, [cls, num, pinHash(ss, cls, num, pin), nowKST()]);
  return { ok: true, created: true };
}

function pinStatus(ss, d) {
  const cls = normNum(d.cls), num = normNum(d.num);
  if (!cls || !num) return { status: 'error', message: '반·번호를 확인해 주세요.' };
  const sheet = ss.getSheetByName(SHEET_PIN);
  return { status: 'ok', hasPin: !!(sheet && findPinRow(sheet, cls, num)) };
}

// 학생이 자기 제출 내용 전부 불러오기 (반+번호+핀 필요. 핀 없는 예전 제출 학생은 이름이 같으면 핀을 새로 만들어요)
function getMySubmission(ss, d) {
  const cls = normNum(d.cls), num = normNum(d.num);
  const name = String(d.name || '').trim().slice(0, 30);
  if (!cls || !num || !name) return { status: 'error', message: '반·번호·이름을 확인해 주세요.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let gate;
  try { gate = pinGate(ss, cls, num, name, d.pin, true); } finally { lock.releaseLock(); }
  if (!gate.ok) return gate.res;

  const logAt = {};
  const log = ss.getSheetByName(SHEET_LOG);
  if (log && log.getLastRow() > 1) {
    log.getRange(2, 1, log.getLastRow() - 1, 8).getValues().forEach(r => {
      if (normNum(r[5]) === cls && normNum(r[6]) === num) logAt[String(r[3])] = String(r[1]);
    });
  }
  const hubSheet = ss.getSheetByName(SHEET_HUB);
  const ids = [];
  if (hubSheet && hubSheet.getLastRow() > 1) {
    const hd = hubSheet.getDataRange().getValues();
    const ci = hd[0].indexOf('ID'), cg = hd[0].indexOf('구분');
    hd.slice(1).forEach(r => { if (String(r[cg]) === 'hub' && r[ci] !== '') ids.push(String(r[ci])); });
  }
  const cache = {};
  const items = [];
  ids.forEach(id => {
    const it = findHubItem(ss, id);
    if (!it || it.type !== 'submit') return;
    if (!(it.area in cache)) {
      const sh = ss.getSheetByName((AREAS[it.area] || AREAS.etc).sheet);
      cache[it.area] = (sh && sh.getLastRow() > 1) ? sh.getDataRange().getValues() : null;
    }
    const data = cache[it.area];
    const qa = [];
    if (data) {
      const headers = data[0].map(String);
      const iB = headers.indexOf('반'), iN = headers.indexOf('번호');
      const prefix = it.name + ' · ';
      for (let i = 1; i < data.length; i++) {
        if (normNum(data[i][iB]) === cls && normNum(data[i][iN]) === num) {
          headers.forEach((h, c) => {
            if (h.indexOf(prefix) === 0) {
              const v = String(data[i][c] == null ? '' : data[i][c]);
              if (v !== '') qa.push({ label: h.slice(prefix.length), value: v });
            }
          });
          break;
        }
      }
    }
    if (!qa.length && !logAt[id]) return;
    items.push({ itemId: id, at: logAt[id] || '', qa: qa });
  });
  return { status: 'ok', created: !!gate.created, items: items };
}

// 관리자: 학생 핀 초기화 (학생이 다음 제출·조회 때 새 핀을 만들 수 있어요)
function adminResetPin(ss, d) {
  if (!ADMIN_PW) return { status: 'auth', message: 'GAS 파일 맨 위 ADMIN_PW에 암호를 먼저 적고 "수정 배포"해 주세요.' };
  if (String(d.pw || '') !== ADMIN_PW) {
    adminFail();   // v11: 틀린 암호 기록(+지연)
    return { status: 'auth', message: '관리자 암호가 달라요.' };
  }
  const cls = normNum(d.cls), num = normNum(d.num);
  if (!cls || !num) return { status: 'error', message: '반·번호를 확인해 주세요.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getPinSheet(ss);
    const found = findPinRow(sheet, cls, num);
    if (found) sheet.deleteRow(found.row);
    CacheService.getScriptCache().remove('pf_' + cls + '_' + num);
    return { status: 'ok', reset: !!found };
  } finally {
    lock.releaseLock();
  }
}

// v11: 앱 관리자 로그인 확인 — 암호를 HTML에 두지 않고 여기(GAS)에서만 확인해요
function adminFail() {
  const c = CacheService.getScriptCache();
  c.put('adm_fail', String(Number(c.get('adm_fail') || 0) + 1), 600);
  Utilities.sleep(1500);
}
function adminLockedOut() {
  return Number(CacheService.getScriptCache().get('adm_fail') || 0) >= ADMIN_MAX_FAIL;
}
function adminLogin(d) {
  if (!ADMIN_PW) return { status: 'auth', message: 'GAS 파일 맨 위 ADMIN_PW에 암호를 먼저 적고 "수정 배포"해 주세요.' };
  if (String(d.pw || '') !== ADMIN_PW) { adminFail(); return { status: 'auth', message: '관리자 암호가 달라요.' }; }
  CacheService.getScriptCache().remove('adm_fail');
  return { status: 'ok' };
}

function doGet(e) {
  const params = e.parameter;
  const action = params.action || '';
  let result;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === 'getStatus') {
      result = {
        isClosed: getSetting(ss,'isClosed')==='true',
        closedAt: getSetting(ss,'closedAt')
      };

    // ★ 신규: 이름+반+번호만으로 일반스터디 + 왕멘토 전체 기록을 한 번에 반환
    } else if (action === 'getAllMyRecords') {
      let records = [];
      const memberSheet = ss.getSheetByName(SHEET_MEMBER);
      if (memberSheet && memberSheet.getLastRow() > 1) {
        const data = memberSheet.getDataRange().getValues();
        const h = data[0];
        records = records.concat(
          data.slice(1).filter(row=>
            String(row[h.indexOf('이름')])==String(params.name)&&
            String(row[h.indexOf('반')])==String(params.classNum)&&
            String(row[h.indexOf('번호')])==String(params.studentNum)
          ).map(row=>({
            type:'member',
            round:row[h.indexOf('회차')],
            quarter:row[h.indexOf('분기')],
            studyName:row[h.indexOf('스터디명')],
            date:row[h.indexOf('활동날짜')],
            time:row[h.indexOf('활동시간')],
            totalMinutes:row[h.indexOf('누적시간(분)')]||0
          }))
        );
      }
      const kingSheet = ss.getSheetByName(SHEET_KING);
      if (kingSheet && kingSheet.getLastRow() > 1) {
        const data = kingSheet.getDataRange().getValues();
        const h = data[0];
        records = records.concat(
          data.slice(1).filter(row=>
            String(row[h.indexOf('이름')])==String(params.name)&&
            String(row[h.indexOf('반')])==String(params.classNum)&&
            String(row[h.indexOf('번호')])==String(params.studentNum)
          ).map(row=>({
            type:'king',
            round:row[h.indexOf('회차')],
            quarter:row[h.indexOf('분기')],
            subject:row[h.indexOf('담당과목')],
            date:row[h.indexOf('활동날짜')],
            time:row[h.indexOf('활동시간')],
            totalMinutes:row[h.indexOf('누적시간(분)')]||0
          }))
        );
      }
      result = {records};

    // ★ 수정: quarter=all 이면 분기 필터 없이 전체 반환, studyName 생략 가능
    } else if (action === 'getMyRecords') {
      const sheet = ss.getSheetByName(SHEET_MEMBER);
      if(!sheet){result={records:[]};} else {
        const data=sheet.getDataRange().getValues();const h=data[0];
        const allQuarters = String(params.quarter)==='all';
        const allStudy   = !params.studyName || String(params.studyName)==='all';
        const records=data.slice(1).filter(row=>
          String(row[h.indexOf('이름')])==String(params.name)&&
          String(row[h.indexOf('반')])==String(params.classNum)&&
          String(row[h.indexOf('번호')])==String(params.studentNum)&&
          (allQuarters || String(row[h.indexOf('분기')])==String(params.quarter))&&
          (allStudy    || String(row[h.indexOf('스터디명')])==String(params.studyName))
        ).map(row=>({
          type:'member',
          round:row[h.indexOf('회차')],
          quarter:row[h.indexOf('분기')],
          studyName:row[h.indexOf('스터디명')],
          date:row[h.indexOf('활동날짜')],
          time:row[h.indexOf('활동시간')],
          totalMinutes:row[h.indexOf('누적시간(분)')]||0
        }));
        result={records};
      }

    } else if (action === 'getMyRecordDetail') {
      const sheet = ss.getSheetByName(SHEET_MEMBER);
      if(!sheet){result={record:null};} else {
        const data=sheet.getDataRange().getValues();const h=data[0];
        const found=data.slice(1).find(row=>
          String(row[h.indexOf('이름')])==String(params.name)&&
          String(row[h.indexOf('반')])==String(params.classNum)&&
          String(row[h.indexOf('번호')])==String(params.studentNum)&&
          String(row[h.indexOf('분기')])==String(params.quarter)&&
          String(row[h.indexOf('스터디명')])==String(params.studyName)&&
          String(row[h.indexOf('회차')])==String(params.round)
        );
        result={record:found?{
          round:found[h.indexOf('회차')],memberRole:found[h.indexOf('멘토/멘티')],
          date:found[h.indexOf('활동날짜')],time:found[h.indexOf('활동시간')],
          place:found[h.indexOf('활동장소')],title:found[h.indexOf('활동제목')],
          content:found[h.indexOf('활동내용')],reflection:found[h.indexOf('느낀점')],
          nextPlan:found[h.indexOf('다음회차예고')],present:found[h.indexOf('출석')],
          absent:found[h.indexOf('결석')],totalMinutes:found[h.indexOf('누적시간(분)')]||0,
          photoUrl:found[h.indexOf('사진첨부')]
        }:null};
      }

    } else if (action === 'getKingRecordDetail') {
      const sheet = ss.getSheetByName(SHEET_KING);
      if(!sheet){result={record:null};} else {
        const data=sheet.getDataRange().getValues();const h=data[0];
        const found=data.slice(1).find(row=>
          String(row[h.indexOf('이름')])==String(params.name)&&
          String(row[h.indexOf('반')])==String(params.classNum)&&
          String(row[h.indexOf('번호')])==String(params.studentNum)&&
          String(row[h.indexOf('분기')])==String(params.quarter)&&
          String(row[h.indexOf('담당과목')])==String(params.subject)&&
          String(row[h.indexOf('회차')])==String(params.round)
        );
        result={record:found?{
          round:found[h.indexOf('회차')],date:found[h.indexOf('활동날짜')],
          time:found[h.indexOf('활동시간')],selfStudy:found[h.indexOf('학습내용')],
          reflection:found[h.indexOf('느낀점')],nextPlan:found[h.indexOf('다음활동예고')],
          questions:found[h.indexOf('질문처리내역(JSON)')],
          totalMinutes:found[h.indexOf('누적시간(분)')]||0,photoUrl:found[h.indexOf('사진첨부')]
        }:null};
      }

    // ★ 수정: quarter=all 이면 분기 필터 없이 전체 반환, subject 생략 가능
    } else if (action === 'getKingRecords') {
      const sheet = ss.getSheetByName(SHEET_KING);
      if(!sheet){result={records:[]};} else {
        const data=sheet.getDataRange().getValues();const h=data[0];
        const allQuarters = String(params.quarter)==='all';
        const allSubject  = !params.subject || String(params.subject)==='all';
        const records=data.slice(1).filter(row=>
          String(row[h.indexOf('이름')])==String(params.name)&&
          String(row[h.indexOf('반')])==String(params.classNum)&&
          String(row[h.indexOf('번호')])==String(params.studentNum)&&
          (allQuarters || String(row[h.indexOf('분기')])==String(params.quarter))&&
          (allSubject  || String(row[h.indexOf('담당과목')])==String(params.subject))
        ).map(row=>({
          type:'king',
          round:row[h.indexOf('회차')],
          quarter:row[h.indexOf('분기')],
          subject:row[h.indexOf('담당과목')],
          date:row[h.indexOf('활동날짜')],
          time:row[h.indexOf('활동시간')],
          totalMinutes:row[h.indexOf('누적시간(분)')]||0
        }));
        result={records};
      }

    } else if (action === 'getNotice') {
      const sheet=ss.getSheetByName(SHEET_NOTICE);
      if(!sheet){result={notice:''};} else {
        const data=sheet.getDataRange().getValues();
        const found=data.slice(1).find(row=>String(row[0])==String(params.quarter)&&String(row[1])==String(params.studyName));
        result={notice:found?found[2]:''};
      }

    } else if (action === 'getPlanRecords') {
      const sheet = ss.getSheetByName(SHEET_PLAN);
      if (!sheet || sheet.getLastRow() < 2) { result = {records:[]}; } else {
        const data = sheet.getDataRange().getValues(); const h = data[0];
        result = { records: data.slice(1).map(row => ({
          timestamp: row[h.indexOf('제출시각')],
          quarter: row[h.indexOf('분기')],
          mentorName: row[h.indexOf('멘토이름')],
          classNum: row[h.indexOf('반')],
          studentNum: row[h.indexOf('번호')],
          subject: row[h.indexOf('과목')],
          teamName: row[h.indexOf('팀명')],
          members: row[h.indexOf('팀원명단')],
          timeSlots: row[h.indexOf('스터디시간대')],
          goalTime: row[h.indexOf('목표시간')],
          content: row[h.indexOf('계획내용')]
        }))};
      }

    } else if (action === 'getRecords') {
      const sheet=ss.getSheetByName(SHEET_MEMBER);
      const kingSheet=ss.getSheetByName(SHEET_KING);
      let records=[];
      if(sheet&&sheet.getLastRow()>1){
        const data=sheet.getDataRange().getValues();const h=data[0];
        records=records.concat(data.slice(1).map(row=>({
          type:'member',quarter:row[h.indexOf('분기')],studyName:row[h.indexOf('스터디명')],
          round:row[h.indexOf('회차')],name:row[h.indexOf('이름')],
          classNum:row[h.indexOf('반')],studentNum:row[h.indexOf('번호')],date:row[h.indexOf('활동날짜')]
        })));
      }
      if(kingSheet&&kingSheet.getLastRow()>1){
        const data=kingSheet.getDataRange().getValues();const h=data[0];
        records=records.concat(data.slice(1).map(row=>({
          type:'king',quarter:row[h.indexOf('분기')],subject:row[h.indexOf('담당과목')],
          round:row[h.indexOf('회차')],name:row[h.indexOf('이름')],
          classNum:row[h.indexOf('반')],studentNum:row[h.indexOf('번호')],date:row[h.indexOf('활동날짜')]
        })));
      }
      result={records};

    } else if (action === 'getHubItems') {
      result = getHubItems(ss);

    } else if (action === 'checkSubmit') {
      result = checkSubmit(ss, e.parameter);

    } else {
      result={status:'ok',message:'자율스터디 GAS v11 정상 작동 중'};
    }

  } catch(err) {
    result={status:'error',message:err.toString()};
  }

  const callback=params.callback;
  return callback
    ?ContentService.createTextOutput(`${callback}(${JSON.stringify(result)})`).setMimeType(ContentService.MimeType.JAVASCRIPT)
    :ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}
