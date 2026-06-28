/**
 * ============================================================
 * © 2026 GEG화성 (깊이e끌림). All rights reserved.
 *
 * 본 코드는 「저작권법」의 보호받는 저작물입니다.
 * - 복제권(제 16조)·공중송신권(제 18조)·배포권(제 20조)은
 *   저작권자에게 있습니다.
 * - 정 경로로 받은 이용자라도 코드의 무단 복제·재배포·
 *   재판매·리브랜딩은 허용되지 않습니다.
 * - 무단 이용 시 「저작권법」 제 136조(5년 이하 징역 또는
 *   5천만 원 이하 벌금) 및 제 125조(손해배상) 적용 대상이
 *   될 수 있습니다.
 * - 이용 문의: bacusiki777@gmail.com, for2102@jimj.kr
 * ============================================================
 */

// 빌드 서명
const _BUILD_SIG = 'GEGHS-DEEPE-2026';

// 출처 확인용 함수
function getBuildInfo() {
  return {
    sig: _BUILD_SIG,
    owner: 'GEG화성 (깊이e끌림)',
    year: 2026
  };
}

/***********************************************************************
 * SDG's 리틀 히어로 — 선생님 시트 연동 스크립트 (Code.gs)
 * ---------------------------------------------------------------------
 * 이 스크립트는 "선생님 본인의 스프레드시트"에 들어 있는 코드입니다.
 * 자기 시트만 읽고 쓰므로, 다른 사람의 데이터에는 접근하지 않습니다.
 *
 * [선생님이 한 번만 하는 일]
 *   1) 이 시트를 "사본 만들기" 해서 내 드라이브에 가져온다.
 *   2) "학생 명단" 탭에 번호(A열)·이름(B열)을 적는다.
 *   3) 상단 메뉴 ▸ 확장 프로그램 ▸ Apps Script ▸ 배포 ▸ 새 배포
 *        ▸ 유형 "웹 앱" ▸ 실행 "나" ▸ 액세스 "모든 사용자" ▸ 배포
 *        ▸ (처음이면) 권한 검토 ▸ 고급 ▸ "(안전하지 않음)으로 이동"
 *   4) 나온 "웹 앱 URL"(끝이 /exec)을 복사해 앱의 [선생님 설정]에 붙여넣는다.
 *
 * ※ 코드를 수정한 경우에는 [배포 ▸ 배포 관리 ▸ 편집(연필) ▸ 버전 "새 버전" ▸ 배포]
 *   를 해야 변경이 반영됩니다.
 ***********************************************************************/

// ── 시트/폴더 이름 (원하면 바꿔도 됩니다) ──────────────────────────
var ROSTER_SHEET = '학생 명단';        // 기본 명단 탭 이름
// 아래 후보 중 "먼저 있는" 탭을 학생 명단으로 사용한다. (이미 만들어 둔 탭과 호환)
var ROSTER_ALIASES = ['학생 명단', '학생명단', '명단'];
var SUBMIT_SHEET = '실천 기록';  // 실천 기록이 쌓이는 탭 이름 (없으면 자동 생성)
var SUBMIT_ALIASES = ['실천 기록', '실천 기록 항목', '실천기록', '제출'];
var FOLDER_PROP_KEY = 'submitFolderId';

// 앱(index.html)의 [선생님 설정] 비밀번호와 똑같이 적어주세요.
// 이 값이 '사용 설명' 탭에 안내로 표시됩니다.
var TEACHER_PASSCODE = '1234';

// ── 제출 탭의 열 순서 ────────────────────────────────────────────
var SUBMIT_HEADERS = [
  '기록시각', '이름', '달성률(%)', '실천 내용', '오늘의 한마디', '나의 다짐', '미션설정(JSON)', '시작일', '종료일'
];

/*=====================================================================
 * 메뉴 & 초기 셋업
 *====================================================================*/
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🌱 SDG 리틀 히어로')
    .addItem('대시보드 만들기', 'createDashboardTab')
    .addItem('사용 설명 만들기', 'createGuideTab')
    .addItem('명단과 기록 관련 만들기', 'createRosterAndSubmitTabs')
    .addToUi();
}

function createDashboardTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureDashboardSheet_(ss);
  SpreadsheetApp.getUi().alert('완료', '대시보드 탭이 준비되었습니다.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function createGuideTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureGuideSheet_(ss);
  SpreadsheetApp.getUi().alert('완료', '사용 설명 탭이 준비되었습니다.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function createRosterAndSubmitTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var roster = getRosterSheet_(ss);
  if (!roster) {
    roster = ss.insertSheet(ROSTER_SHEET, 1);
    roster.getRange(1, 1, 1, 3).setValues([['번호', '이름', '팀(선택)']]);
    roster.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#E8EAF6').setFontColor('#1A237E');
    roster.setFrozenRows(1);
    roster.setColumnWidth(2, 140);
  }

  ensureSubmitSheet_(ss);

  SpreadsheetApp.getUi().alert(
    '완료',
    '학생 명단 탭과 실천 기록 탭이 준비되었습니다.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function ensureDashboardSheet_(ss) {
  var sh = ss.getSheetByName('대시보드');
  if (!sh) {
    sh = ss.insertSheet('대시보드', 0);
  } else {
    sh.activate();
  }
  sh.clear();
  sh.getRange(1, 1).setValue('📊 우리 반 통계 대시보드').setFontSize(16).setFontWeight('bold');
  sh.getRange(3, 1).setValue('※ 현재 학생들의 상세한 누적 발자취 및 통계 현황은 앱(웹) 화면에서 더 직관적으로 확인하실 수 있습니다.').setFontSize(11);
  return sh;
}

// 명단으로 쓸 탭을 후보 이름 순서대로 찾는다. (없으면 null)
function getRosterSheet_(ss) {
  for (var i = 0; i < ROSTER_ALIASES.length; i++) {
    var s = ss.getSheetByName(ROSTER_ALIASES[i]);
    if (s) return s;
  }
  return null;
}

/*=====================================================================
 * 웹 요청 처리 (앱이 호출)
 *====================================================================*/
function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || '';
    if (action === 'getRoster' || action === 'getStudents') {
      return json_({ ok: true, students: getRoster_() });
    }
    if (action === 'getHistory') {
      var name = (e && e.parameter && e.parameter.name) || '';
      return json_({ ok: true, history: getHistory_(name) });
    }
    if (action === 'getClassStats') {
      return json_({ ok: true, stats: getClassStats_() });
    }
    if (action === 'ping') {
      return json_({ ok: true, pong: true });
    }
    return json_({ ok: false, error: '알 수 없는 요청: ' + action });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }
    var action = body.action || '';

    if (action === 'uploadImage') {
      return json_(uploadImage_(body));
    }
    if (action === 'saveRecord') {
      return json_(saveRecord_(body));
    }
    if (action === 'saveReport') {
      return json_(saveReport_(body));
    }
    return json_({ ok: false, error: '알 수 없는 요청: ' + action });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/*=====================================================================
 * 기능 구현
 *====================================================================*/

// 명단 읽기 → [{id, name, team}]
function getRoster_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getRosterSheet_(ss);
  if (!sh) throw new Error('“학생 명단” 탭이 없어요. 메뉴 ▸ 🌱 SDG 리틀 히어로 ▸ 초기 설정을 눌러 주세요.');

  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  // 1행(헤더)에서 번호/이름/팀 열 위치 찾기 (열 순서가 달라도 됨)
  var header = values[0].map(function (h) { return String(h).replace(/\s/g, ''); });
  var col = function (keys) {
    for (var i = 0; i < header.length; i++) {
      for (var k = 0; k < keys.length; k++) {
        if (header[i].indexOf(keys[k]) >= 0) return i;
      }
    }
    return -1;
  };
  var iId = col(['번호', '학번', 'id', 'ID']);
  var iName = col(['이름', '성명', 'name']);
  var iTeam = col(['팀', '모둠', '조', 'team']);
  if (iName < 0) iName = 1; // 못 찾으면 두 번째 열을 이름으로 가정

  var out = [];
  for (var r = 1; r < values.length; r++) {
    var name = iName >= 0 ? String(values[r][iName]).trim() : '';
    if (!name) continue; // 이름 없는 행은 건너뜀
    out.push({
      id: iId >= 0 ? String(values[r][iId]).trim() : '',
      name: name,
      team: iTeam >= 0 ? String(values[r][iTeam]).trim() : ''
    });
  }
  return out;
}

// base64 파일을 드라이브에 저장 → {ok, id, url}
function uploadImage_(p) {
  if (!p.dataBase64) throw new Error('파일 데이터가 없습니다.');
  var mime = p.mimeType || 'application/octet-stream';
  var name = sanitize_(p.filename || 'upload');
  if (mime === 'application/pdf' && !/\.pdf$/i.test(name)) name += '.pdf';
  if (mime === 'image/png' && !/\.png$/i.test(name)) name += '.png';

  var bytes = Utilities.base64Decode(p.dataBase64);
  var blob = Utilities.newBlob(bytes, mime, name);

  var folder = getSubmitFolder_();
  var file = folder.createFile(blob);
  // 링크가 있는 사람은 보기 가능 → 시트의 IMAGE()/링크가 정상 표시됨
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}

  var id = file.getId();
  var url = (mime === 'application/pdf')
    ? file.getUrl()
    : ('https://drive.google.com/uc?export=view&id=' + id);
  return { ok: true, id: id, url: url };
}

// 제출 기록을 "제출" 탭에 한 줄 추가 → {ok}
function saveReport_(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ensureSubmitSheet_(ss);

  var row = [
    new Date(),
    p.studentId || '',
    p.studentName || '',
    p.team || '',
    p.mode || '',
    p.activityType || '',
    p.title || '',
    p.content || '',
    '', // 보고서(PDF) — 아래에서 하이퍼링크 수식으로 채움
    ''  // 그래프 — 아래에서 IMAGE 수식으로 채움
  ];
  sh.appendRow(row);
  var r = sh.getLastRow();

  // 보고서 PDF: 클릭하면 열리는 링크
  if (p.reportUrl) {
    sh.getRange(r, 9).setFormula('=HYPERLINK("' + p.reportUrl + '","PDF 열기")');
  }

  // 그래프 이미지: 여러 장이면 9열(그래프) 오른쪽으로 한 칸씩 펼쳐서 표시
  var ids = p.graphImageIds || [];
  for (var i = 0; i < ids.length; i++) {
    var imgUrl = 'https://drive.google.com/uc?export=view&id=' + ids[i];
    sh.getRange(r, 10 + i).setFormula('=IMAGE("' + imgUrl + '")');
  }
  if (ids.length > 0) sh.setRowHeight(r, 120);

  return { ok: true };
}

// SDG 히어로 실천 기록을 '실천 기록' 탭에 한 줄 추가 → {ok}
function saveRecord_(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ensureSubmitSheet_(ss);

  var tasks = p.tasks || '';
  if (Array.isArray(tasks)) {
    tasks = tasks.map(function(t) {
      return (t.theme || '') + ', ' + (t.task || '') + ', ' + (t.score || 0);
    }).join('\n');
  }

  var row = [
    new Date(),
    p.heroName || '',
    Number(p.progress) || 0,
    String(tasks),
    p.reflection || '',
    p.myGoal || '',
    p.missionsJson || '',  // 미션설정(JSON) - 상태 복원용
    p.startDate || '',     // 시작일
    p.endDate || ''        // 종료일
  ];
  sh.appendRow(row);
  return { ok: true };
}

// '실천 기록' 탭(없으면 '제출' 탭)에서 특정 학생의 기록을 반환 (최신 순)
function getHistory_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shInfo = getRecordSheet_(ss);
  if (!shInfo) return [];
  var sh = shInfo.sheet;
  var legacy = shInfo.legacy;

  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  // 헤더 기반 컬럼 자동 매핑
  var header = values[0].map(function(h) { return String(h).trim(); });
  var colTs   = header.indexOf('기록시각') >= 0 ? header.indexOf('기록시각') : 0;
  var colName = header.indexOf('이름')     >= 0 ? header.indexOf('이름')     : (legacy ? 2 : 1);
  var colProg = header.indexOf('달성률(%)') >= 0 ? header.indexOf('달성률(%)') : (legacy ? -1 : 2);
  var colTask = header.indexOf('실천 내용') >= 0 ? header.indexOf('실천 내용') : (legacy ? 7 : 3);
  var colRef  = header.indexOf('오늘의 한마디') >= 0 ? header.indexOf('오늘의 한마디') : 4;
  var colGoal = header.indexOf('나의 다짐') >= 0 ? header.indexOf('나의 다짐') : 5;
  var colMissions = header.indexOf('미션설정(JSON)') >= 0 ? header.indexOf('미션설정(JSON)') : -1;
  var colStart    = header.indexOf('시작일') >= 0 ? header.indexOf('시작일') : -1;
  var colEnd      = header.indexOf('종료일') >= 0 ? header.indexOf('종료일') : -1;

  var cleanReqName = String(name || '').replace(/^\d+\.\s*/, '').trim();

  var out = [];
  for (var r = 1; r < values.length; r++) {
    var rowNameRaw = String(values[r][colName] || '').trim();
    var rowNameClean = rowNameRaw.replace(/^\d+\.\s*/, '').trim();
    
    if (!cleanReqName || rowNameRaw === name || rowNameClean === cleanReqName) {
      var ts = values[r][colTs];
      out.push({
        timestamp: (ts instanceof Date) ? ts.toLocaleString('ko-KR') : String(ts),
        heroName: rowNameClean,
        progress: colProg >= 0 ? (Number(values[r][colProg]) || 0) : 0,
        tasks: colTask >= 0 ? String(values[r][colTask] || '') : '',
        reflection: colRef < values[r].length ? String(values[r][colRef] || '') : '',
        myGoal: colGoal < values[r].length ? String(values[r][colGoal] || '') : '',
        missionsJson: colMissions >= 0 ? String(values[r][colMissions] || '') : '',
        startDate: colStart >= 0 ? String(values[r][colStart] || '') : '',
        endDate: colEnd >= 0 ? String(values[r][colEnd] || '') : ''
      });
    }
  }
  return out.reverse(); // 최신 순
}

// 실천 기록 탭을 취득 (SUBMIT_SHEET 및 SUBMIT_ALIASES 시도)
function getRecordSheet_(ss) {
  for (var i = 0; i < SUBMIT_ALIASES.length; i++) {
    var sh = ss.getSheetByName(SUBMIT_ALIASES[i]);
    if (sh) {
      // '제출' 등 구형 탭이면 legacy: true
      var isLegacy = (SUBMIT_ALIASES[i] === '제출');
      return { sheet: sh, legacy: isLegacy };
    }
  }
  return null;
}

// 우리 반 전체 현황 통계 집계 → stats 객체
function getClassStats_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shInfo = getRecordSheet_(ss);
  if (!shInfo) return { totalStudents: 0, totalRecords: 0, totalTasks: 0, avgProgress: 0, dailyStats: [] };
  var sh = shInfo.sheet;
  var legacy = shInfo.legacy;

  var values = sh.getDataRange().getValues();
  if (values.length < 2) return { totalStudents: 0, totalRecords: 0, totalTasks: 0, avgProgress: 0, dailyStats: [] };

  var nameSet = {};
  var totalRecords = 0;
  var totalTasks = 0;
  var progressSum = 0;
  var dailyMap = {};

  // 코럼 인덱스: 간단하게 헤더로 스마트하게 매핑
  // 새 탭: ['기록시각', '이름', '달성률(%)', '실천 내용', '오늘의 한마디', '나의 다짐']
  // 구 탭: ['제출시각', '번호', '이름', '팀', ...]
  var header = values[0].map(function(h) { return String(h).trim(); });
  var colTs   = header.indexOf('기록시각') >= 0 ? header.indexOf('기록시각') : 0;
  var colName = header.indexOf('이름') >= 0 ? header.indexOf('이름') : (legacy ? 2 : 1);
  var colProg = header.indexOf('달성률(%)') >= 0 ? header.indexOf('달성률(%)') : (legacy ? -1 : 2);
  var colTask = header.indexOf('실천 내용') >= 0 ? header.indexOf('실천 내용') : (legacy ? 7 : 3);

  for (var r = 1; r < values.length; r++) {
    var ts = values[r][colTs];
    var nameRaw = String(values[r][colName] || '').trim();
    var nameClean = nameRaw.replace(/^\d+\.\s*/, '').trim();
    var progress = colProg >= 0 ? (Number(values[r][colProg]) || 0) : 0;
    var tasks = colTask >= 0 ? String(values[r][colTask] || '') : '';

    if (!nameClean) continue;
    nameSet[nameClean] = true;
    totalRecords++;
    progressSum += progress;

    var taskLines = tasks.split('\n').filter(function(l) { return l.trim(); });
    if (taskLines.length > 0) totalTasks += taskLines.length;
    else if (tasks.trim()) totalTasks++;

    var dateKey = '';
    if (ts instanceof Date) {
      dateKey = (ts.getMonth() + 1) + '/' + ts.getDate();
    } else {
      var d = new Date(ts);
      if (!isNaN(d)) dateKey = (d.getMonth() + 1) + '/' + d.getDate();
    }
    if (dateKey) {
      if (!dailyMap[dateKey]) dailyMap[dateKey] = { date: dateKey, records: 0, progressSum: 0 };
      dailyMap[dateKey].records++;
      dailyMap[dateKey].progressSum += progress;
    }
  }

  var dailyStats = Object.keys(dailyMap).sort().map(function(k) {
    var d = dailyMap[k];
    return { date: d.date, records: d.records, avgProgress: Math.round(d.progressSum / d.records) };
  });

  return {
    totalStudents: Object.keys(nameSet).length,
    totalRecords: totalRecords,
    totalTasks: totalTasks,
    avgProgress: totalRecords > 0 ? Math.round(progressSum / totalRecords) : 0,
    dailyStats: dailyStats.slice(-10)
  };
}

/*=====================================================================
 * 보조 함수
 *====================================================================*/
function ensureSubmitSheet_(ss) {
  var sh = ss.getSheetByName(SUBMIT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SUBMIT_SHEET);
    sh.getRange(1, 1, 1, SUBMIT_HEADERS.length).setValues([SUBMIT_HEADERS]);
    sh.getRange(1, 1, 1, SUBMIT_HEADERS.length).setFontWeight('bold').setBackground('#C8E6C9');
    sh.setFrozenRows(1);
    sh.setColumnWidth(4, 320); // 실천 내용 열
    sh.setColumnWidth(5, 200); // 한마디 열
    sh.setColumnWidth(6, 200); // 다짐 열
  }
  return sh;
}

function getSubmitFolder_() {
  var props = PropertiesService.getDocumentProperties();
  var id = props.getProperty(FOLDER_PROP_KEY);
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) {}
  }
  var ssName = SpreadsheetApp.getActiveSpreadsheet().getName();
  var folder = DriveApp.createFolder('📁 ' + ssName + ' - 제출물');
  props.setProperty(FOLDER_PROP_KEY, folder.getId());
  return folder;
}

function sanitize_(s) {
  return String(s).replace(/[\\/:*?"<>|]/g, '_').slice(0, 80) || 'file';
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// "사용 설명" 탭을 만든다.
function ensureGuideSheet_(ss) {
  var sh = ss.getSheetByName('사용 설명');
  if (!sh) sh = ss.insertSheet('사용 설명', 0);
  else sh.clear();
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(1); // 맨 앞으로

  sh.setHiddenGridlines(true);
  sh.setColumnWidth(1, 46);   // 번호/아이콘 배지
  sh.setColumnWidth(2, 740);  // 본문

  var r = 1;

  // ── 제목 ──
  band_(sh, r, '사용 설명 — SDG 리틀 히어로',
    { bg: '#1A237E', fg: '#FFFFFF', size: 16, bold: true, height: 50 });
  r++;
  // ── 소개 ──
  band_(sh, r, '학생 명단과 기록은 모두 선생님 구글 계정에만 저장돼요. 아래 4단계를 한 번만 하면 됩니다. (약 3분)',
    { bg: '#E8EAF6', fg: '#303F9F', size: 10, height: 48 });
  r++;

  // ── 선생님 설정 비밀번호 (학생에게 비공개) ──
  band_(sh, r, '앱 [선생님 설정] 비밀번호 :   ' + TEACHER_PASSCODE,
    { bg: '#FFE082', fg: '#5D4037', size: 13, bold: true, height: 38 });
  r++;
  band_(sh, r, '학생이 설정에 들어가지 못하게 막는 비밀번호예요. 학생에게는 알려주지 마세요. (앱에서 [선생님 설정]을 열 때 입력)',
    { bg: '#FFF8E1', fg: '#8A6D3B', size: 9, height: 26 });
  r++;
  r = gap_(sh, r);

  // ── 단계 카드 (번호 배지 + 색 본문) ──
  var steps = [
    { num: '1', head: '#10B981', bodyBg: '#E8F5E9', bodyFg: '#1B5E20',
      title: '학생 명단 입력',
      lines: ['아래 “학생 명단” 탭에서 번호(A열)와 이름(B열)을 적습니다. 여기 적은 학생이 그대로 앱 드롭다운에 연동돼요.'] },
    { num: '2', head: '#2563EB', bodyBg: '#E3F2FD', bodyFg: '#0D47A1',
      title: 'Apps Script 열기',
      lines: ['상단 메뉴에서  [확장 프로그램] ▸ [Apps Script]  를 누릅니다.'] },
    { num: '3', head: '#F59E0B', bodyBg: '#FFF8E1', bodyFg: '#7A4F01',
      title: '웹 앱으로 배포  (한 번만)',
      lines: [
        '①  오른쪽 위  [배포] ▸ [새 배포]',
        '②  톱니바퀴(⚙)를 눌러 유형  “웹 앱”  선택',
        '③  실행: “나”      /      액세스 권한: “모든 사용자”',
        '④  [배포]  클릭',
        '⑤  (처음이면) 권한 화면 →  [고급] ▸ [(안전하지 않음)으로 이동] ▸ [허용]',
        '⑥  배포 후 보이는 “웹 앱 URL”(끝이 /exec)을  복사'
      ],
      note: '※ 내가 만든 내 시트라서 안전합니다. 이 권한 화면은 처음 한 번만 나옵니다.' },
    { num: '4', head: '#7C3AED', bodyBg: '#F3E8FF', bodyFg: '#4A148C',
      title: '앱에 주소 붙여넣기',
      lines: [
        '①  앱 주소를 엽니다.',
        '②  첫 화면 아래쪽  [⚙ 선생님 설정]  버튼을 누릅니다. (비밀번호 입력)',
        '③  복사한 웹 앱 URL을 붙여넣습니다.',
        '④  [연결 테스트]를 눌러  “학생 N명”  이 나오는지 확인하고  [저장].'
      ],
      note: '※ 저장하면 브라우저 주소창이  「앱주소?api=…」  형태로 바뀝니다. 이것이 학생용 링크예요.' },
    { num: '5', head: '#0EA5E9', bodyBg: '#E0F2FE', bodyFg: '#075985',
      title: '학생에게 링크 나눠주기  (중요)',
      lines: [
        '①  [선생님 설정] 화면의  [학생용 링크 복사]  버튼을 누릅니다.',
        '②  복사된 링크를 클래스룸·메신저로 학생에게 보냅니다.',
        '③  학생은 그 링크로 들어오면 설정 없이 바로 이름 선택 화면이 떠요.'
      ],
      note: '※ 맨 앱 주소만 주면 학생 화면이 비어 있어요. 반드시  「?api=…」  가 붙은 학생용 링크를 주세요.' }
  ];
  for (var i = 0; i < steps.length; i++) {
    r = sectionCard_(sh, r, steps[i].num, steps[i].num + '단계.  ' + steps[i].title, steps[i]);
  }

  // ── 정보 섹션 ──
  var infos = [
    { num: '▶', head: '#0D9488', bodyBg: '#E0F2F1', bodyFg: '#004D40',
      title: '이제부터',
      lines: [
        '• 학생은 앱에서 실천 기록을 작성하고  [저장]  하면, “실천 기록” 탭에 자동으로 쌓입니다.',
        '• 링크는 우리 반 전용이에요. 학생에게만 공유하고, [선생님 설정] 비밀번호는 알려주지 마세요.'
      ] },
    { num: '📊', head: '#E65100', bodyBg: '#FFF3E0', bodyFg: '#E65100',
      title: '대시보드 보기',
      lines: [
        '• 우리 반 학생들이 기록한 통계 현황과 발자취는 앱 화면에서 한눈에 확인할 수 있습니다.',
        '• 데이터는 “실천 기록” 탭에서 실시간으로 불러옵니다.'
      ] },
    { num: '?', head: '#E11D48', bodyBg: '#FFE4E6', bodyFg: '#881337',
      title: '자주 묻는 것',
      lines: [
        '• 데이터는 누가 보나요?  →  선생님 본인만. 시트에만 저장됩니다.',
        '• 명단 열 순서를 바꿔도 되나요?  →  됩니다. “번호 / 이름” 글자가 든 제목을 자동으로 찾습니다.',
        '• 코드를 수정했어요.  →  [배포] ▸ [배포 관리] ▸ 편집(연필) ▸ 버전 “새 버전” ▸ [배포]  해야 반영됩니다.'
      ] }
  ];
  for (var k = 0; k < infos.length; k++) {
    r = sectionCard_(sh, r, infos[k].num, infos[k].title, infos[k]);
  }

  return sh;
}

// 색 헤더 띠 + 색 본문 줄들(+선택 메모)로 한 섹션을 그린다. 다음 행 번호 반환.
function sectionCard_(sh, r, badge, headerText, s) {
  sh.getRange(r, 1).setValue(badge)
    .setBackground(s.head).setFontColor('#FFFFFF').setFontSize(13).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(r, 2).setValue(headerText)
    .setBackground(s.head).setFontColor('#FFFFFF').setFontSize(12).setFontWeight('bold')
    .setVerticalAlignment('middle').setWrap(true);
  sh.setRowHeight(r, 32);
  r++;

  for (var j = 0; j < s.lines.length; j++) {
    sh.getRange(r, 1, 1, 2).setBackground(s.bodyBg);
    sh.getRange(r, 2).setValue(s.lines[j])
      .setFontColor(s.bodyFg).setFontSize(10).setVerticalAlignment('middle').setWrap(true);
    sh.setRowHeight(r, 32);
    r++;
  }

  if (s.note) {
    sh.getRange(r, 1, 1, 2).setBackground('#FFE0B2');
    sh.getRange(r, 2).setValue(s.note)
      .setFontColor('#8A4B00').setFontSize(9).setFontStyle('italic')
      .setVerticalAlignment('middle').setWrap(true);
    sh.setRowHeight(r, 26);
    r++;
  }

  return gap_(sh, r);
}

// A:B 한 행을 한 가지 색으로 채우고 본문을 쓴다.
function band_(sh, r, text, o) {
  sh.getRange(r, 1, 1, 2).setBackground(o.bg);
  var cell = sh.getRange(r, 2).setValue(text).setVerticalAlignment('middle').setWrap(true);
  cell.setFontColor(o.fg).setFontSize(o.size);
  if (o.bold) cell.setFontWeight('bold');
  sh.setRowHeight(r, o.height || 26);
}

// 얇은 간격 행
function gap_(sh, r) {
  sh.setRowHeight(r, 10);
  return r + 1;
}
