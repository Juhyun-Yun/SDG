/**
 * ============================================================
 * © 2026 GEG 화성(깊이 e끌림). All rights reserved.
 *
 * 본 코드는 「저작권법」상 보호받는 저작물입니다.
 * - 복제권(제16조)·공중송신권(제18조)·배포권(제20조)은
 *   저작권자에게 있습니다.
 * - 정식 경로로 받은 이용자라도 코드의 무단 복제·재배포·
 *   재판매·리브랜딩은 허용되지 않습니다.
 * - 무단 이용 시 「저작권법」 제136조(5년 이하 징역 또는
 *   5천만 원 이하 벌금) 및 제125조(손해배상) 적용 대상이
 *   될 수 있습니다.
 * - 이용 문의: bacusiki777@gmail.com, for2102@jimj.kr
 * ============================================================
 */

// 빌드 서명
const _BUILD_SIG = 'GEGHS-DEEPE-2026';

// 출처 확인용 함수
function getBuildInfo() {
  return {
    sig:   _BUILD_SIG,
    owner: 'GEG 화성(깊이 e끌림)',
    year:  2026
  };
}

const SHEET_NAME        = '실천 기록';
const SHEET_NAME_LEGACY = 'SDG_Records';   // 기존 배포 호환
const SHEET_STUDENTS    = '학생 명단';
const SHEET_STUDENTS_LEGACY = 'Students';  // 기존 배포 호환
const SHEET_DASHBOARD   = '📊 대시보드';

// ── 내부 헬퍼: 실천 기록 시트 가져오기 (없으면 생성) ──
function getRecordsSheet_(ss, create) {
  let sheet = ss.getSheetByName(SHEET_NAME)
           || ss.getSheetByName(SHEET_NAME_LEGACY)
           || ss.getSheetByName('records');
  if (!sheet && create) {
    sheet = ss.insertSheet(SHEET_NAME);
    const header = sheet.getRange(1, 1, 1, 6);
    sheet.appendRow(["기록 시간", "히어로 이름", "성장 지수 (%)", "실천한 내용", "느낀 점", "나의 다짐"]);
    header.setBackground("#10b981").setFontColor("white").setFontWeight("bold");
  }
  return sheet;
}

// ── 내부 헬퍼: 학생 명단 시트 가져오기 (없으면 생성) ──
function getStudentsSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_STUDENTS)
            || ss.getSheetByName(SHEET_STUDENTS_LEGACY)
            || ss.getSheetByName('학생명단')
            || ss.getSheetByName('명단')
            || ss.getSheetByName('student');
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_STUDENTS);
    sheet.appendRow(["번호", "이름"]);
    sheet.appendRow([1, "용기있는 사자"]);
    sheet.appendRow([2, "지혜로운 코끼리"]);
    sheet.getRange(1, 1, 1, 2).setBackground("#6366f1").setFontColor("white").setFontWeight("bold");
  }
  return sheet;
}

// ══════════════════════════════════════════════
//   doGet: 학생 명단 / 통계
// ══════════════════════════════════════════════
function doGet(e) {
  const params   = (e && e.parameter) || {};
  const action   = params.action;
  const heroName = params.heroName;

  if (action === 'getStudents') {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getStudentsSheet_(ss);
    const rows  = sheet.getDataRange().getValues().slice(1);
    const studentList = rows.map(function(r, i) {
      const cellA = String(r[0] || '').trim();
      const cellB = String(r[1] || '').trim();
      if (cellB) {
        return { num: cellA || (i + 1), name: cellB };
      } else if (cellA && isNaN(cellA)) {
        return { num: i + 1, name: cellA };
      }
      return null;
    }).filter(function(s){ return s !== null; });
    const names = studentList.map(function(s){ return s.name; });
    return ContentService.createTextOutput(JSON.stringify({ students: names, studentList: studentList }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'getClassStats') {
    const stats = computeClassStats_();
    return ContentService.createTextOutput(JSON.stringify({ stats: stats }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (heroName) {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getRecordsSheet_(ss, false);
    if (!sheet) return ContentService.createTextOutput(JSON.stringify({ history: [] }))
                  .setMimeType(ContentService.MimeType.JSON);
    const history = sheet.getDataRange().getValues().slice(1)
      .filter(function(r){ return r[1] === heroName; })
      .map(function(r){
        return { timestamp: r[0], heroName: r[1], progress: parseInt(r[2]),
                 tasks: r[3], reflection: r[4], myGoal: r[5] };
      }).reverse();
    return ContentService.createTextOutput(JSON.stringify({ history: history }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({ error: 'Unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════
//   doPost: 기록 저장 / 대시보드 생성 / 초기화
// ══════════════════════════════════════════════
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.getActiveSpreadsheet();

    if (data.action === 'generateDashboard') {
      return generateDashboardSheet();
    }

    if (data.action === 'clearData') {
      const sheet = getRecordsSheet_(ss, false);
      if (sheet) {
        sheet.clearContents();
        sheet.appendRow(["기록 시간", "히어로 이름", "성장 지수 (%)", "실천한 내용", "느낀 점", "나의 다짐"]);
        sheet.getRange(1, 1, 1, 6).setBackground("#10b981").setFontColor("white").setFontWeight("bold");
      }
      return ContentService.createTextOutput(JSON.stringify({ result: "success" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 기록 저장
    const sheet = getRecordsSheet_(ss, true);
    sheet.appendRow([
      data.timestamp || new Date().toLocaleString('ko-KR'),
      data.heroName,
      data.progress,
      data.tasks,
      data.reflection,
      data.myGoal || ''
    ]);
    try { maybeRegenerateDashboard(); } catch (err) { /* 무시 */ }

    return ContentService.createTextOutput(JSON.stringify({ result: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ result: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ══════════════════════════════════════════════
//   대시보드 자동 갱신 (쓰로틀: 30초)
// ══════════════════════════════════════════════
function maybeRegenerateDashboard() {
  const props   = PropertiesService.getScriptProperties();
  const lastStr = props.getProperty('dashboard_last_run');
  const now     = Date.now();
  if (lastStr && now - parseInt(lastStr, 10) < 30000) return;
  props.setProperty('dashboard_last_run', String(now));
  generateDashboardSheet();
}

// ══════════════════════════════════════════════
//   통계 집계 (내부 공통 함수)
// ══════════════════════════════════════════════
function computeClassStats_() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getRecordsSheet_(ss, false);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues().slice(1);
  if (data.length === 0) return null;

  const byHero = {}, byDate = {};
  let totalRecords = 0, totalProgressSum = 0, totalTasks = 0;

  data.forEach(function(row) {
    const ts         = String(row[0] || '');
    const name       = row[1];
    const progress   = parseInt(row[2]) || 0;
    const tasksStr   = String(row[3] || '');
    if (!name) return;
    const tasksCount = tasksStr.split('\n').filter(function(t){ return t.trim().length > 0; }).length;

    if (!byHero[name]) byHero[name] = { records: 0, progressSum: 0, latestProgress: 0, lastTimestamp: '' };
    byHero[name].records++;
    byHero[name].progressSum   += progress;
    byHero[name].latestProgress = progress;
    byHero[name].lastTimestamp  = ts;

    const dm = ts.match(/(\d+)\.\s*(\d+)\.\s*(\d+)/);
    if (dm) {
      const dateKey = dm[1] + '-' + ('0'+dm[2]).slice(-2) + '-' + ('0'+dm[3]).slice(-2);
      const label   = parseInt(dm[2],10) + '/' + parseInt(dm[3],10);
      if (!byDate[dateKey]) byDate[dateKey] = { label: label, totalTasks: 0, studentLatest: {}, participants: {} };
      byDate[dateKey].totalTasks += tasksCount;
      byDate[dateKey].studentLatest[name] = progress;
      byDate[dateKey].participants[name]   = true;
    }
    totalRecords++;
    totalProgressSum += progress;
    totalTasks       += tasksCount;
  });

  const heroStats = Object.keys(byHero).map(function(n) {
    return { name: n, records: byHero[n].records,
             avgProgress: Math.round(byHero[n].progressSum / byHero[n].records),
             latestProgress: byHero[n].latestProgress,
             lastTimestamp: byHero[n].lastTimestamp };
  }).sort(function(a,b){ return b.avgProgress - a.avgProgress; });

  const dailyStats = Object.keys(byDate).sort().map(function(k) {
    const scores = Object.values(byDate[k].studentLatest);
    const avg    = scores.length ? Math.round(scores.reduce(function(a,b){return a+b;},0)/scores.length) : 0;
    return { dateKey: k, label: byDate[k].label,
             totalTasks: byDate[k].totalTasks, avgProgress: avg,
             participants: Object.keys(byDate[k].participants).length };
  }).filter(function(d){ return d.avgProgress > 0; });

  return {
    totalStudents: heroStats.length,
    totalRecords:  totalRecords,
    totalTasks:    totalTasks,
    avgProgress:   totalRecords ? Math.round(totalProgressSum / totalRecords) : 0,
    heroStats:     heroStats,
    dailyStats:    dailyStats
  };
}

// ══════════════════════════════════════════════
//   대시보드 시트 생성/갱신
// ══════════════════════════════════════════════
function generateDashboardSheet() {
  try {
    const ss          = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = getRecordsSheet_(ss, false);
    if (!sourceSheet) return jsonOut_({ result:'error', message:'실천 기록 시트가 없습니다.' });

    const stats = computeClassStats_();
    if (!stats) return jsonOut_({ result:'error', message:'집계할 데이터가 없습니다.' });

    const { heroStats, dailyStats, totalStudents, totalRecords, totalTasks, avgProgress } = stats;

    let dash = ss.getSheetByName(SHEET_DASHBOARD);
    if (dash) {
      dash.getCharts().forEach(function(c){ dash.removeChart(c); });
      dash.clear();
    } else {
      dash = ss.insertSheet(SHEET_DASHBOARD, 0);
    }

    dash.setHiddenGridlines(true);
    try { dash.setTabColor('#10b981'); } catch(e){}
    dash.setFrozenRows(2);

    dash.getRange('A1').setValue('📊 우리 반 SDG 히어로 대시보드')
      .setFontSize(20).setFontWeight('bold').setFontColor('#10b981');
    dash.getRange('A2').setValue('마지막 업데이트: ' + new Date().toLocaleString('ko-KR'))
      .setFontSize(10).setFontColor('#64748b').setFontStyle('italic');

    dash.getRange('A4').setValue('▼ 요약 통계')
      .setFontSize(13).setFontWeight('bold').setFontColor('#0369a1');
    dash.getRange('A5:D5').setValues([['참여 친구 수','전체 기록 수','실천 횟수','평균 성장률']])
      .setBackground('#0369a1').setFontColor('white').setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    dash.getRange('A6:D6').setValues([[totalStudents+'명', totalRecords+'개', totalTasks+'회', avgProgress+'%']])
      .setBackground('#f0f9ff').setFontSize(16).setFontWeight('bold')
      .setHorizontalAlignment('center').setFontColor('#0369a1');
    dash.setRowHeight(5, 30);
    dash.setRowHeight(6, 45);

    let row = 8;
    dash.getRange(row,1).setValue('▼ 친구별 성장 통계 (평균 성장률 높은 순)')
      .setFontSize(13).setFontWeight('bold').setFontColor('#0369a1');
    row++;
    dash.getRange(row,1,1,5).setValues([['이름','기록 수','평균 성장률(%)','최근 성장률(%)','마지막 기록']])
      .setBackground('#0369a1').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center');
    row++;

    const heroStartRow = row;
    if (heroStats.length > 0) {
      const heroData = heroStats.map(function(h){ return [h.name,h.records,h.avgProgress,h.latestProgress,h.lastTimestamp]; });
      dash.getRange(row,1,heroData.length,5).setValues(heroData).setHorizontalAlignment('center');
      dash.getRange(row,1,heroData.length,1).setHorizontalAlignment('left').setFontWeight('bold');
      row += heroData.length;
    }

    row += 2;
    dash.getRange(row,1).setValue('▼ 날짜별 우리 반 평균 성장률')
      .setFontSize(13).setFontWeight('bold').setFontColor('#0369a1');
    row++;
    dash.getRange(row,1,1,4).setValues([['날짜','평균 성장률(%)','실천 횟수','참여 인원']])
      .setBackground('#0369a1').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center');
    row++;

    const dateStartRow = row;
    if (dailyStats.length > 0) {
      const dateData = dailyStats.map(function(d){ return [d.label, d.avgProgress||0, d.totalTasks, d.participants]; });
      dash.getRange(row,1,dateData.length,4).setValues(dateData).setHorizontalAlignment('center');
      row += dateData.length;
    }

    dash.setColumnWidth(1,160); dash.setColumnWidth(2,100);
    dash.setColumnWidth(3,130); dash.setColumnWidth(4,130); dash.setColumnWidth(5,200);

    if (dailyStats.length > 0) {
      const dateChart = dash.newChart().asLineChart()
        .addRange(dash.getRange(dateStartRow,1,dailyStats.length,1))
        .addRange(dash.getRange(dateStartRow,2,dailyStats.length,1))
        .setPosition(4,7,0,0)
        .setOption('title','날짜별 우리 반 평균 성장률 (%)')
        .setOption('legend',{position:'none'})
        .setOption('width',520).setOption('height',320)
        .setOption('colors',['#10b981']).setOption('curveType','function')
        .setOption('pointSize',6).setOption('lineWidth',3)
        .setOption('vAxis',{title:'평균 성장률 (%)',minValue:0,maxValue:100})
        .setOption('hAxis',{title:'날짜'})
        .build();
      dash.insertChart(dateChart);
    }

    if (heroStats.length > 0) {
      const n = Math.min(5, heroStats.length);
      const heroChart = dash.newChart().asBarChart()
        .addRange(dash.getRange(heroStartRow,1,n,1))
        .addRange(dash.getRange(heroStartRow,3,n,1))
        .setPosition(22,7,0,0)
        .setOption('title', n < heroStats.length ? '친구별 평균 성장률 TOP 5 (%)' : '친구별 평균 성장률 (%)')
        .setOption('legend',{position:'none'})
        .setOption('width',520).setOption('height',Math.max(280, n*55+120))
        .setOption('colors',['#10b981'])
        .setOption('hAxis',{title:'평균 성장률 (%)',minValue:0,maxValue:100})
        .setOption('vAxis',{textStyle:{fontSize:12,bold:true}})
        .build();
      dash.insertChart(heroChart);
    }

    ss.setActiveSheet(dash);
    return jsonOut_({ result:'success', message:'대시보드가 생성되었습니다.' });

  } catch(err) {
    return jsonOut_({ result:'error', message:err.toString() });
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════
//   Spreadsheet UI: 메뉴
// ══════════════════════════════════════════════
function renameLegacySheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // SDG_Records → 실천 기록
  const oldRecords = ss.getSheetByName(SHEET_NAME_LEGACY);
  if (oldRecords && !ss.getSheetByName(SHEET_NAME)) oldRecords.setName(SHEET_NAME);

  // Students → 학생 명단
  const oldStudents = ss.getSheetByName(SHEET_STUDENTS_LEGACY);
  if (oldStudents && !ss.getSheetByName(SHEET_STUDENTS)) oldStudents.setName(SHEET_STUDENTS);

  // 'records' 탭: 실천 기록이 없으면 이름 변경, 있으면 삭제
  const recordsTab = ss.getSheetByName('records');
  if (recordsTab) {
    if (!ss.getSheetByName(SHEET_NAME)) {
      recordsTab.setName(SHEET_NAME);
    } else {
      try { ss.deleteSheet(recordsTab); } catch(e) {}
    }
  }

  // 'student' 탭: 학생 명단이 없으면 이름 변경, 있으면 삭제
  const studentTab = ss.getSheetByName('student');
  if (studentTab) {
    if (!ss.getSheetByName(SHEET_STUDENTS)) {
      studentTab.setName(SHEET_STUDENTS);
    } else {
      try { ss.deleteSheet(studentTab); } catch(e) {}
    }
  }
}

function renameLegacySheetsFromMenu() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let changed = [];
    const oldRecords  = ss.getSheetByName(SHEET_NAME_LEGACY);
    if (oldRecords && !ss.getSheetByName(SHEET_NAME)) {
      oldRecords.setName(SHEET_NAME);
      changed.push(SHEET_NAME_LEGACY + ' → ' + SHEET_NAME);
    }
    const oldStudents = ss.getSheetByName(SHEET_STUDENTS_LEGACY);
    if (oldStudents && !ss.getSheetByName(SHEET_STUDENTS)) {
      oldStudents.setName(SHEET_STUDENTS);
      changed.push(SHEET_STUDENTS_LEGACY + ' → ' + SHEET_STUDENTS);
    }
    if (changed.length > 0) {
      ui.alert('✅ 탭 이름 변경 완료\n\n' + changed.join('\n'));
    } else {
      ui.alert('변경할 영문 탭이 없습니다.\n이미 한글 이름으로 되어 있거나, 해당 탭이 존재하지 않습니다.');
    }
  } catch(err) {
    ui.alert('❌ 오류\n\n' + err.toString());
  }
}

function onOpen() {
  try { renameLegacySheets_(); } catch(e) {}
  SpreadsheetApp.getUi()
    .createMenu('🌍 SDG 히어로')
    .addItem('📈 통계 사이드바 열기',       'showStatsSidebar')
    .addSeparator()
    .addItem('📊 대시보드 보기',             'showDashboardPopup')
    .addItem('🔄 대시보드 시트 새로고침',    'refreshDashboardFromMenu')
    .addSeparator()
    .addItem('🗑️ 기록 데이터 초기화',       'clearRecordsFromMenu')
    .addSeparator()
    .addItem('📋 사용 설명 만들기',          'setupGuideSheetFromMenu')
    .addToUi();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName('사용 설명')) {
    try { setupGuideSheet(); } catch(e) {}
  }
}

function showDashboardPopup() {
  const stats = computeClassStats_();
  const html  = HtmlService.createHtmlOutput(buildDashboardPopupHtml_(stats))
    .setWidth(520).setHeight(600).setTitle('📊 우리 반 SDG 대시보드');
  SpreadsheetApp.getUi().showModalDialog(html, '📊 우리 반 SDG 대시보드');
}

function buildDashboardPopupHtml_(stats) {
  const esc = function(n){ return String(n).replace(/[<>&"]/g, function(c){return({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;'}[c]);}); };
  if (!stats) {
    return '<div style="font-family:sans-serif;padding:30px;text-align:center;color:#64748b;">' +
           '<div style="font-size:2.5rem;margin-bottom:12px;">📭</div>' +
           '<p>아직 기록이 없어요.<br>학생이 한 번이라도 저장하면 통계가 나타납니다.</p></div>';
  }
  const top = stats.heroStats.slice(0, 10);
  const days = stats.dailyStats.slice(-7);

  const summaryHtml =
    '<div style="display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:8px;margin-bottom:16px;">' +
    [{ num: stats.totalStudents, label:'참여 친구', color:'#10b981' },
     { num: stats.totalRecords,  label:'전체 기록', color:'#0ea5e9' },
     { num: stats.totalTasks,    label:'실천 횟수', color:'#f59e0b' },
     { num: stats.avgProgress+'%', label:'평균 성장', color:'#8b5cf6' }
    ].map(function(c){
      return '<div style="background:#f8fafc;border-radius:10px;padding:10px 6px;text-align:center;border:1px solid #e2e8f0;">' +
             '<div style="font-size:1.5rem;font-weight:900;color:'+c.color+';">'+c.num+'</div>' +
             '<div style="font-size:0.7rem;color:#64748b;font-weight:700;margin-top:2px;">'+c.label+'</div></div>';
    }).join('') + '</div>';

  const heroRows = top.map(function(h, i){
    const medal = ['🥇','🥈','🥉'][i] || (i+1)+'.';
    const bar = '<div style="height:6px;background:#e2e8f0;border-radius:3px;margin-top:3px;">' +
                '<div style="height:6px;background:#10b981;border-radius:3px;width:'+h.avgProgress+'%;"></div></div>';
    return '<tr style="border-bottom:1px solid #f1f5f9;">' +
           '<td style="padding:7px 8px;font-size:0.85rem;">'+medal+'</td>' +
           '<td style="padding:7px 8px;font-weight:700;font-size:0.85rem;">'+esc(h.name)+'</td>' +
           '<td style="padding:7px 8px;text-align:center;">' +
           '<span style="font-weight:900;color:#10b981;font-size:0.9rem;">'+h.avgProgress+'%</span>'+bar+'</td>' +
           '<td style="padding:7px 8px;text-align:center;color:#64748b;font-size:0.8rem;">'+h.records+'회</td></tr>';
  }).join('');

  const dayRows = days.map(function(d){
    const bar = '<div style="height:6px;background:#e2e8f0;border-radius:3px;margin-top:3px;">' +
                '<div style="height:6px;background:#6366f1;border-radius:3px;width:'+d.avgProgress+'%;"></div></div>';
    return '<tr style="border-bottom:1px solid #f1f5f9;">' +
           '<td style="padding:7px 8px;font-weight:700;font-size:0.85rem;">'+d.label+'</td>' +
           '<td style="padding:7px 8px;text-align:center;">'+
           '<span style="font-weight:900;color:#6366f1;font-size:0.9rem;">'+d.avgProgress+'%</span>'+bar+'</td>' +
           '<td style="padding:7px 8px;text-align:center;color:#64748b;font-size:0.8rem;">'+d.participants+'명</td></tr>';
  }).join('');

  return '<div style="font-family:-apple-system,sans-serif;padding:16px;color:#0f172a;font-size:0.9rem;">' +
    '<p style="color:#64748b;font-size:0.75rem;margin:0 0 12px;">업데이트: ' + new Date().toLocaleString('ko-KR') + '</p>' +
    summaryHtml +
    '<h3 style="font-size:0.88rem;color:#0369a1;margin:0 0 6px;font-weight:900;">🏆 친구별 성장 순위</h3>' +
    '<table style="width:100%;border-collapse:collapse;margin-bottom:16px;">' +
    '<thead><tr style="background:#f0f9ff;">' +
    '<th style="padding:6px 8px;text-align:left;font-size:0.75rem;color:#64748b;width:32px;"></th>' +
    '<th style="padding:6px 8px;text-align:left;font-size:0.75rem;color:#64748b;">이름</th>' +
    '<th style="padding:6px 8px;text-align:center;font-size:0.75rem;color:#64748b;">평균 성장률</th>' +
    '<th style="padding:6px 8px;text-align:center;font-size:0.75rem;color:#64748b;">기록 수</th></tr></thead>' +
    '<tbody>'+heroRows+'</tbody></table>' +
    '<h3 style="font-size:0.88rem;color:#0369a1;margin:0 0 6px;font-weight:900;">📅 최근 7일 현황</h3>' +
    '<table style="width:100%;border-collapse:collapse;margin-bottom:16px;">' +
    '<thead><tr style="background:#f0f9ff;">' +
    '<th style="padding:6px 8px;text-align:left;font-size:0.75rem;color:#64748b;">날짜</th>' +
    '<th style="padding:6px 8px;text-align:center;font-size:0.75rem;color:#64748b;">평균 성장률</th>' +
    '<th style="padding:6px 8px;text-align:center;font-size:0.75rem;color:#64748b;">참여 인원</th></tr></thead>' +
    '<tbody>'+dayRows+'</tbody></table>' +
    '<button onclick="google.script.host.close()" style="width:100%;background:#6366f1;color:white;border:none;' +
    'border-radius:8px;padding:10px;font-weight:800;cursor:pointer;font-size:0.85rem;">닫기</button>' +
    '</div>';
}

function refreshDashboardFromMenu() {
  const ui = SpreadsheetApp.getUi();
  try {
    generateDashboardSheet();
    ui.alert('✅ 대시보드 시트가 갱신되었습니다.');
  } catch(err) {
    ui.alert('❌ 오류\n\n' + (err && err.message ? err.message : err));
  }
}

function clearRecordsFromMenu() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.alert(
    '⚠️ 기록 데이터 초기화',
    '실천 기록 시트의 모든 데이터가 영구적으로 삭제됩니다.\n학생 명단과 대시보드 시트는 유지됩니다.\n\n정말 초기화하시겠습니까?',
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) return;

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getRecordsSheet_(ss, false);
  if (sheet) {
    sheet.clearContents();
    sheet.appendRow(["기록 시간","히어로 이름","성장 지수 (%)","실천한 내용","느낀 점","나의 다짐"]);
    sheet.getRange(1,1,1,6).setBackground("#10b981").setFontColor("white").setFontWeight("bold");
    ui.alert('✅ 기록 데이터가 초기화되었습니다.');
  } else {
    ui.alert('실천 기록 시트가 없습니다. 초기화할 데이터가 없어요.');
  }
}

// ══════════════════════════════════════════════
//   통계 사이드바
// ══════════════════════════════════════════════
function showStatsSidebar() {
  const stats = computeClassStats_();
  const html  = HtmlService.createHtmlOutput(buildSidebarHtml_(stats))
    .setTitle('🌍 우리 반 SDG 통계');
  SpreadsheetApp.getUi().showSidebar(html);
}

function buildSidebarHtml_(stats) {
  if (!stats) {
    return '<div style="font-family:sans-serif;padding:20px;text-align:center;color:#64748b;">' +
           '아직 데이터가 없어요.<br><br>학생이 한 번이라도 저장하면<br>통계가 표시됩니다 ✨</div>';
  }
  const top5    = stats.heroStats.slice(0, 5);
  const recent5 = stats.dailyStats.slice(-5);
  const esc     = function(n){ return String(n).replace(/[<>&"]/g, function(c){return({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;'}[c]);}); };

  const cards = [
    { num: stats.totalStudents,       label:'참여 친구', color:'#10b981' },
    { num: stats.totalRecords,        label:'전체 기록', color:'#0ea5e9' },
    { num: stats.totalTasks,          label:'실천 횟수', color:'#f59e0b' },
    { num: stats.avgProgress + '%',   label:'평균 성장', color:'#8b5cf6' }
  ].map(function(c){
    return '<div style="background:white;border:1px solid #e2e8f0;border-radius:10px;padding:10px 6px;text-align:center;">' +
           '<div style="font-size:1.4rem;font-weight:900;color:'+c.color+';line-height:1.1;">'+c.num+'</div>' +
           '<div style="font-size:0.7rem;color:#64748b;font-weight:700;margin-top:4px;">'+c.label+'</div></div>';
  }).join('');

  const top5Html = top5.length ? top5.map(function(h,i){
    const medal = ['🥇','🥈','🥉','4️⃣','5️⃣'][i];
    return '<div style="display:flex;align-items:center;gap:8px;padding:6px 8px;background:#f8fafc;border-radius:8px;margin-bottom:4px;">' +
           '<span>'+medal+'</span>' +
           '<span style="flex:1;font-size:0.78rem;font-weight:700;color:#0f172a;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">'+esc(h.name)+'</span>' +
           '<span style="font-size:0.78rem;font-weight:900;color:#10b981;">'+h.avgProgress+'%</span></div>';
  }).join('') : '<div style="color:#94a3b8;font-size:0.78rem;text-align:center;padding:8px;">기록 없음</div>';

  const recentHtml = recent5.length ? recent5.map(function(d){
    return '<div style="display:flex;justify-content:space-between;padding:5px 8px;background:#f8fafc;border-radius:8px;margin-bottom:3px;">' +
           '<span style="font-size:0.78rem;font-weight:700;color:#475569;">'+d.label+'</span>' +
           '<span style="font-size:0.78rem;color:#475569;">실천 <b style="color:#0369a1;">'+d.totalTasks+'</b> · 참여 '+d.participants+'명</span></div>';
  }).join('') : '<div style="color:#94a3b8;font-size:0.78rem;text-align:center;padding:8px;">기록 없음</div>';

  return '<div style="font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',\'Noto Sans KR\',sans-serif;padding:14px 12px;background:#f1f5f9;min-height:100vh;color:#0f172a;">' +
    '<h2 style="font-size:1rem;color:#10b981;margin:0 0 10px;font-weight:900;">🌍 우리 반 현황</h2>' +
    '<div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:14px;">' + cards + '</div>' +
    '<h3 style="font-size:0.85rem;color:#0369a1;margin:0 0 6px;font-weight:900;">🏆 친구별 TOP 5</h3>' +
    '<div style="margin-bottom:14px;">' + top5Html + '</div>' +
    '<h3 style="font-size:0.85rem;color:#0369a1;margin:0 0 6px;font-weight:900;">📅 최근 5일 실천</h3>' +
    '<div style="margin-bottom:14px;">' + recentHtml + '</div>' +
    '<button onclick="google.script.run.withSuccessHandler(function(){ google.script.run.showStatsSidebar(); }).generateDashboardSheet()" ' +
      'style="width:100%;background:#10b981;color:white;border:none;border-radius:8px;padding:9px;font-weight:800;cursor:pointer;font-size:0.82rem;margin-bottom:6px;">🔄 통계 새로고침 + 대시보드 갱신</button>' +
    '<button onclick="google.script.run.openDashboardSheet()" ' +
      'style="width:100%;background:white;color:#475569;border:1px solid #cbd5e1;border-radius:8px;padding:9px;font-weight:800;cursor:pointer;font-size:0.82rem;">📊 대시보드 시트로 이동</button>' +
    '<p style="font-size:0.65rem;color:#94a3b8;text-align:center;margin-top:10px;">학생 저장 시 자동으로 갱신됩니다</p>' +
    '</div>';
}

// ══════════════════════════════════════════════
//   사용 설명 시트 생성
// ══════════════════════════════════════════════
function setupGuideSheet() {
  var ss         = SpreadsheetApp.getActiveSpreadsheet();
  var GUIDE_NAME = '사용 설명';
  var OLD_NAMES  = ['📋 사용법', '사용 설명'];

  // 1. 임시 탭 생성 (마지막 탭 삭제 오류 방지)
  var tempName = '_sdg_tmp_' + Date.now();
  var guide    = ss.insertSheet(tempName);

  // 2. 기존 사용 안내 탭 삭제
  OLD_NAMES.forEach(function(n) {
    var s = ss.getSheetByName(n);
    if (s) { try { ss.deleteSheet(s); } catch(e) {} }
  });

  // 3. 탭 이름 변경 후 첫 번째 위치로 이동
  guide.setName(GUIDE_NAME);
  ss.setActiveSheet(guide);
  ss.moveActiveSheet(1);

  // 4. 시트 기본 설정
  guide.setHiddenGridlines(true);
  try { guide.setTabColor('#10b981'); } catch(e) {}
  guide.setColumnWidth(1, 170);
  guide.setColumnWidth(2, 510);

  // 5. 콘텐츠 배열 구성 및 행 위치 추적
  var rows = [];
  var sn   = 1;
  var fmt  = {
    titleRow:     -1,
    sectionRows:  [],
    tableHdrRows: [],
    mergeRows:    [],
    noticeRow:    -1,
    copyrightRow: -1,
    borderRanges: [],         // [r1, c1, r2, c2] 1-indexed
    sidebarData:  { s: -1, e: -1 }
  };

  function push(a, b) {
    rows.push([a === undefined ? '' : a, b === undefined ? '' : b]);
    return rows.length;
  }
  function blank() { return push('', ''); }

  // ── 제목 ──
  fmt.titleRow = push("SDG's 리틀 히어로 — 시트 사용 설명", '');
  blank();

  // ── 섹션 1: 사본을 만든 뒤 설정하기 ──
  fmt.sectionRows.push(push(sn++ + '. 사본을 만든 뒤 설정하기', ''));
  fmt.tableHdrRows.push(push('단계', '내용'));
  var s1s = rows.length + 1;
  push('①',
    'Apps Script 웹앱 배포\n' +
    '구글 시트 상단 메뉴에서 [확장 프로그램] → [Apps Script] 클릭\n' +
    '[배포] → [새 배포] 클릭 → 유형: 웹앱 선택\n' +
    '액세스 권한: 반드시 "모든 사용자(Anyone)"로 설정\n' +
    '[배포] 버튼 클릭 후 표시된 웹 앱 URL 복사');
  push('②',
    '앱에 배포 URL 등록\n' +
    'GitHub Pages에 배포된 SDG 리틀 히어로 앱 접속\n' +
    '하단 [선생님 메뉴] → [설정] 탭\n' +
    '[본부(Apps Script) URL] 칸에 복사한 URL 붙여넣고 저장\n' +
    '저장 후 학생 접속 주소가 자동으로 생성됩니다');
  push('③',
    '학생 명단 입력\n' +
    '이 스프레드시트의 [학생 명단] 탭 클릭\n' +
    '1행은 헤더(번호 / 이름)이므로 수정하지 마세요\n' +
    '2행부터 A열에 번호, B열에 학생 이름을 한 명씩 입력\n' +
    '앱 새로고침 시 학생 선택 목록에 자동 반영');
  push('④',
    '학생용 접속 주소 공유\n' +
    '앱 [선생님 메뉴] → [설정] 탭 → 학생용 접속 주소 복사\n' +
    '카카오톡·구글 클래스룸·QR코드 등으로 학생에게 공유\n' +
    '이 주소로 접속하면 우리 학급으로 자동 연결됩니다\n\n' +
    '※ 왜 따로 공유해야 하나요?\n' +
    '여러 선생님이 같은 앱을 사용할 때, 학생용 주소에는\n' +
    '우리 학급 시트 정보가 포함되어 있어서 학생이\n' +
    '그 링크로 접속하면 자동으로 우리 반에 연결됩니다.');
  var s1e = rows.length;
  fmt.borderRanges.push([s1s - 1, 1, s1e, 2]);
  blank();

  // ── 섹션 2: 시트 탭 안내 ──
  fmt.sectionRows.push(push(sn++ + '. 시트 탭 안내', ''));
  fmt.tableHdrRows.push(push('탭 이름', '역할 및 안내'));
  var s2s = rows.length + 1;
  push('사용 설명',         '현재 보고 있는 시트입니다. 처음 사용할 때 참고하세요.');
  push('학생 명단',         '학생 이름을 입력하는 시트입니다.\nA열: 번호, B열: 이름 형식으로 입력하세요.\n2행부터 한 명씩 입력하면 앱에 자동으로 반영됩니다.');
  push('실천 기록',         '학생이 앱에서 저장한 실천 내용이 자동으로 기록됩니다.\n직접 수정하지 마세요.');
  push('📊 대시보드',      '학급 전체 성장 통계와 차트가 자동으로 생성됩니다.\n메뉴에서 새로고침하거나 학생 저장 시 자동 갱신됩니다.');
  var s2e = rows.length;
  fmt.borderRanges.push([s2s - 1, 1, s2e, 2]);
  blank();

  // ── 주의 안내 (한 번만) ──
  fmt.noticeRow = push(
    '⚠ 주의',
    '데이터나 설정을 변경할 때는 앱 화면이 아니라 해당 시트 탭에서 직접 수정하세요.\n탭 이름은 코드와 연결되어 있으므로 삭제하거나 변경하지 마세요.'
  );
  blank();

  // ── 섹션 3: 메뉴 사용 방법 ──
  fmt.sectionRows.push(push(sn++ + '. 메뉴 사용 방법', ''));
  fmt.mergeRows.push(push('구글 시트 상단 메뉴 [🌍 SDG 히어로]를 클릭하면 다음 기능을 사용할 수 있습니다.', ''));
  fmt.tableHdrRows.push(push('메뉴 항목', '기능'));
  var s3s = rows.length + 1;
  push('📈 통계 사이드바 열기',     '참여 학생 수, 기록 수, 실천 횟수, 평균 성장률 및 친구별 TOP 5 순위를\n화면 오른쪽 사이드바로 표시합니다.');
  push('📊 대시보드 보기',           '팝업 창에서 학급 통계 요약과 최근 7일 현황을 확인합니다.');
  push('🔄 대시보드 시트 새로고침', '📊 대시보드 시트를 최신 데이터로 다시 생성합니다.');
  push('🗑️ 기록 데이터 초기화',    '실천 기록 시트의 모든 내용을 삭제합니다. (복구 불가)');
  push('📋 사용 설명 만들기',       '이 사용 설명 시트를 새로 만듭니다.');
  var s3e = rows.length;
  fmt.borderRanges.push([s3s - 1, 1, s3e, 2]);
  blank();

  // ── 섹션 4: 통계 사이드바 ──
  fmt.sectionRows.push(push(sn++ + '. 통계 사이드바', ''));
  fmt.sidebarData.s = rows.length + 1;
  push('사용 방법',
    '구글 시트 상단 메뉴 [🌍 SDG 히어로] → [📈 통계 사이드바 열기] 클릭\n' +
    '시트 오른쪽에 사이드바가 열립니다.');
  push('확인 가능한 정보',
    '• 참여 학생 수, 전체 기록 수, 실천 횟수, 반 평균 성장률\n' +
    '• 친구별 성장 순위 TOP 5\n' +
    '• 최근 5일간의 실천 현황\n' +
    '사이드바 내 [통계 새로고침 + 대시보드 갱신] 버튼으로 즉시 갱신됩니다.');
  fmt.sidebarData.e = rows.length;
  fmt.borderRanges.push([fmt.sidebarData.s, 1, fmt.sidebarData.e, 2]);
  blank();

  // ── 섹션 5: 교사 대시보드 ──
  fmt.sectionRows.push(push(sn++ + '. 교사 대시보드', ''));
  fmt.tableHdrRows.push(push('방법', '안내'));
  var s5s = rows.length + 1;
  push('팝업 대시보드',
    '[🌍 SDG 히어로] → [📊 대시보드 보기] 선택\n' +
    '참여 친구 수, 평균 성장률, 친구별 순위, 최근 7일 현황을\n팝업 창으로 확인합니다.');
  push('대시보드 시트',
    '[🌍 SDG 히어로] → [🔄 대시보드 시트 새로고침] 선택\n' +
    '📊 대시보드 시트에 차트와 통계표가 생성됩니다.\n' +
    '학생 기록이 저장될 때 자동으로 갱신됩니다 (30초 간격).');
  var s5e = rows.length;
  fmt.borderRanges.push([s5s - 1, 1, s5e, 2]);
  blank();

  // ── 섹션 6 (마지막): 저작권 안내 ──
  fmt.sectionRows.push(push(sn++ + '. 저작권 안내', ''));
  fmt.copyrightRow = push(
    '저작권',
    '본 구글 시트 및 관련 자료(앱, 코드, 콘텐츠 포함)의 저작권은 GEG 화성(깊이 e끌림)에게 있습니다.\n\n' +
    '1. 본 자료는 책을 구입한 자에 한해 이용이 허락됩니다(교사일 경우는 해당 학급, 학부모일 경우 자녀). 정상 경로로 구매하거나 배포받은 이용자라 하더라도 앱 코드의 무단 수정 및 2차 배포는 허용되지 않습니다.\n\n' +
    '2. 다음 행위를 금합니다.\n' +
    '· 무단 복제·전송·배포·공유(타인에게 시트 링크 또는 사본 전달 포함)\n' +
    '· 영리 목적의 사용 또는 배포(학원에서의 사용 포함)\n' +
    '· 영리 목적의 재판매 또는 재배포\n' +
    '· 무단 수정·편집을 통한 2차적 저작물 작성\n\n' +
    '3. 「저작권법」 제136조(벌칙) 제1항 제1호에 따라, 저작재산권을 복제·공연·공중송신·전시·배포·대여·2차적저작물 작성의 방법으로 침해한 자는 5년 이하의 징역 또는 5천만원 이하의 벌금에 처하거나 이를 병과할 수 있습니다.\n\n' +
    'ⓒ 2026 GEG 화성(깊이 e끌림)'
  );

  // 6. setValues() 한 번에 입력
  var totalRows    = rows.length;
  var contentRange = guide.getRange(1, 1, totalRows, 2);
  contentRange.setValues(rows);

  // 7. 기본 서식 (전체 영역)
  contentRange
    .setFontSize(10)
    .setFontColor('#1e293b')
    .setBackground('#ffffff')
    .setWrap(true)
    .setVerticalAlignment('top')
    .setHorizontalAlignment('left');

  // 8. 제목 서식
  guide.getRange(fmt.titleRow, 1, 1, 2).merge()
    .setFontSize(14).setFontWeight('bold').setFontColor('#10b981')
    .setBackground('#f0fdf4').setHorizontalAlignment('center').setVerticalAlignment('middle');

  // 9. 섹션 제목 서식
  fmt.sectionRows.forEach(function(r) {
    guide.getRange(r, 1, 1, 2).merge()
      .setFontSize(11).setFontWeight('bold').setFontColor('#ffffff')
      .setBackground('#10b981').setVerticalAlignment('middle');
  });

  // 10. 표 헤더 서식
  fmt.tableHdrRows.forEach(function(r) {
    guide.getRange(r, 1, 1, 2)
      .setFontWeight('bold').setFontColor('#0f172a')
      .setBackground('#d1fae5').setHorizontalAlignment('center');
  });

  // 11. 병합 행 서식 (메뉴 섹션 안내 줄)
  fmt.mergeRows.forEach(function(r) {
    guide.getRange(r, 1, 1, 2).merge()
      .setFontColor('#475569').setFontStyle('italic').setBackground('#f8fafc');
  });

  // 12. 주의 안내 서식
  if (fmt.noticeRow > 0) {
    guide.getRange(fmt.noticeRow, 1)
      .setFontWeight('bold').setFontColor('#92400e').setBackground('#fef9c3');
    guide.getRange(fmt.noticeRow, 2)
      .setBackground('#fef9c3').setFontColor('#78350f');
  }

  // 13. 저작권 행 서식
  if (fmt.copyrightRow > 0) {
    guide.getRange(fmt.copyrightRow, 1)
      .setFontWeight('bold').setFontColor('#374151').setBackground('#f1f5f9');
    guide.getRange(fmt.copyrightRow, 2)
      .setFontColor('#374151').setBackground('#f8fafc');
  }

  // 14. 사이드바 데이터 행 서식
  if (fmt.sidebarData.s > 0) {
    var sdRows = fmt.sidebarData.e - fmt.sidebarData.s + 1;
    guide.getRange(fmt.sidebarData.s, 1, sdRows, 1)
      .setFontWeight('bold').setBackground('#f0fdf4').setFontColor('#065f46');
    guide.getRange(fmt.sidebarData.s, 2, sdRows, 1)
      .setBackground('#ffffff');
  }

  // 15. 표 테두리
  fmt.borderRanges.forEach(function(br) {
    guide.getRange(br[0], br[1], br[2] - br[0] + 1, br[3] - br[1] + 1)
      .setBorder(true, true, true, true, true, true,
                 '#94a3b8', SpreadsheetApp.BorderStyle.SOLID);
  });
}

function setupGuideSheetFromMenu() {
  var ui = SpreadsheetApp.getUi();
  try {
    setupGuideSheet();
    ui.alert('✅ 사용 설명 시트가 생성되었습니다.');
  } catch(err) {
    ui.alert('❌ 오류\n\n' + err.toString());
  }
}
