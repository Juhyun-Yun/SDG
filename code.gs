/*
  SDG's 리틀 히어로 - Google Apps Script Backend (code.gs)
  1. 구글 시트에서 확장 프로그램 → Apps Script
  2. 이 코드 전체를 붙여넣기
  3. 배포 → 새 배포 → 웹 앱 → 액세스: 모든 사용자(Anyone)
  4. 배포 후 웹 앱 URL을 앱 선생님 메뉴 → 설정 → 본부 URL에 붙여넣기
*/

const SHEET_NAME        = '실천 기록';
const SHEET_NAME_LEGACY = 'SDG_Records';   // 기존 배포 호환
const SHEET_STUDENTS    = '학생 명단';
const SHEET_STUDENTS_LEGACY = 'Students';  // 기존 배포 호환
const SHEET_DASHBOARD   = '📊 대시보드';
const SHEET_GUIDE       = '📋 선생님 가이드';

// ── 내부 헬퍼: 실천 기록 시트 가져오기 (없으면 생성) ──
function getRecordsSheet_(ss, create) {
  let sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheetByName(SHEET_NAME_LEGACY);
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
            || ss.getSheetByName('명단');
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
//   doGet: 학생 명단 / 통계 / HTML 서빙
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
        // A=번호, B=이름 구조
        return { num: cellA || (i + 1), name: cellB };
      } else if (cellA && isNaN(cellA)) {
        // A=이름만 있는 구조 (구버전 호환)
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
  const oldRecords  = ss.getSheetByName(SHEET_NAME_LEGACY);
  if (oldRecords && !ss.getSheetByName(SHEET_NAME)) oldRecords.setName(SHEET_NAME);
  const oldStudents = ss.getSheetByName(SHEET_STUDENTS_LEGACY);
  if (oldStudents && !ss.getSheetByName(SHEET_STUDENTS)) oldStudents.setName(SHEET_STUDENTS);
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
    .addToUi();

  // 선생님 가이드 시트가 없으면 자동 생성
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SHEET_GUIDE)) {
    try { createGuideSheet_(); } catch(e) {}
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
//   선생님 가이드 시트 생성
// ══════════════════════════════════════════════
function createGuideSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let guide = ss.getSheetByName(SHEET_GUIDE);
  if (guide) {
    guide.clear();
    guide.getCharts().forEach(function(c){ guide.removeChart(c); });
  } else {
    guide = ss.insertSheet(SHEET_GUIDE);
  }

  guide.setHiddenGridlines(true);
  try { guide.setTabColor('#6366f1'); } catch(e){}
  guide.setColumnWidth(1, 22);
  guide.setColumnWidth(2, 320);
  guide.setColumnWidth(3, 440);
  guide.setColumnWidth(4, 22);

  function title(row, text) {
    guide.getRange(row, 1, 1, 4).merge()
      .setValue(text)
      .setFontSize(15).setFontWeight('bold').setFontColor('#ffffff')
      .setBackground('#6366f1')
      .setVerticalAlignment('middle').setHorizontalAlignment('left')
      .setPaddings ? null : null;
    guide.setRowHeight(row, 38);
  }

  function subtitle(row, text) {
    guide.getRange(row, 2, 1, 3).merge()
      .setValue(text)
      .setFontSize(11).setFontWeight('bold').setFontColor('#4338ca')
      .setBackground('#eef2ff');
    guide.setRowHeight(row, 28);
  }

  function row2col(row, label, value) {
    guide.getRange(row, 2).setValue(label)
      .setFontSize(10).setFontWeight('bold').setFontColor('#374151')
      .setBackground('#f9fafb').setVerticalAlignment('top').setWrap(true);
    guide.getRange(row, 3).setValue(value)
      .setFontSize(10).setFontColor('#1e293b')
      .setBackground('#ffffff').setVerticalAlignment('top').setWrap(true);
    guide.setRowHeight(row, 52);
  }

  function blank(row) {
    guide.getRange(row, 1, 1, 4).setBackground('#f8fafc');
    guide.setRowHeight(row, 10);
  }

  function note(row, text) {
    guide.getRange(row, 2, 1, 3).merge()
      .setValue(text)
      .setFontSize(9.5).setFontColor('#64748b').setFontStyle('italic')
      .setBackground('#f8fafc').setWrap(true);
    guide.setRowHeight(row, 36);
  }

  let r = 1;

  // 헤더
  guide.getRange(r, 1, 1, 4).merge()
    .setValue('📋 SDG\'s 리틀 히어로 — 선생님 가이드')
    .setFontSize(20).setFontWeight('bold').setFontColor('#4f46e5')
    .setBackground('#eef2ff').setHorizontalAlignment('center').setVerticalAlignment('middle');
  guide.setRowHeight(r, 52); r++;

  guide.getRange(r, 1, 1, 4).merge()
    .setValue('이 시트는 앱과 구글 시트 연결, 배포, 데이터 관리 방법을 안내합니다.')
    .setFontSize(10).setFontColor('#64748b').setBackground('#eef2ff')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  guide.setRowHeight(r, 26); r++;
  blank(r); r++;

  // ─ 섹션 1 ─
  title(r, '  🚀  1단계 · 처음 시작하기 (앱-스프레드시트 연결)'); r++;
  blank(r); r++;

  subtitle(r, '  ① 스프레드시트 사본 만들기'); r++;
  row2col(r,
    '방법',
    '제공받은 SDG 리틀 히어로 템플릿 시트를 열고\n파일 → 사본 만들기 클릭\n사본 이름은 자유롭게 지정 (예: "3학년 2반 SDG 히어로")\nApps Script 코드도 자동으로 함께 복사됩니다.');
  r++;

  subtitle(r, '  ② Apps Script 배포하기'); r++;
  row2col(r, '경로', '내 사본 시트에서\n확장 프로그램 → Apps Script'); r++;
  row2col(r,
    '배포 설정',
    '우상단 [배포] → [새 배포] 클릭\n유형: 웹 앱 선택\n⚠️ 액세스 권한: 반드시 "모든 사용자(Anyone)" 설정\n   ("나만" 또는 "본인만"으로 두면 학생 데이터를 받지 못합니다!)\n[배포] 버튼 클릭 후 표시된 웹 앱 URL 복사');
  r++;

  subtitle(r, '  ③ 앱에 URL 등록하기'); r++;
  row2col(r,
    '경로',
    '앱 하단 [선생님 메뉴] → 설정 탭\n→ 본부(Apps Script) URL 칸에 복사한 URL 붙여넣기 → 저장\n→ 스프레드시트 주소 칸에 이 시트의 브라우저 URL 붙여넣기 → 저장');
  r++;
  note(r, '※ 저장 후 페이지를 새로고침(F5)하면 우리 학급 학생 명단이 표시됩니다.'); r++;
  blank(r); r++;

  // ─ 섹션 2 ─
  title(r, '  👩‍🏫  2단계 · 학생 명단 관리'); r++;
  blank(r); r++;
  row2col(r,
    '"학생 명단" 시트 편집',
    '이 스프레드시트의 "학생 명단" 시트를 클릭하세요.\n1행: 헤더 (A1=번호, B1=이름) — 수정하지 마세요.\n2행부터 A열에 번호, B열에 학생 이름을 한 줄씩 입력하면\n앱 로그인 화면에 자동으로 반영됩니다.');
  r++;
  note(r, '※ 이름 추가·삭제 후 앱을 새로고침하면 바로 적용됩니다.'); r++;
  blank(r); r++;

  // ─ 섹션 3 ─
  title(r, '  🔗  3단계 · 학생용 접속 주소 공유'); r++;
  blank(r); r++;
  row2col(r,
    '학생용 주소 확인',
    '앱 [선생님 메뉴] → 설정 탭\n→ "학생용 접속 주소" 항목에서 주소 복사\n→ 카카오톡, 구글 클래스룸, QR코드 등으로 학생에게 공유');
  r++;
  row2col(r,
    '왜 따로 공유해야 하나요?',
    '여러 선생님이 같은 앱 주소를 공유할 때,\n학생용 주소에는 우리 학급 시트가 인코딩되어 있어서\n학생이 그 링크로 접속하면 자동으로 우리 반에 연결됩니다.');
  r++;
  blank(r); r++;

  // ─ 섹션 4 ─
  title(r, '  📊  4단계 · 데이터 확인 및 관리'); r++;
  blank(r); r++;

  subtitle(r, '  통계 확인'); r++;
  row2col(r,
    '사이드바 통계',
    '상단 메뉴 [🌍 SDG 히어로] → [📈 통계 사이드바 열기]\n참여 친구 수, 전체 기록, 실천 횟수, 평균 성장률\nTOP 5 친구, 최근 5일 실천 현황 확인');
  r++;
  row2col(r,
    '대시보드 시트',
    '상단 메뉴 [🌍 SDG 히어로] → [🔄 대시보드 새로고침]\n→ "📊 대시보드" 시트에 자동으로 차트와 통계표 생성\n학생 기록 저장 시 자동 갱신 (30초 쓰로틀)');
  r++;

  subtitle(r, '  데이터 초기화'); r++;
  row2col(r,
    '기록 데이터만 삭제',
    '상단 메뉴 [🌍 SDG 히어로] → [🗑️ 기록 데이터 초기화]\n실천 기록 시트의 모든 데이터를 삭제합니다.\n학생 명단과 대시보드 시트는 유지됩니다.\n\n또는 앱 [선생님 메뉴] → 설정 탭 → "시트 데이터 비우기"');
  r++;
  row2col(r,
    '앱 기기 데이터만 삭제',
    '앱 [선생님 메뉴] → 설정 탭 → "기기 데이터 초기화"\n해당 기기의 localStorage 데이터만 삭제됩니다.\n서버(구글 시트) 기록은 보존됩니다.');
  r++;
  blank(r); r++;

  // ─ 섹션 5 ─
  title(r, '  ⚠️  주의사항'); r++;
  blank(r); r++;

  const warnings = [
    ['기록 초기화는 되돌릴 수 없습니다', '초기화 전 반드시 시트를 다운로드하거나\n중요한 데이터를 별도로 저장해두세요.'],
    ['재배포 시 URL이 변경됩니다', 'Apps Script를 수정 후 "새 배포"를 하면 URL이 바뀝니다.\n기존 URL을 유지하려면 [배포 관리] → ✏️ 편집 → 버전: 새 버전 선택 후 배포하세요.\n같은 URL이 유지됩니다.'],
    ['액세스 권한 확인', '"모든 사용자(Anyone)" 설정을 반드시 확인하세요.\n학교 도메인 내부로만 설정하면 학생이 개인 구글 계정으로\n접속할 수 없어 데이터가 저장되지 않습니다.'],
    ['학생 명단 직접 편집 가능', '"학생 명단" 시트를 직접 수정하면 즉시 앱에 반영됩니다.\n학생 이름을 변경하면 기존 기록과 분리될 수 있으니\n학기 초에 한 번만 설정하는 것을 권장합니다.']
  ];
  warnings.forEach(function(w){ row2col(r, w[0], w[1]); r++; });
  blank(r); r++;

  // 하단 안내
  guide.getRange(r, 1, 1, 4).merge()
    .setValue('문의 및 피드백: 앱 화면 하단 [선생님 메뉴]에서 본부 URL을 확인하거나, 학교 담당자에게 문의하세요.')
    .setFontSize(9).setFontColor('#94a3b8').setBackground('#f1f5f9')
    .setHorizontalAlignment('center').setWrap(true);
  guide.setRowHeight(r, 30);

  // 가이드 시트를 맨 앞으로 이동
  ss.setActiveSheet(guide);
  ss.moveActiveSheet(1);
}
