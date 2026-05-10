/* 
  SDG's 리틀 히어로 - Google Apps Script Backend (code.gs)
  1. Open Google Sheets.
  2. Extensions > Apps Script.
  3. Paste this code.
  4. Deploy > New Deployment > Web App (Set "Who has access" to "Anyone").
  5. Copy the Web App URL and paste it into script.js.
*/

const SHEET_NAME = 'SDG_Records';
const SHEET_STUDENTS = 'Students'; // New sheet for student names

function doGet(e) {
  const action = e.parameter.action;
  const heroName = e.parameter.heroName;

  // 1. Fetch student list
  if (action === 'getStudents') {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Try possible sheet names: Students, 학생명단, 명단
    const possibleNames = [SHEET_STUDENTS, '학생명단', '명단', '학생', '학생명부'];
    let sheet = null;
    for (let name of possibleNames) {
      sheet = ss.getSheetByName(name);
      if (sheet) break;
    }

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_STUDENTS);
      sheet.appendRow(["학생 이름"]);
      sheet.appendRow(["용기있는 사자"]);
      sheet.appendRow(["지혜로운 코끼리"]);
    }
    
    // Get all non-header values from first column
    const names = sheet.getDataRange().getValues().slice(1).map(row => row[0]).filter(n => n && n !== "");

    return ContentService.createTextOutput(JSON.stringify({ students: names }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 1.5 Fetch class-wide aggregated statistics
  if (action === 'getClassStats') {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({ stats: null }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const data = sheet.getDataRange().getValues().slice(1);
    const byHero = {};
    const byDate = {};
    let totalRecords = 0;
    let totalProgressSum = 0;
    let totalTasks = 0;

    data.forEach(row => {
      const ts = String(row[0] || '');
      const name = row[1];
      const progress = parseInt(row[2]) || 0;
      const tasksStr = String(row[3] || '');
      if (!name) return;

      const tasksCount = tasksStr.split('\n').filter(t => t.trim().length > 0).length;

      if (!byHero[name]) byHero[name] = { records: 0, progressSum: 0, latestProgress: 0, lastTimestamp: '' };
      byHero[name].records++;
      byHero[name].progressSum += progress;
      byHero[name].latestProgress = progress;
      byHero[name].lastTimestamp = ts;

      const dateMatch = ts.match(/(\d+)\.\s*(\d+)\.\s*(\d+)/);
      if (dateMatch) {
        const yyyy = dateMatch[1];
        const mm = ('0' + dateMatch[2]).slice(-2);
        const dd = ('0' + dateMatch[3]).slice(-2);
        const dateKey = yyyy + '-' + mm + '-' + dd;
        const label = parseInt(dateMatch[2], 10) + '/' + parseInt(dateMatch[3], 10);
        if (!byDate[dateKey]) byDate[dateKey] = { dateKey: dateKey, label: label, totalTasks: 0, studentLatest: {}, participants: {} };
        byDate[dateKey].totalTasks += tasksCount;
        byDate[dateKey].studentLatest[name] = progress; // last row wins (sheet append order)
        byDate[dateKey].participants[name] = true;
      }

      totalRecords++;
      totalProgressSum += progress;
      totalTasks += tasksCount;
    });

    const heroStats = Object.keys(byHero).map(name => ({
      name: name,
      records: byHero[name].records,
      avgProgress: Math.round(byHero[name].progressSum / byHero[name].records),
      latestProgress: byHero[name].latestProgress
    })).sort(function(a, b) { return b.avgProgress - a.avgProgress; });

    const dailyStats = Object.keys(byDate).sort().map(function(k) {
      const scores = Object.keys(byDate[k].studentLatest).map(function(n) { return byDate[k].studentLatest[n]; });
      const avg = scores.length > 0
        ? Math.round(scores.reduce(function(a, b) { return a + b; }, 0) / scores.length)
        : 0;
      return {
        dateKey: byDate[k].dateKey,
        label: byDate[k].label,
        totalTasks: byDate[k].totalTasks,
        avgProgress: avg,
        participants: Object.keys(byDate[k].participants).length
      };
    }).filter(function(d) { return d.avgProgress > 0; });

    return ContentService.createTextOutput(JSON.stringify({
      stats: {
        totalStudents: heroStats.length,
        totalRecords: totalRecords,
        totalTasks: totalTasks,
        avgProgress: totalRecords ? Math.round(totalProgressSum / totalRecords) : 0,
        heroStats: heroStats,
        dailyStats: dailyStats
      }
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // 2. Fetch history for a specific hero
  if (heroName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return ContentService.createTextOutput(JSON.stringify({ history: [] })).setMimeType(ContentService.MimeType.JSON);
    
    const data = sheet.getDataRange().getValues();
    const history = data.slice(1) // Skip header
      .filter(row => row[1] === heroName)
      .map(row => ({
        timestamp: row[0],
        heroName: row[1],
        progress: parseInt(row[2]),
        tasks: row[3],
        reflection: row[4],
        myGoal: row[5]
      }))
      .reverse(); 
      
    return ContentService.createTextOutput(JSON.stringify({ history: history }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 3. Default: serve the app HTML
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle("SDG's 리틀 히어로")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 2.5 Generate dashboard sheet
    if (data.action === 'generateDashboard') {
      return generateDashboardSheet();
    }

    // 3. Clear record data
    if (data.action === 'clearData') {
      let sheet = ss.getSheetByName(SHEET_NAME);
      if (sheet) {
        sheet.clearContents();
        sheet.appendRow(["기록 시간", "히어로 이름", "성장 지수 (%)", "실천한 내용", "느낀 점", "나의 다짐"]);
        sheet.getRange(1, 1, 1, 6).setBackground("#10b981").setFontColor("white").setFontWeight("bold");
      }
      return ContentService.createTextOutput(JSON.stringify({ "result": "success", "message": "Records cleared" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 4. Save new record
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(["기록 시간", "히어로 이름", "성장 지수 (%)", "실천한 내용", "느낀 점", "나의 다짐"]);
      sheet.getRange(1, 1, 1, 6).setBackground("#10b981").setFontColor("white").setFontWeight("bold");
    }
    
    sheet.appendRow([
      data.timestamp || new Date().toLocaleString('ko-KR'),
      data.heroName,
      data.progress,
      data.tasks,
      data.reflection,
      data.myGoal || ''
    ]);

    // Auto-refresh dashboard (throttled to avoid overload during simultaneous saves)
    try { maybeRegenerateDashboard(); } catch (dashErr) { /* ignore */ }

    return ContentService.createTextOutput(JSON.stringify({ "result": "success" }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ "result": "error", "message": err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function maybeRegenerateDashboard() {
  const props = PropertiesService.getScriptProperties();
  const lastStr = props.getProperty('dashboard_last_run');
  const now = Date.now();
  if (lastStr) {
    const last = parseInt(lastStr, 10);
    if (now - last < 30000) return; // throttle: skip if last run < 30 sec ago
  }
  props.setProperty('dashboard_last_run', String(now));
  generateDashboardSheet();
}

function generateDashboardSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(SHEET_NAME);
    if (!sourceSheet) {
      return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: '기록 시트가 없습니다.' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const data = sourceSheet.getDataRange().getValues().slice(1);
    if (data.length === 0) {
      return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: '집계할 데이터가 없습니다.' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const byHero = {};
    const byDate = {};
    let totalRecords = 0;
    let totalProgressSum = 0;
    let totalTasks = 0;

    data.forEach(function(row) {
      const ts = String(row[0] || '');
      const name = row[1];
      const progress = parseInt(row[2]) || 0;
      const tasksStr = String(row[3] || '');
      if (!name) return;

      const tasksCount = tasksStr.split('\n').filter(function(t) { return t.trim().length > 0; }).length;

      if (!byHero[name]) byHero[name] = { records: 0, progressSum: 0, latestProgress: 0, lastTimestamp: '' };
      byHero[name].records++;
      byHero[name].progressSum += progress;
      byHero[name].latestProgress = progress;
      byHero[name].lastTimestamp = ts;

      const dateMatch = ts.match(/(\d+)\.\s*(\d+)\.\s*(\d+)/);
      if (dateMatch) {
        const yyyy = dateMatch[1];
        const mm = ('0' + dateMatch[2]).slice(-2);
        const dd = ('0' + dateMatch[3]).slice(-2);
        const dateKey = yyyy + '-' + mm + '-' + dd;
        const label = parseInt(dateMatch[2], 10) + '월 ' + parseInt(dateMatch[3], 10) + '일';
        if (!byDate[dateKey]) byDate[dateKey] = { dateKey: dateKey, label: label, totalTasks: 0, studentLatest: {}, participants: {} };
        byDate[dateKey].totalTasks += tasksCount;
        byDate[dateKey].studentLatest[name] = progress; // last row wins
        byDate[dateKey].participants[name] = true;
      }

      totalRecords++;
      totalProgressSum += progress;
      totalTasks += tasksCount;
    });

    const heroStats = Object.keys(byHero).map(function(name) {
      return {
        name: name,
        records: byHero[name].records,
        avgProgress: Math.round(byHero[name].progressSum / byHero[name].records),
        latestProgress: byHero[name].latestProgress,
        lastTimestamp: byHero[name].lastTimestamp
      };
    }).sort(function(a, b) { return b.avgProgress - a.avgProgress; });

    const dailyStats = Object.keys(byDate).sort().map(function(k) {
      const scores = Object.keys(byDate[k].studentLatest).map(function(n) { return byDate[k].studentLatest[n]; });
      const avg = scores.length > 0
        ? Math.round(scores.reduce(function(a, b) { return a + b; }, 0) / scores.length)
        : 0;
      return {
        label: byDate[k].label,
        totalTasks: byDate[k].totalTasks,
        avgProgress: avg,
        participants: Object.keys(byDate[k].participants).length
      };
    }).filter(function(d) { return d.avgProgress > 0; });

    const totalStudents = heroStats.length;
    const avgProgress = totalRecords ? Math.round(totalProgressSum / totalRecords) : 0;

    const DASHBOARD_NAME = '📊 대시보드';
    let dash = ss.getSheetByName(DASHBOARD_NAME);
    if (dash) {
      dash.getCharts().forEach(function(c) { dash.removeChart(c); });
      dash.clear();
    } else {
      dash = ss.insertSheet(DASHBOARD_NAME, 0);
    }

    // Clean look: hide gridlines, set tab color, freeze title row
    dash.setHiddenGridlines(true);
    try { dash.setTabColor('#10b981'); } catch (e) { /* older API */ }
    dash.setFrozenRows(2);

    dash.getRange('A1').setValue('📊 우리 반 SDG 히어로 대시보드')
      .setFontSize(20).setFontWeight('bold').setFontColor('#10b981');
    dash.getRange('A2').setValue('마지막 업데이트: ' + new Date().toLocaleString('ko-KR'))
      .setFontSize(10).setFontColor('#64748b').setFontStyle('italic');

    dash.getRange('A4').setValue('▼ 요약 통계')
      .setFontSize(13).setFontWeight('bold').setFontColor('#0369a1');
    dash.getRange('A5:D5').setValues([['참여 친구 수', '전체 기록 수', '실천 횟수', '평균 성장률']])
      .setBackground('#0369a1').setFontColor('white').setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    dash.getRange('A6:D6').setValues([[totalStudents + '명', totalRecords + '개', totalTasks + '회', avgProgress + '%']])
      .setBackground('#f0f9ff').setFontSize(16).setFontWeight('bold')
      .setHorizontalAlignment('center').setFontColor('#0369a1');
    dash.setRowHeight(5, 30);
    dash.setRowHeight(6, 45);

    let row = 8;
    dash.getRange(row, 1).setValue('▼ 친구별 성장 통계 (평균 성장률 높은 순)')
      .setFontSize(13).setFontWeight('bold').setFontColor('#0369a1');
    row++;
    dash.getRange(row, 1, 1, 5).setValues([['이름', '기록 수', '평균 성장률(%)', '최근 성장률(%)', '마지막 기록']])
      .setBackground('#0369a1').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center');
    row++;

    const heroStartRow = row;
    if (heroStats.length > 0) {
      const heroData = heroStats.map(function(h) {
        return [h.name, h.records, h.avgProgress, h.latestProgress, h.lastTimestamp];
      });
      dash.getRange(row, 1, heroData.length, 5).setValues(heroData)
        .setHorizontalAlignment('center');
      dash.getRange(row, 1, heroData.length, 1).setHorizontalAlignment('left').setFontWeight('bold');
      row += heroData.length;
    }

    row += 2;
    dash.getRange(row, 1).setValue('▼ 날짜별 우리 반 평균 성장률')
      .setFontSize(13).setFontWeight('bold').setFontColor('#0369a1');
    row++;
    dash.getRange(row, 1, 1, 4).setValues([['날짜', '평균 성장률(%)', '실천 횟수', '참여 인원']])
      .setBackground('#0369a1').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center');
    row++;

    const dateStartRow = row;
    if (dailyStats.length > 0) {
      const dateData = dailyStats.map(function(d) {
        return [d.label, (typeof d.avgProgress === 'number' ? d.avgProgress : 0), d.totalTasks, d.participants];
      });
      dash.getRange(row, 1, dateData.length, 4).setValues(dateData).setHorizontalAlignment('center');
      row += dateData.length;
    }

    dash.setColumnWidth(1, 160);
    dash.setColumnWidth(2, 100);
    dash.setColumnWidth(3, 130);
    dash.setColumnWidth(4, 130);
    dash.setColumnWidth(5, 200);

    // ─ Date chart on TOP (line chart, daily class avg growth) ─
    if (dailyStats.length > 0) {
      const dateChart = dash.newChart()
        .asLineChart()
        .addRange(dash.getRange(dateStartRow, 1, dailyStats.length, 1))
        .addRange(dash.getRange(dateStartRow, 2, dailyStats.length, 1))
        .setPosition(4, 7, 0, 0)
        .setOption('title', '날짜별 우리 반 평균 성장률 (%)')
        .setOption('legend', { position: 'none' })
        .setOption('width', 520)
        .setOption('height', 320)
        .setOption('colors', ['#10b981'])
        .setOption('curveType', 'function')
        .setOption('pointSize', 6)
        .setOption('lineWidth', 3)
        .setOption('vAxis', { title: '평균 성장률 (%)', minValue: 0, maxValue: 100 })
        .setOption('hAxis', { title: '날짜' })
        .build();
      dash.insertChart(dateChart);
    }

    // ─ Hero chart BELOW the date chart (bar chart, top 5 friends) ─
    if (heroStats.length > 0) {
      const chartRowCount = Math.min(5, heroStats.length);
      const chartTitle = heroStats.length > 5
        ? '친구별 평균 성장률 TOP 5 (%)'
        : '친구별 평균 성장률 (%)';
      const heroChart = dash.newChart()
        .asBarChart()
        .addRange(dash.getRange(heroStartRow, 1, chartRowCount, 1))
        .addRange(dash.getRange(heroStartRow, 3, chartRowCount, 1))
        .setPosition(22, 7, 0, 0)
        .setOption('title', chartTitle)
        .setOption('legend', { position: 'none' })
        .setOption('width', 520)
        .setOption('height', Math.max(280, chartRowCount * 55 + 120))
        .setOption('colors', ['#10b981'])
        .setOption('hAxis', { title: '평균 성장률 (%)', minValue: 0, maxValue: 100 })
        .setOption('vAxis', { textStyle: { fontSize: 12, bold: true } })
        .build();
      dash.insertChart(heroChart);
    }

    ss.setActiveSheet(dash);

    return ContentService.createTextOutput(JSON.stringify({ result: 'success', message: 'Dashboard generated' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ===== Spreadsheet UI: top menu + sidebar =====

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('🌍 SDG 히어로')
    .addItem('📈 요약 통계 사이드바 열기', 'showStatsSidebar')
    .addSeparator()
    .addItem('🔄 대시보드 시트 새로고침', 'refreshDashboardFromMenu')
    .addItem('📊 대시보드 시트로 이동', 'openDashboardSheet')
    .addToUi();
}

function refreshDashboardFromMenu() {
  const ui = SpreadsheetApp.getUi();
  try {
    generateDashboardSheet();
    ui.alert('✅ 대시보드가 갱신되었습니다.');
  } catch (err) {
    ui.alert('❌ 오류가 발생했습니다.\n\n' + (err && err.message ? err.message : err));
  }
}

function openDashboardSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dash = ss.getSheetByName('📊 대시보드');
  if (!dash) {
    SpreadsheetApp.getUi().alert('아직 대시보드 시트가 없습니다.\n"🔄 대시보드 시트 새로고침"을 먼저 눌러주세요.');
    return;
  }
  ss.setActiveSheet(dash);
}

function computeClassStats_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues().slice(1);
  if (data.length === 0) return null;

  const byHero = {};
  const byDate = {};
  let totalRecords = 0, totalProgressSum = 0, totalTasks = 0;

  data.forEach(function(row) {
    const ts = String(row[0] || '');
    const name = row[1];
    const progress = parseInt(row[2]) || 0;
    const tasksStr = String(row[3] || '');
    if (!name) return;
    const tasksCount = tasksStr.split('\n').filter(function(t) { return t.trim().length > 0; }).length;

    if (!byHero[name]) byHero[name] = { records: 0, progressSum: 0, latestProgress: 0 };
    byHero[name].records++;
    byHero[name].progressSum += progress;
    byHero[name].latestProgress = progress;

    const dateMatch = ts.match(/(\d+)\.\s*(\d+)\.\s*(\d+)/);
    if (dateMatch) {
      const dateKey = dateMatch[1] + '-' + ('0' + dateMatch[2]).slice(-2) + '-' + ('0' + dateMatch[3]).slice(-2);
      const label = parseInt(dateMatch[2], 10) + '/' + parseInt(dateMatch[3], 10);
      if (!byDate[dateKey]) byDate[dateKey] = { label: label, totalTasks: 0, participants: {} };
      byDate[dateKey].totalTasks += tasksCount;
      byDate[dateKey].participants[name] = true;
    }

    totalRecords++;
    totalProgressSum += progress;
    totalTasks += tasksCount;
  });

  const heroStats = Object.keys(byHero).map(function(name) {
    return { name: name, records: byHero[name].records,
      avgProgress: Math.round(byHero[name].progressSum / byHero[name].records),
      latestProgress: byHero[name].latestProgress };
  }).sort(function(a, b) { return b.avgProgress - a.avgProgress; });

  const dailyStats = Object.keys(byDate).sort().map(function(k) {
    return { label: byDate[k].label, totalTasks: byDate[k].totalTasks,
      participants: Object.keys(byDate[k].participants).length };
  });

  return {
    totalStudents: heroStats.length,
    totalRecords: totalRecords,
    totalTasks: totalTasks,
    avgProgress: totalRecords ? Math.round(totalProgressSum / totalRecords) : 0,
    heroStats: heroStats,
    dailyStats: dailyStats
  };
}

function showStatsSidebar() {
  const stats = computeClassStats_();
  const html = HtmlService.createHtmlOutput(buildSidebarHtml_(stats))
    .setTitle('🌍 우리 반 SDG 통계');
  SpreadsheetApp.getUi().showSidebar(html);
}

function buildSidebarHtml_(stats) {
  if (!stats) {
    return '<div style="font-family:sans-serif; padding:20px; text-align:center; color:#64748b;">아직 데이터가 없어요.<br><br>학생이 한 번이라도 저장하면<br>통계가 표시됩니다 ✨</div>';
  }
  const top5 = stats.heroStats.slice(0, 5);
  const recent5 = stats.dailyStats.slice(-5);
  const safeName = function(n) { return String(n).replace(/[<>&"]/g, function(c){return ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;'}[c]);}); };

  const summaryCards = [
    { num: stats.totalStudents, label: '참여 친구', color: '#10b981' },
    { num: stats.totalRecords, label: '전체 기록', color: '#0ea5e9' },
    { num: stats.totalTasks, label: '실천 횟수', color: '#f59e0b' },
    { num: stats.avgProgress + '%', label: '평균 성장', color: '#8b5cf6' }
  ].map(function(c) {
    return '<div style="background:white; border:1px solid #e2e8f0; border-radius:10px; padding:10px 6px; text-align:center;">' +
      '<div style="font-size:1.4rem; font-weight:900; color:' + c.color + '; line-height:1.1;">' + c.num + '</div>' +
      '<div style="font-size:0.7rem; color:#64748b; font-weight:700; margin-top:4px;">' + c.label + '</div>' +
      '</div>';
  }).join('');

  const top5Html = top5.length ? top5.map(function(h, i) {
    const medal = ['🥇','🥈','🥉','4️⃣','5️⃣'][i] || '•';
    return '<div style="display:flex; align-items:center; gap:8px; padding:6px 8px; background:#f8fafc; border-radius:8px; margin-bottom:4px;">' +
      '<span style="font-size:0.95rem;">' + medal + '</span>' +
      '<span style="flex:1; font-size:0.78rem; font-weight:700; color:#0f172a; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">' + safeName(h.name) + '</span>' +
      '<span style="font-size:0.78rem; font-weight:900; color:#10b981;">' + h.avgProgress + '%</span>' +
      '</div>';
  }).join('') : '<div style="color:#94a3b8; font-size:0.78rem; text-align:center; padding:8px;">기록 없음</div>';

  const recentHtml = recent5.length ? recent5.map(function(d) {
    return '<div style="display:flex; justify-content:space-between; padding:5px 8px; background:#f8fafc; border-radius:8px; margin-bottom:3px;">' +
      '<span style="font-size:0.78rem; font-weight:700; color:#475569;">' + d.label + '</span>' +
      '<span style="font-size:0.78rem; color:#475569;">실천 <b style="color:#0369a1;">' + d.totalTasks + '</b> · 참여 ' + d.participants + '명</span>' +
      '</div>';
  }).join('') : '<div style="color:#94a3b8; font-size:0.78rem; text-align:center; padding:8px;">기록 없음</div>';

  return '<div style="font-family: -apple-system, BlinkMacSystemFont, \'Segoe UI\', \'Noto Sans KR\', sans-serif; padding:14px 12px; background:#f1f5f9; min-height:100vh; color:#0f172a;">' +
    '<h2 style="font-size:1rem; color:#10b981; margin:0 0 10px; font-weight:900;">🌍 우리 반 현황</h2>' +
    '<div style="display:grid; grid-template-columns:1fr 1fr; gap:6px; margin-bottom:14px;">' + summaryCards + '</div>' +
    '<h3 style="font-size:0.85rem; color:#0369a1; margin:0 0 6px; font-weight:900;">🏆 친구별 TOP 5</h3>' +
    '<div style="margin-bottom:14px;">' + top5Html + '</div>' +
    '<h3 style="font-size:0.85rem; color:#0369a1; margin:0 0 6px; font-weight:900;">📅 최근 5일 실천</h3>' +
    '<div style="margin-bottom:14px;">' + recentHtml + '</div>' +
    '<button onclick="google.script.run.withSuccessHandler(function(){ google.script.run.showStatsSidebar(); }).generateDashboardSheet()" style="width:100%; background:#10b981; color:white; border:none; border-radius:8px; padding:9px; font-weight:800; cursor:pointer; font-size:0.82rem; margin-bottom:6px;">🔄 통계 새로고침 + 대시보드 갱신</button>' +
    '<button onclick="google.script.run.openDashboardSheet()" style="width:100%; background:white; color:#475569; border:1px solid #cbd5e1; border-radius:8px; padding:9px; font-weight:800; cursor:pointer; font-size:0.82rem;">📊 대시보드 시트로 이동</button>' +
    '<p style="font-size:0.65rem; color:#94a3b8; text-align:center; margin-top:10px;">학생 저장 시 자동으로 갱신됩니다</p>' +
    '</div>';
}