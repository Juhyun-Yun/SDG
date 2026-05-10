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
        if (!byDate[dateKey]) byDate[dateKey] = { dateKey: dateKey, label: label, totalTasks: 0, participants: {} };
        byDate[dateKey].totalTasks += tasksCount;
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
      return {
        dateKey: byDate[k].dateKey,
        label: byDate[k].label,
        totalTasks: byDate[k].totalTasks,
        participants: Object.keys(byDate[k].participants).length
      };
    });

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
        if (!byDate[dateKey]) byDate[dateKey] = { dateKey: dateKey, label: label, totalTasks: 0, participants: {} };
        byDate[dateKey].totalTasks += tasksCount;
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
      return {
        label: byDate[k].label,
        totalTasks: byDate[k].totalTasks,
        participants: Object.keys(byDate[k].participants).length
      };
    });

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
    dash.getRange(row, 1).setValue('▼ 날짜별 실천 횟수')
      .setFontSize(13).setFontWeight('bold').setFontColor('#0369a1');
    row++;
    dash.getRange(row, 1, 1, 3).setValues([['날짜', '실천 횟수', '참가자 수']])
      .setBackground('#0369a1').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center');
    row++;

    const dateStartRow = row;
    if (dailyStats.length > 0) {
      const dateData = dailyStats.map(function(d) {
        return [d.label, d.totalTasks, d.participants];
      });
      dash.getRange(row, 1, dateData.length, 3).setValues(dateData).setHorizontalAlignment('center');
      row += dateData.length;
    }

    dash.setColumnWidth(1, 160);
    dash.setColumnWidth(2, 100);
    dash.setColumnWidth(3, 130);
    dash.setColumnWidth(4, 130);
    dash.setColumnWidth(5, 200);

    if (heroStats.length > 0) {
      const heroChart = dash.newChart()
        .asBarChart()
        .addRange(dash.getRange(heroStartRow, 1, heroStats.length, 1))
        .addRange(dash.getRange(heroStartRow, 3, heroStats.length, 1))
        .setPosition(heroStartRow, 7, 0, 0)
        .setOption('title', '친구별 평균 성장률 (%)')
        .setOption('legend', { position: 'none' })
        .setOption('width', 520)
        .setOption('height', Math.max(220, heroStats.length * 35 + 100))
        .setOption('colors', ['#10b981'])
        .setOption('hAxis', { title: '평균 성장률 (%)', minValue: 0, maxValue: 100 })
        .build();
      dash.insertChart(heroChart);
    }

    if (dailyStats.length > 0) {
      const dateChart = dash.newChart()
        .asColumnChart()
        .addRange(dash.getRange(dateStartRow, 1, dailyStats.length, 1))
        .addRange(dash.getRange(dateStartRow, 2, dailyStats.length, 1))
        .setPosition(dateStartRow, 7, 0, 0)
        .setOption('title', '날짜별 실천 횟수')
        .setOption('legend', { position: 'none' })
        .setOption('width', 520)
        .setOption('height', 320)
        .setOption('colors', ['#6366f1'])
        .setOption('vAxis', { title: '실천 횟수', minValue: 0 })
        .build();
      dash.insertChart(dateChart);
    }

    ss.setActiveSheet(dash);

    return ContentService.createTextOutput(JSON.stringify({ result: 'success', message: 'Dashboard generated' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}