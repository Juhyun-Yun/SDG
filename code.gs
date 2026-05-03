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
