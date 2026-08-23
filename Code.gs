// ==========================================
// Google Apps Script 백엔드 메인 로직
// ==========================================

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('일정 & 일기 관리')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 1. 구글 시트 초기화 함수 (initSheets 복구)
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Events 시트 확인 및 생성
  let eventsSheet = ss.getSheetByName('Events');
  if (!eventsSheet) {
    eventsSheet = ss.insertSheet('Events');
    eventsSheet.appendRow(['ID', 'Date', 'Title', 'Category']);
    eventsSheet.getRange('A1:D1').setFontWeight('bold');
  }

  // Diaries 시트 확인 및 생성
  let diariesSheet = ss.getSheetByName('Diaries');
  if (!diariesSheet) {
    diariesSheet = ss.insertSheet('Diaries');
    diariesSheet.appendRow(['ID', 'Date', 'Content']);
    diariesSheet.getRange('A1:C1').setFontWeight('bold');
  }

  return "시트 초기화 완료";
}

// 2. 초기 데이터 일괄 수신 (일정 + 일기)
function getInitialData() {
  try {
    // 필요시 자동 시트 생성
    initSheets();

    const events = getEventsFromSheet();
    const diaries = getDiariesFromSheet();
    return {
      success: true,
      events: events || [],
      diaries: diaries || []
    };
  } catch (e) {
    return {
      success: false,
      error: e.toString(),
      events: [],
      diaries: []
    };
  }
}

// 3. 구글 시트에서 일정 읽기
function getEventsFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Events');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const events = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0]) { // ID가 존재하는 경우
      let dateStr = row[1];
      if (dateStr instanceof Date) {
        const y = dateStr.getFullYear();
        const m = String(dateStr.getMonth() + 1).padStart(2, '0');
        const d = String(dateStr.getDate()).padStart(2, '0');
        dateStr = `${y}-${m}-${d}`;
      }
      events.push({
        id: String(row[0]),
        date: String(dateStr),
        title: String(row[2] || ''),
        category: String(row[3] || '기타')
      });
    }
  }
  return events;
}

// 4. 구글 시트에서 일기 읽기
function getDiariesFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Diaries');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const diaries = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0]) {
      let dateStr = row[1];
      if (dateStr instanceof Date) {
        const y = dateStr.getFullYear();
        const m = String(dateStr.getMonth() + 1).padStart(2, '0');
        const d = String(dateStr.getDate()).padStart(2, '0');
        dateStr = `${y}-${m}-${d}`;
      }
      diaries.push({
        id: String(row[0]),
        date: String(dateStr),
        content: String(row[2] || '')
      });
    }
  }
  return diaries;
}

// 5. 일정 저장/수정
function saveEventToSheet(eventData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Events');
  if (!sheet) {
    initSheets();
    sheet = ss.getSheetByName('Events');
  }

  const data = sheet.getDataRange().getValues();
  let foundIndex = -1;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(eventData.id)) {
      foundIndex = i + 1;
      break;
    }
  }

  if (foundIndex > 0) {
    sheet.getRange(foundIndex, 1, 1, 4).setValues([[eventData.id, eventData.date, eventData.title, eventData.category]]);
  } else {
    sheet.appendRow([eventData.id, eventData.date, eventData.title, eventData.category]);
  }
  return getInitialData();
}

// 6. 일정 삭제
function deleteEventFromSheet(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Events');
  if (!sheet) return getInitialData();

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return getInitialData();
}