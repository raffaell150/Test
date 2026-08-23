// ==========================================
// 1. 설정 및 초기화
// ==========================================
const GEMINI_API_KEY = "YOUR_GEMINI_API_KEY_HERE";
const LOG_SHEET_NAME = 'Sheet1';
const REPEAT_SHEET_NAME = 'RepeatingTasks';

function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let repeatSheet = ss.getSheetByName(REPEAT_SHEET_NAME);
  if (!repeatSheet) {
    repeatSheet = ss.insertSheet(REPEAT_SHEET_NAME);
    repeatSheet.appendRow(['id', 'title', 'type', 'days']);
  }
}

// ==========================================
// 2. 웹 앱 진입점
// ==========================================
function doGet(e) {
  try {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('일정 및 기록 관리')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    return ContentService.createTextOutput("Google Apps Script Web App is running correctly.")
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

function doPost(e) {
  try {
    let data;
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      data = e.parameter;
    } else {
      throw new Error("전송된 데이터가 없습니다.");
    }

    const { date, schedule, diary, email, name, userTrait } = data;
    const aiResponse = generateAiResponse(name, userTrait, diary);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(LOG_SHEET_NAME) || ss.getSheets()[0];
    const now = new Date();
    const regDate = date || Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");

    sheet.appendRow([
      regDate, schedule || "", diary || "", email || "", name || "", userTrait || "", aiResponse, "SUCCESS"
    ]);

    return ContentService.createTextOutput(JSON.stringify({
      result: "success", message: "처리 완료", aiResponse: aiResponse
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      result: "error", error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// 3. 일정 및 반복 일정 관리 백엔드 함수
// ==========================================
function loadDataByDate(dateStr) {
  initSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LOG_SHEET_NAME) || ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const rowDate = data[i][0];
    let formattedDate = '';
    
    if (rowDate instanceof Date) {
      formattedDate = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else {
      formattedDate = String(rowDate).trim();
    }
    
    if (formattedDate === dateStr) {
      return { tasks: data[i][1] || '', diary: data[i][2] || '' };
    }
  }
  return { tasks: '', diary: '' };
}

function loadWeeklyData(dateList) {
  const result = {};
  for (let i = 0; i < dateList.length; i++) {
    const dateStr = dateList[i];
    result[dateStr] = loadDataByDate(dateStr);
  }
  return result;
}

function saveData(e) {
  initSheets();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(LOG_SHEET_NAME) || ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    const targetDate = e.date;
    
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      const rowDate = data[i][0];
      let formattedDate = '';
      if (rowDate instanceof Date) {
        formattedDate = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        formattedDate = String(rowDate).trim();
      }
      if (formattedDate === targetDate) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 2).setValue(e.tasks);
      sheet.getRange(rowIndex, 3).setValue(e.diary);
    } else {
      sheet.appendRow([targetDate, e.tasks, e.diary]);
    }
    
    SpreadsheetApp.flush();
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

// [핵심 해결] 반복 일정 불러오기
function loadRepeatingTasks() {
  initSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(REPEAT_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  const tasks = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      let daysArray = [];
      const rawDays = String(data[i][3] || '');
      
      try {
        daysArray = JSON.parse(rawDays);
      } catch (e) {
        daysArray = rawDays ? rawDays.split(',') : [];
      }

      tasks.push({
        id: String(data[i][0]),
        title: String(data[i][1]),
        type: String(data[i][2]),
        days: Array.isArray(daysArray) ? daysArray : [daysArray]
      });
    }
  }
  return tasks;
}

// [핵심 해결] 반복 일정 서버 저장
function addRepeatingTaskServer(task) {
  initSheets();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // 락으로 동시 접근 충돌 방지
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(REPEAT_SHEET_NAME);
    
    // days 배열 처리
    const daysStr = Array.isArray(task.days) ? task.days.join(',') : String(task.days);

    sheet.appendRow([
      task.id,
      task.title,
      task.type,
      daysStr
    ]);
    
    SpreadsheetApp.flush(); // ★ 핵심: 구글 서버 디스크에 즉시 저장 강제
    return loadRepeatingTasks(); // 저장 완료 후 최신 목록 반환
  } finally {
    lock.releaseLock();
  }
}

function deleteRepeatingTaskServer(id) {
  initSheets();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(REPEAT_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    // 행을 뒤에서부터 탐색하여 삭제 (행 삭제 시 인덱스 꼬임 방지)
    for (let i = data.length - 1; i >= 1; i--) {
      // String()으로 변환 후 trim()하여 숫자/문자열 타입 차이 및 공백 무시 비교
      if (String(data[i][0]).trim() === String(id).trim()) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    
    SpreadsheetApp.flush(); // ★ 시트에 삭제 결과 즉시 물리적 저장
    return loadRepeatingTasks(); // 삭제 후 최신 목록 전체 반환
  } catch (e) {
    Logger.log("삭제 중 에러: " + e.toString());
    return loadRepeatingTasks();
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 4. Gemini API
// ==========================================
function generateAiResponse(userName, userTrait, userDiary) {
  if (!GEMINI_API_KEY || GEMINI_API_KEY === "YOUR_GEMINI_API_KEY_HERE") {
    return "API 키가 설정되지 않아 기본 답장을 반환합니다.";
  }
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`;
  const promptText = `사용자:${userName || "회원"}, 성향:${userTrait || "없음"}, 일기:${userDiary || "없음"}`;
  const payload = { contents: [{ parts: [{ text: promptText }] }] };
  const options = { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    return json.candidates?.[0]?.content?.parts?.[0]?.text?.trim() || "응답 생성 실패";
  } catch (err) {
    return "AI 오류: " + err.toString();
  }
}