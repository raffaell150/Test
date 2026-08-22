// ==========================================
// 1. 설정 및 초기화
// ==========================================
const GEMINI_API_KEY = "YOUR_GEMINI_API_KEY_HERE"; // Gemini API 키 (필요 시 수정)
const LOG_SHEET_NAME = 'Sheet1';                   // 일기/일정 저장 시트 (필요 시 '시트1'로 변경)
const REPEAT_SHEET_NAME = 'RepeatingTasks';         // 반복 일정 저장 시트

// 시트가 없으면 자동으로 생성하는 함수
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let repeatSheet = ss.getSheetByName(REPEAT_SHEET_NAME);
  if (!repeatSheet) {
    repeatSheet = ss.insertSheet(REPEAT_SHEET_NAME);
    repeatSheet.appendRow(['id', 'title', 'type', 'days']); // 헤더 추가
  }
}

// ==========================================
// 2. 웹 앱 진입점 (doGet / doPost)
// ==========================================

/**
 * GET 요청 처리 (웹 브라우저 접속)
 */
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

/**
 * POST 요청 처리 (회원가입 또는 외부 API 요청 수신)
 */
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
      regDate,
      schedule || "",
      diary || "",
      email || "",
      name || "",
      userTrait || "",
      aiResponse,
      "SUCCESS"
    ]);

    return ContentService.createTextOutput(JSON.stringify({
      result: "success",
      message: "처리 완료",
      aiResponse: aiResponse
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      result: "error",
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// 3. 일정 관리 전용 백엔드 함수들 (google.script.run 전용)
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
      return {
        tasks: data[i][1] || '',
        diary: data[i][2] || ''
      };
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
  
  return { success: true };
}

function loadRepeatingTasks() {
  initSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(REPEAT_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  const tasks = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      tasks.push({
        id: data[i][0],
        title: data[i][1],
        type: data[i][2],
        days: String(data[i][3]).split(',')
      });
    }
  }
  return tasks;
}

function addRepeatingTaskServer(task) {
  initSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(REPEAT_SHEET_NAME);
  
  sheet.appendRow([
    task.id,
    task.title,
    task.type,
    task.days.join(',')
  ]);
  
  return loadRepeatingTasks();
}

function deleteRepeatingTaskServer(id) {
  initSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(REPEAT_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  
  return loadRepeatingTasks();
}

// ==========================================
// 4. Gemini API 호출 전용 함수
// ==========================================

function generateAiResponse(userName, userTrait, userDiary) {
  if (!GEMINI_API_KEY || GEMINI_API_KEY === "YOUR_GEMINI_API_KEY_HERE") {
    return "API 키가 설정되지 않아 기본 답장을 반환합니다.";
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`;
  
  const promptText = `
    당신은 사용자 맞춤형 메시지를 작성하는 AI 에이전트입니다.
    아래 사용자의 정보를 바탕으로 맞춤형 답장을 3-4문장으로 작성해 주세요.
    - 사용자 이름: ${userName || "회원"}
    - 사용자 성향: ${userTrait || "정보 없음"}
    - 일기/기록 내용: ${userDiary || "정보 없음"}
  `;

  const payload = {
    contents: [{ parts: [{ text: promptText }] }]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    if (json.candidates && json.candidates[0].content.parts[0].text) {
      return json.candidates[0].content.parts[0].text.trim();
    } else {
      return "맞춤 메시지를 생성하는 중에 문제가 발생했습니다.";
    }
  } catch (err) {
    return "AI 오류: " + err.toString();
  }
}