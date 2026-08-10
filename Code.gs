// ============================================================
// 補習社出席記錄系統 — Google Apps Script 後端
// 版本: 3.7 (加入學校欄位、學校篩選及編輯功能)
// ============================================================

const SPREADSHEET_ID = "1ibit_1LtSGfhaI2NuBIIJwL5fhYIVnM-gRruF44pLak";
const STUDENTS_SHEET = "學生名單";
const ATTENDANCE_SHEET = "出席記錄";

// ── CORS 輔助 ─────────────────────────────────────────────────
function corsResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 主入口：GET 請求 ──────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action || "";

  if (!action) {
    return HtmlService.createHtmlOutputFromFile("index")
      .setTitle("補習社出席系統")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  try {
    if (action === "getGrades") {
      return corsResponse(getGrades());
    }
    if (action === "getSchools") { // 新增：讀取學校清單
      return corsResponse(getSchools());
    }
    if (action === "getStudents") {
      return corsResponse(getStudents(e.parameter.grade, e.parameter.school));
    }
    if (action === "getAttendance") {
      return corsResponse(getAttendance(e.parameter.date, e.parameter.grade));
    }
    if (action === "getSummary") {
      return corsResponse(getSummary(e.parameter.startDate, e.parameter.endDate, e.parameter.grade));
    }
    if (action === "getTodayStatus") {
      return corsResponse(getTodayStatus(e.parameter.grade));
    }
    if (action === "getStudentAttendance") {
      return corsResponse(getStudentAttendance(e.parameter.grade, e.parameter.name, e.parameter.startDate, e.parameter.endDate));
    }
    if (action === "getStudentInfo") { // 更新：同時讀取備註及學校
      return corsResponse(getStudentInfo(e.parameter.grade, e.parameter.name));
    }
    return corsResponse({ success: false, error: "Unknown action: " + action });
  } catch (err) {
    return corsResponse({ success: false, error: err.message });
  }
}

// ── 主入口：POST 請求 ─────────────────────────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action || "";

    if (action === "saveAttendance") {
      return corsResponse(saveAttendance(payload.date, payload.grade, payload.records));
    }
    if (action === "addStudent") {
      return corsResponse(addStudent(payload.grade, payload.name, payload.remark, payload.school));
    }
    if (action === "deleteStudent") {
      return corsResponse(deleteStudent(payload.grade, payload.name));
    }
    if (action === "updateStudentInfo") { // 更新：同時更新備註及學校
      return corsResponse(updateStudentInfo(payload.grade, payload.name, payload.remark, payload.school));
    }
    if (action === "updateAttendanceRemark") {
      return corsResponse(updateAttendanceRemark(payload.date, payload.grade, payload.name, payload.remark));
    }
    if (action === "updateAttendanceStatus") {
      return corsResponse(updateAttendanceStatus(payload.date, payload.grade, payload.name, payload.status));
    }
    return corsResponse({ success: false, error: "Unknown action: " + action });
  } catch (err) {
    return corsResponse({ success: false, error: err.message });
  }
}

// ── 取得所有年級 ──────────────────────────────
function getGrades() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(STUDENTS_SHEET);
  const data = sheet.getDataRange().getValues();
  const grades = new Set();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) grades.add(data[i][0]);
  }
  const gradeOrder = ["F.1","F.2","F.3","F.4","F.5","F.6"];
  const sorted = gradeOrder.filter(g => grades.has(g));
  grades.forEach(g => { if (!gradeOrder.includes(g)) sorted.push(g); });
  return { success: true, grades: sorted };
}

// ── 取得所有學校清單 ──────────────────────────────
function getSchools() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(STUDENTS_SHEET);
  const data = sheet.getDataRange().getValues();
  const schools = new Set();
  for (let i = 1; i < data.length; i++) {
    if (data[i][3]) schools.add(data[i][3]); // 讀取 D 欄 (索引 3)
  }
  return { success: true, schools: Array.from(schools).sort() };
}

// ── 取得學生名單 (支援按學校篩選) ──────────────────────────────────
function getStudents(grade, school) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(STUDENTS_SHEET);
  const data = sheet.getDataRange().getValues();
  const students = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && !row[1]) continue;
    if (grade && row[0] !== grade) continue;
    if (school && row[3] !== school) continue; // 學校篩選
    students.push({ grade: row[0], name: row[1], remark: row[2] || "", school: row[3] || "" });
  }
  return { success: true, students };
}

// ── 儲存出席記錄 ─────────────────────────
function saveAttendance(date, grade, records) {
  if (!date || !records || records.length === 0) {
    return { success: false, error: "缺少必要資料" };
  }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(ATTENDANCE_SHEET);
  const data = sheet.getDataRange().getValues();
  const rowsToDelete = [];
  
  const gradesToReplace = new Set(records.map(r => r.grade));
  if (grade) gradesToReplace.add(grade);

  for (let i = data.length - 1; i >= 1; i--) {
    const rawDate = data[i][0];
    const rowDate = rawDate instanceof Date
      ? Utilities.formatDate(rawDate, "Asia/Hong_Kong", "yyyy-MM-dd")
      : String(rawDate).split("T")[0];
    if (rowDate === date && gradesToReplace.has(data[i][2])) {
      rowsToDelete.push(i + 1);
    }
  }
  rowsToDelete.forEach(r => sheet.deleteRow(r));

  const dateObj = new Date(date);
  const weekdays = ["日", "一", "二", "三", "四", "五", "六"];
  const weekday = "星期" + weekdays[dateObj.getDay()];
  const now = new Date();
  const timestamp = Utilities.formatDate(now, "Asia/Hong_Kong", "yyyy-MM-dd HH:mm:ss");

  const newRows = records.map(r => [
    date, weekday, r.grade || grade, r.name, r.status, r.remark || "", timestamp
  ]);

  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 7).setValues(newRows);
    const lastRow = sheet.getLastRow();
    const startRow = lastRow - newRows.length + 1;
    for (let i = 0; i < newRows.length; i++) {
      const statusCell = sheet.getRange(startRow + i, 5);
      if (newRows[i][4] === "出席") {
        statusCell.setBackground("#d9ead3").setFontColor("#274e13");
      } else {
        statusCell.setBackground("#fce8e6").setFontColor("#a61c00");
      }
    }
  }
  const presentCount = records.filter(r => r.status === "出席").length;
  const absentCount = records.filter(r => r.status === "缺席").length;
  return {
    success: true, message: `已儲存 ${records.length} 條記錄`,
    summary: { present: presentCount, absent: absentCount, total: records.length }
  };
}

// ── 取得指定日期出席記錄 (強制排序) ────────────────
function getAttendance(date, grade) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(ATTENDANCE_SHEET);
  const data = sheet.getDataRange().getValues();
  const records = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const rawDate = row[0];
    const rowDate = rawDate instanceof Date
      ? Utilities.formatDate(rawDate, "Asia/Hong_Kong", "yyyy-MM-dd")
      : String(rawDate).split("T")[0];
    if (date && rowDate !== date) continue;
    if (grade && row[2] !== grade) continue;
    records.push({
      date: rowDate, weekday: row[1], grade: row[2],
      name: row[3], status: row[4], remark: row[5] || "", timestamp: row[6] || ""
    });
  }

  const gradeOrder = ["F.1","F.2","F.3","F.4","F.5","F.6"];
  records.sort((a, b) => {
    const gIndexA = gradeOrder.indexOf(a.grade) !== -1 ? gradeOrder.indexOf(a.grade) : 99;
    const gIndexB = gradeOrder.indexOf(b.grade) !== -1 ? gradeOrder.indexOf(b.grade) : 99;
    if (gIndexA !== gIndexB) return gIndexA - gIndexB;
    return a.name.localeCompare(b.name, "zh-HK");
  });

  return { success: true, records };
}

function getTodayStatus(grade) {
  const today = Utilities.formatDate(new Date(), "Asia/Hong_Kong", "yyyy-MM-dd");
  return getAttendance(today, grade);
}

// ── 查詢某學生某時段出席記錄 ───────────────────────────
function getStudentAttendance(grade, name, startDate, endDate) {
  if (!name) return { success: false, error: "請提供學生姓名" };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(ATTENDANCE_SHEET);
  const data = sheet.getDataRange().getValues();
  const records = [];
  let presentCount = 0;
  let absentCount = 0;
  let foundGrade = grade || "未知年級";

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const rawDate = row[0];
    const rowDate = rawDate instanceof Date
      ? Utilities.formatDate(rawDate, "Asia/Hong_Kong", "yyyy-MM-dd")
      : String(rawDate).split("T")[0];

    if (grade && row[2] !== grade) continue;
    if (row[3] !== name) continue;
    if (startDate && rowDate < startDate) continue;
    if (endDate && rowDate > endDate) continue;

    foundGrade = row[2];
    records.push({ date: rowDate, weekday: row[1], status: row[4], remark: row[5] || "" });
    if (row[4] === "出席") presentCount++;
    else absentCount++;
  }
  records.sort((a, b) => b.date.localeCompare(a.date));
  const total = records.length;
  const rate = total > 0 ? Math.round((presentCount / total) * 100) : 0;
  return {
    success: true, student: { grade: foundGrade, name }, records,
    summary: { present: presentCount, absent: absentCount, total, rate },
    dateRange: { start: startDate || "", end: endDate || "" }
  };
}

// ── 取得某學生資訊 (包含學校) ──────────────────────────────────────
function getStudentInfo(grade, name) {
  if (!name) return { success: false, error: "請提供姓名" };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(STUDENTS_SHEET);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((!grade || data[i][0] === grade) && data[i][1] === name) {
      return { success: true, grade: data[i][0], name, remark: data[i][2] || "", school: data[i][3] || "" };
    }
  }
  return { success: false, error: "找不到該學生" };
}

// ── 更新學生資料 (備註及學校) ────────────────────────────
function updateStudentInfo(grade, name, remark, school) {
  if (!name) return { success: false, error: "請提供姓名" };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(STUDENTS_SHEET);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((!grade || data[i][0] === grade) && data[i][1] === name) {
      sheet.getRange(i + 1, 3).setValue(remark || "");
      sheet.getRange(i + 1, 4).setValue(school || ""); // 寫入 D 欄
      return { success: true, message: `已更新 ${name} 的資料` };
    }
  }
  return { success: false, error: "找不到該學生" };
}

// ── 更新某條出席記錄的備註 ─────────────────────────────
function updateAttendanceRemark(date, grade, name, remark) {
  if (!date || !name) return { success: false, error: "缺少必要參數" };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(ATTENDANCE_SHEET);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rawDate = data[i][0];
    const rowDate = rawDate instanceof Date
      ? Utilities.formatDate(rawDate, "Asia/Hong_Kong", "yyyy-MM-dd")
      : String(rawDate).split("T")[0];
    if (rowDate === date && data[i][3] === name) {
      sheet.getRange(i + 1, 6).setValue(remark || "");
      return { success: true, message: "已更新備註" };
    }
  }
  return { success: false, error: "找不到對應記錄" };
}

// ── 快速切換單一學生出席狀態 ──────────────────────────────────
function updateAttendanceStatus(date, grade, name, status) {
  if (!date || !name || !status) return { success: false, error: "缺少必要參數" };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(ATTENDANCE_SHEET);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rawDate = data[i][0];
    const rowDate = rawDate instanceof Date
      ? Utilities.formatDate(rawDate, "Asia/Hong_Kong", "yyyy-MM-dd")
      : String(rawDate).split("T")[0];

    if (rowDate === date && data[i][3] === name) {
      const statusCell = sheet.getRange(i + 1, 5);
      statusCell.setValue(status);
      if (status === "出席") {
        statusCell.setBackground("#d9ead3").setFontColor("#274e13");
      } else {
        statusCell.setBackground("#fce8e6").setFontColor("#a61c00");
      }
      return { success: true, message: `已將 ${name} 的狀態更新為 ${status}` };
    }
  }
  return { success: false, error: "找不到該筆記錄" };
}

// ── 取得出席統計摘要 ──────────────────────────────────────────
function getSummary(startDate, endDate, grade) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(ATTENDANCE_SHEET);
  const data = sheet.getDataRange().getValues();
  const studentStats = {};
  const dateSet = new Set();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const rawDate = row[0];
    const rowDate = rawDate instanceof Date
      ? Utilities.formatDate(rawDate, "Asia/Hong_Kong", "yyyy-MM-dd")
      : String(rawDate).split("T")[0];
    const rowGrade = row[2];
    const rowName = row[3];
    const rowStatus = row[4];

    if (startDate && rowDate < startDate) continue;
    if (endDate && rowDate > endDate) continue;
    if (grade && rowGrade !== grade) continue;

    dateSet.add(rowDate);
    const key = `${rowGrade}|${rowName}`;
    if (!studentStats[key]) {
      studentStats[key] = { grade: rowGrade, name: rowName, present: 0, absent: 0, total: 0 };
    }
    studentStats[key].total++;
    if (rowStatus === "出席") studentStats[key].present++;
    else studentStats[key].absent++;
  }

  const summary = Object.values(studentStats).map(s => ({
    ...s, rate: s.total > 0 ? Math.round((s.present / s.total) * 100) : 0
  }));
  const gradeOrder = ["F.1","F.2","F.3","F.4","F.5","F.6"];
  summary.sort((a, b) => {
    const gi = gradeOrder.indexOf(a.grade) - gradeOrder.indexOf(b.grade);
    if (gi !== 0) return gi;
    return a.name.localeCompare(b.name, "zh-HK");
  });
  return {
    success: true, summary, totalDays: dateSet.size,
    dateRange: { start: startDate || "", end: endDate || "" }
  };
}

// ── 新增學生 (包含學校欄位) ──────────────────────────────────────────────────
function addStudent(grade, name, remark, school) {
  if (!grade || !name) return { success: false, error: "年級和姓名不能為空" };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(STUDENTS_SHEET);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === name) {
      return { success: false, error: `系統內已有叫「${name}」的學生 (於 ${data[i][0]})。請輸入全名或加上姓氏以作識別。` };
    }
  }
  sheet.appendRow([grade, name, remark || "", school || ""]); // 寫入 A, B, C, D 欄
  return { success: true, message: `已新增學生：${grade} ${name}` };
}

function deleteStudent(grade, name) {
  if (!grade || !name) return { success: false, error: "年級和姓名不能為空" };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(STUDENTS_SHEET);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === grade && data[i][1] === name) {
      sheet.deleteRow(i + 1);
      return { success: true, message: `已刪除學生：${grade} ${name}` };
    }
  }
  return { success: false, error: "找不到該學生" };
}
