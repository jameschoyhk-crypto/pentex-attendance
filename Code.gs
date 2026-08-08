// ============================================================
// 補習社出席記錄系統 — Google Apps Script 後端
// 版本: 3.0 (新增學生個人查詢、備註編輯、PWA)
// ============================================================

const SPREADSHEET_ID = "1ibit_1LtSGfhaI2NuBIIJwL5fhYIVnM-gRruF44pLak";
const STUDENTS_SHEET = "學生名單";
const ATTENDANCE_SHEET = "出席記錄";
const SCHOOLS_SHEET = "學校名單";

// ── CORS 輔助 ─────────────────────────────────────────────────
function corsResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 主入口：GET 請求 ──────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action || "";

  // 無 action 參數時返回 HTML 介面
  if (!action) {
    return HtmlService.createHtmlOutputFromFile("index")
      .setTitle("補習社出席系統")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  try {
    if (action === "getStudents") {
      return corsResponse(getStudents(e.parameter.grade));
    }
    if (action === "getAttendance") {
      return corsResponse(getAttendance(e.parameter.date, e.parameter.grade));
    }
    if (action === "getSummary") {
      return corsResponse(getSummary(e.parameter.startDate, e.parameter.endDate, e.parameter.grade));
    }
    if (action === "getGrades") {
      return corsResponse(getGrades());
    }
    if (action === "getTodayStatus") {
      return corsResponse(getTodayStatus(e.parameter.grade));
    }
    // 新增：查詢某學生某時段出席記錄
    if (action === "getStudentAttendance") {
      return corsResponse(getStudentAttendance(
        e.parameter.grade,
        e.parameter.name,
        e.parameter.startDate,
        e.parameter.endDate
      ));
    }
    // 新增：取得某學生備註
    if (action === "getStudentRemark") {
      return corsResponse(getStudentRemark(e.parameter.grade, e.parameter.name));
    }
    // 新增：查詢日期範圍內的出席記錄
    if (action === "getAttendanceRange") {
      return corsResponse(getAttendanceRange(
        e.parameter.startDate,
        e.parameter.endDate,
        e.parameter.grade,
        e.parameter.school
      ));
    }
    if (action === "getSchools") {
      return corsResponse(getSchools());
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
      return corsResponse(addStudent(payload.grade, payload.name, payload.remark));
    }
    if (action === "deleteStudent") {
      return corsResponse(deleteStudent(payload.grade, payload.name));
    }
    // 新增：修改學生資料（年級、姓名）
    if (action === "updateStudent") {
      return corsResponse(updateStudent(
        payload.oldGrade, payload.oldName,
        payload.newGrade, payload.newName,
        payload.newRemark
      ));
    }
    // 新增：更新學生備註
    if (action === "updateStudentRemark") {
      return corsResponse(updateStudentRemark(payload.grade, payload.name, payload.remark));
    }
    // 新增：更新某條出席記錄的備註
    if (action === "updateAttendanceRemark") {
      return corsResponse(updateAttendanceRemark(payload.date, payload.grade, payload.name, payload.remark));
    }
    // 新增：修改出席記錄（日期、狀態、備註）
    if (action === "updateAttendanceRecord") {
      return corsResponse(updateAttendanceRecord(
        payload.oldDate, payload.grade, payload.name,
        payload.newDate, payload.newStatus, payload.newRemark
      ));
    }
    if (action === "addSchool") {
      return corsResponse(addSchool(payload.name));
    }
    if (action === "deleteSchool") {
      return corsResponse(deleteSchool(payload.name));
    }
    if (action === "updateStudentSchool") {
      return corsResponse(updateStudentSchool(payload.grade, payload.name, payload.school));
    }
    return corsResponse({ success: false, error: "Unknown action: " + action });
  } catch (err) {
    return corsResponse({ success: false, error: err.message });
  }
}

// ── 取得所有年級 ──────────────────────────────────────────────
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

// ── 取得學生名單（可按年級篩選）────────────────────────────────
function getStudents(grade) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(STUDENTS_SHEET);
  const data = sheet.getDataRange().getValues();

  const students = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && !row[1]) continue;
    if (grade && row[0] !== grade) continue;
        students.push({ grade: row[0], name: row[1], remark: row[2] || "", school: row[3] || "" });
  }
  return { success: true, students };
}

// ── 儲存出席記錄 ──────────────────────────────────────────────
function saveAttendance(date, grade, records) {
  if (!date || !grade || !records || records.length === 0) {
    return { success: false, error: "缺少必要參數" };
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(ATTENDANCE_SHEET);
  const now = new Date();
  const timestamp = Utilities.formatDate(now, "Asia/Hong_Kong", "yyyy-MM-dd HH:mm:ss");

  // 先刪除同一日期同一年級的舊記錄
  const allData = sheet.getDataRange().getValues();
  const rowsToDelete = [];
  for (let i = allData.length - 1; i >= 1; i--) {
    if (allData[i][0] === date && allData[i][2] === grade) {
      rowsToDelete.push(i + 1);
    }
  }
  rowsToDelete.forEach(r => sheet.deleteRow(r));

  // 計算星期
  const dateObj = new Date(date);
  const weekdays = ["日", "一", "二", "三", "四", "五", "六"];
  const weekday = "星期" + weekdays[dateObj.getDay()];

  // 寫入新記錄
  const newRows = records.map(r => [
    date, weekday, grade, r.name, r.status, r.remark || "", timestamp
  ]);

  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 7).setValues(newRows);

    const lastRow = sheet.getLastRow();
    const startRow = lastRow - newRows.length + 1;
    for (let i = 0; i < newRows.length; i++) {
      const statusCell = sheet.getRange(startRow + i, 5);
      if (newRows[i][4] === "出席") {
        statusCell.setBackground("#d9ead3");
        statusCell.setFontColor("#274e13");
      } else {
        statusCell.setBackground("#fce8e6");
        statusCell.setFontColor("#a61c00");
      }
    }
  }

  const presentCount = records.filter(r => r.status === "出席").length;
  const absentCount = records.filter(r => r.status === "缺席").length;

  return {
    success: true,
    message: `已儲存 ${records.length} 條記錄`,
    summary: { present: presentCount, absent: absentCount, total: records.length }
  };
}

// ── 取得指定日期（及年級）的出席記錄 ─────────────────────────
function getAttendance(date, grade) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(ATTENDANCE_SHEET);
  const data = sheet.getDataRange().getValues();

  const records = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    if (date && row[0] !== date) continue;
    if (grade && row[2] !== grade) continue;
    records.push({
      date: row[0], weekday: row[1], grade: row[2],
      name: row[3], status: row[4], remark: row[5] || "", timestamp: row[6] || ""
    });
  }
  return { success: true, records };
}

// ── 取得今日某年級已記錄狀態 ──────────────────────────────────
function getTodayStatus(grade) {
  const today = Utilities.formatDate(new Date(), "Asia/Hong_Kong", "yyyy-MM-dd");
  return getAttendance(today, grade);
}

// ── 新增：查詢某學生某時段出席記錄 ───────────────────────────
function getStudentAttendance(grade, name, startDate, endDate) {
  if (!grade || !name) return { success: false, error: "請提供年級和姓名" };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(ATTENDANCE_SHEET);
  const data = sheet.getDataRange().getValues();

  const records = [];
  let presentCount = 0;
  let absentCount = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    // 將 Date 物件或其他格式統一轉換為 yyyy-MM-dd 字串
    let rowDate;
    if (row[0] instanceof Date) {
      rowDate = Utilities.formatDate(row[0], "Asia/Hong_Kong", "yyyy-MM-dd");
    } else {
      rowDate = String(row[0]);
    }
    if (row[2] !== grade || row[3] !== name) continue;
    if (startDate && rowDate < startDate) continue;
    if (endDate && rowDate > endDate) continue;

    const rec = {
      date: rowDate, weekday: row[1], status: row[4], remark: row[5] || ""
    };
    records.push(rec);
    if (row[4] === "出席") presentCount++;
    else absentCount++;
  }

  // 按日期排序（最新在前）
  records.sort((a, b) => b.date.localeCompare(a.date));

  const total = records.length;
  const rate = total > 0 ? Math.round((presentCount / total) * 100) : 0;

  return {
    success: true,
    student: { grade, name },
    records,
    summary: { present: presentCount, absent: absentCount, total, rate },
    dateRange: { start: startDate || "", end: endDate || "" }
  };
}

// ── 新增：取得某學生備註 ──────────────────────────────────────
function getStudentRemark(grade, name) {
  if (!grade || !name) return { success: false, error: "請提供年級和姓名" };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(STUDENTS_SHEET);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === grade && data[i][1] === name) {
      return { success: true, grade, name, remark: data[i][2] || "" };
    }
  }
  return { success: false, error: "找不到該學生" };
}

// ── 新增：修改學生資料（年級、姓名）────────────────────────────
function updateStudent(oldGrade, oldName, newGrade, newName, newRemark) {
  if (!oldGrade || !oldName) return { success: false, error: "請提供原年級和姓名" };
  if (!newGrade || !newName) return { success: false, error: "新年級和姓名不能為空" };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(STUDENTS_SHEET);
  const data = sheet.getDataRange().getValues();

  // 如果改了姓名或年級，先檢查新資料是否已存在（排除自身）
  if (newGrade !== oldGrade || newName !== oldName) {
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === newGrade && data[i][1] === newName) {
        return { success: false, error: `「${newGrade} ${newName}」已存在` };
      }
    }
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === oldGrade && data[i][1] === oldName) {
      sheet.getRange(i + 1, 1).setValue(newGrade);
      sheet.getRange(i + 1, 2).setValue(newName);
      if (newRemark !== undefined) sheet.getRange(i + 1, 3).setValue(newRemark);
      return { success: true, message: `已更新：${oldGrade} ${oldName} → ${newGrade} ${newName}` };
    }
  }
  return { success: false, error: "找不到該學生" };
}

// ── 新增：更新學生備註（學生名單）────────────────────────────
function updateStudentRemark(grade, name, remark) {
  if (!grade || !name) return { success: false, error: "請提供年級和姓名" };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(STUDENTS_SHEET);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === grade && data[i][1] === name) {
      sheet.getRange(i + 1, 3).setValue(remark || "");
      return { success: true, message: `已更新 ${grade} ${name} 的備註` };
    }
  }
  return { success: false, error: "找不到該學生" };
}

// ── 新增：更新某條出席記錄的備註 ─────────────────────────────
function updateAttendanceRemark(date, grade, name, remark) {
  if (!date || !grade || !name) return { success: false, error: "缺少必要參數" };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(ATTENDANCE_SHEET);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === date && data[i][2] === grade && data[i][3] === name) {
      sheet.getRange(i + 1, 6).setValue(remark || "");
      return { success: true, message: "已更新備註" };
    }
  }
  return { success: false, error: "找不到對應記錄" };
}

// ── 新增：查詢日期範圍內的出席記錄 ────────────────────────────────
function getAttendanceRange(startDate, endDate, grade, school) {
  if (!startDate || !endDate) return { success: false, error: "請提供開始和結束日期" };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(ATTENDANCE_SHEET);
  const data = sheet.getDataRange().getValues();

  // 建立學生學校映射表
  const studSheet = ss.getSheetByName(STUDENTS_SHEET);
  const studData = studSheet.getDataRange().getValues();
  const schoolMap = {};
  for (let i = 1; i < studData.length; i++) {
    if (studData[i][0] && studData[i][1]) {
      schoolMap[studData[i][0] + '|' + studData[i][1]] = studData[i][3] || '';
    }
  }

  const records = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    let rowDate;
    if (row[0] instanceof Date) {
      rowDate = Utilities.formatDate(row[0], "Asia/Hong_Kong", "yyyy-MM-dd");
    } else {
      rowDate = String(row[0]);
    }
    if (rowDate < startDate || rowDate > endDate) continue;
    if (grade && row[2] !== grade) continue;
    const studentSchool = schoolMap[row[2] + '|' + row[3]] || '';
    if (school && studentSchool !== school) continue;
    records.push({
      date: rowDate, weekday: row[1], grade: row[2],
      name: row[3], status: row[4], remark: row[5] || "", timestamp: row[6] || "",
      school: studentSchool
    });
  }

  records.sort((a, b) => b.date.localeCompare(a.date) || a.grade.localeCompare(b.grade) || a.name.localeCompare(b.name, "zh-HK"));

  return { success: true, records, dateRange: { start: startDate, end: endDate } };
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
    const rowDate = row[0];
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
    ...s,
    rate: s.total > 0 ? Math.round((s.present / s.total) * 100) : 0
  }));

  const gradeOrder = ["F.1","F.2","F.3","F.4","F.5","F.6"];
  summary.sort((a, b) => {
    const gi = gradeOrder.indexOf(a.grade) - gradeOrder.indexOf(b.grade);
    if (gi !== 0) return gi;
    return a.name.localeCompare(b.name, "zh-HK");
  });

  return {
    success: true, summary,
    totalDays: dateSet.size,
    dateRange: { start: startDate || "", end: endDate || "" }
  };
}

// ── 新增學生 ──────────────────────────────────────────────────
function addStudent(grade, name, remark) {
  if (!grade || !name) return { success: false, error: "年級和姓名不能為空" };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(STUDENTS_SHEET);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === grade && data[i][1] === name) {
      return { success: false, error: "該學生已存在" };
    }
  }

  sheet.appendRow([grade, name, remark || ""]);
  return { success: true, message: `已新增學生：${grade} ${name}` };
}

// ── 刪除學生 ──────────────────────────────────────────────────
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

// ── 修改出席記錄（日期、狀態、備註）────────────────────────────
function updateAttendanceRecord(oldDate, grade, name, newDate, newStatus, newRemark) {
  if (!oldDate || !grade || !name) return { success: false, error: "缺少必要參數" };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(ATTENDANCE_SHEET);
  const data = sheet.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();

  for (let i = 1; i < data.length; i++) {
    const rawDate = data[i][0];
    const rowDate = rawDate instanceof Date
      ? Utilities.formatDate(rawDate, tz, 'yyyy-MM-dd')
      : String(rawDate).split('T')[0];
    const rowGrade = data[i][2];
    const rowName  = data[i][3];

    if (rowDate === oldDate && rowGrade === grade && rowName === name) {
      const targetDate = newDate || oldDate;
      const targetDateObj = new Date(targetDate + 'T00:00:00');
      const weekdays = ['星期日','星期一','星期二','星期三','星期四','星期五','星期六'];
      const weekday = weekdays[targetDateObj.getDay()];

      sheet.getRange(i + 1, 1).setValue(targetDate);
      sheet.getRange(i + 1, 2).setValue(weekday);
      if (newStatus) sheet.getRange(i + 1, 5).setValue(newStatus);
      if (newRemark !== undefined) sheet.getRange(i + 1, 6).setValue(newRemark);

      return { success: true, message: `已更新：${grade} ${name} 的記錄` };
    }
  }
  return { success: false, error: "找不到對應的出席記錄" };
}

// ── 取得學校名單 ──────────────────────────────────────────────
function getSchools() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SCHOOLS_SHEET);
  if (!sheet) return { success: true, schools: [] };
  const data = sheet.getDataRange().getValues();
  const schools = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) schools.push(data[i][0]);
  }
  return { success: true, schools };
}

// ── 新增學校 ──────────────────────────────────────────────────
function addSchool(name) {
  if (!name) return { success: false, error: "學校名稱不能為空" };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SCHOOLS_SHEET);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === name) return { success: false, error: "該學校已存在" };
  }
  sheet.appendRow([name]);
  return { success: true, message: `已新增學校：${name}` };
}

// ── 刪除學校 ──────────────────────────────────────────────────
function deleteSchool(name) {
  if (!name) return { success: false, error: "學校名稱不能為空" };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SCHOOLS_SHEET);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === name) {
      sheet.deleteRow(i + 1);
      return { success: true, message: `已刪除學校：${name}` };
    }
  }
  return { success: false, error: "找不到該學校" };
}

// ── 更新學生所屬學校 ──────────────────────────────────────────
function updateStudentSchool(grade, name, school) {
  if (!grade || !name) return { success: false, error: "請提供年級和姓名" };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(STUDENTS_SHEET);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === grade && data[i][1] === name) {
      sheet.getRange(i + 1, 4).setValue(school || "");
      return { success: true, message: `已更新 ${grade} ${name} 的學校為「${school || "（無）"}」` };
    }
  }
  return { success: false, error: "找不到該學生" };
}
