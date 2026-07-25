# 補習社出席系統 (Pentex Attendance) - AI Agent 交接文件

這份文件旨在幫助新的 AI Agent 快速了解專案架構、現有功能、Google Sheets 結構、部署流程及目前待修復的問題。請在進行任何修改前，仔細閱讀此文件。

---

## 1. 系統架構與重要連結

本系統採用 **Google Apps Script (GAS)** 作為後端與前端的託管平台，並以 **Google Sheets** 作為資料庫。前端亦可透過 GitHub Pages 靜態部署（PWA 支援）。

*   **Google Apps Script 編輯器**：`https://script.google.com/home/projects/1iURnxVOQmiPm-6w9k0MCSjaNsBDpn732JiTCYRiexyBPLLXYGeKPn3LQ/edit`
*   **Web App 運行網址**：`https://script.google.com/macros/s/AKfycbwf0fRAhoUtLV-DhRw4Iyl8LlM8TFInzWKmVLft1hkCn5EMjEFP_R2WAaADOXAaPWM/exec`
*   **Google Sheets 資料庫**：ID 為 `1ibit_1LtSGfhaI2NuBIIJwL5fhYIVnM-gRruF44pLak`
*   **GitHub 儲存庫**：`https://github.com/jameschoyhk-crypto/pentex-attendance`
*   **GitHub Pages**：`https://jameschoyhk-crypto.github.io/pentex-attendance/`

---

## 2. 檔案結構與說明

專案主要由兩個檔案構成（位於 `attendance_system/` 目錄，需同步至 Apps Script 編輯器）：

### `Code.gs` (後端邏輯)
處理所有與 Google Sheets 互動的邏輯。
*   `doGet(e)`：處理 GET 請求，回傳 HTML 介面或讀取資料（如 `getStudents`, `getAttendance`, `getSummary`, `getStudentAttendance`, `getAttendanceRange`）。
*   `doPost(e)`：處理 POST 請求，寫入資料（如 `saveAttendance`, `addStudent`, `deleteStudent`, `updateStudent`, `updateStudentRemark`, `updateAttendanceRemark`）。
*   **注意**：Google Apps Script 有執行時間限制（通常為 6 分鐘），處理大量數據時需注意效能。

### `index.html` (前端介面)
單頁應用程式 (SPA)，包含 HTML、CSS (內聯) 和 JavaScript (內聯)。
*   **導航機制**：使用 `switchTab(name, btn)` 切換不同的 `page-<name>` 區塊。
*   **API 呼叫**：透過 `apiGet()` 和 `apiPost()` 與 `Code.gs` 通訊。
*   **頁面 ID 對應**：
    *   `page-attendance`：簽到頁面
    *   `page-history`：查閱頁面（所有學生出席記錄）
    *   `page-student`：學生個人查詢頁面
    *   `page-summary`：統計頁面（出席率）
    *   `page-manage`：管理頁面（新增/修改/刪除學生）

---

## 3. Google Sheets 資料庫結構

試算表 ID：`1ibit_1LtSGfhaI2NuBIIJwL5fhYIVnM-gRruF44pLak`

包含兩個主要工作表 (Sheets)：

1.  **學生名單 (`STUDENTS_SHEET`)**
    *   Column A: 年級 (Grade)
    *   Column B: 姓名 (Name)
    *   Column C: 備註 (Remark)
2.  **出席記錄 (`ATTENDANCE_SHEET`)**
    *   Column A: 日期 (Date, 格式為 yyyy-MM-dd)
    *   Column B: 星期 (Weekday)
    *   Column C: 年級 (Grade)
    *   Column D: 姓名 (Name)
    *   Column E: 狀態 (Status, "出席" 或 "缺席")
    *   Column F: 備註 (Remark)
    *   Column G: 時間戳記 (Timestamp)

---

## 4. 已知 Bug 與待修復清單

在最新的自動化測試中，發現了 5 個前端 UI 互動問題。**下一個接手的 Agent 應優先修復這些問題**：

1.  **簽到頁面 (`page-attendance`)**：點擊單一學生的「✅ 出席」或「❌ 缺席」按鈕時，可能會觸發頁面重繪或狀態重置，導致頂部的「年級」(`att-grade`) 選擇丟失。當點擊「儲存記錄」時，系統會提示「請選擇日期和年級」而無法儲存。
2.  **查閱頁面 (`page-history`)**：「年級」下拉選單 (`hist-grade`) 的篩選功能失效，即使選擇了特定年級，仍然會顯示所有年級的記錄。
3.  **查閱頁面 (`page-history`)**：點擊「✏️」編輯備注時，備注模態框 (`remarkModal`) 未能正確彈出顯示（可能是 CSS 的 z-index 問題或 JS 邏輯錯誤）。
4.  **管理頁面 (`page-manage`)**：學生列表的「年級篩選」(`mgr-filter-grade`) 功能失效，無法正確隱藏其他年級的學生。
5.  **管理頁面 (`page-manage`)**：點擊「🗑 刪除」按鈕時，確認對話框未能顯示，導致無法執行刪除動作。

---

## 5. 開發與部署流程

當你修改了 `Code.gs` 或 `index.html` 後，請遵循以下步驟進行部署：

1.  **同步本地檔案**：確保本地的 `attendance_system/` 目錄下的檔案已更新。
2.  **更新 Apps Script 編輯器**：
    *   使用瀏覽器工具開啟 Apps Script 編輯器。
    *   將修改後的 `Code.gs` 或 `index.html` 內容全選並貼上覆蓋。
    *   點擊「儲存專案」（或按 Ctrl+S），並**務必等待雲端硬碟圖示顯示打勾 (`cloud_done`)**，確認儲存完成（檔案較大時可能需要數十秒）。
3.  **部署新版本**：
    *   點擊右上角「部署」 > 「管理部署作業」。
    *   點擊「編輯」按鈕（鉛筆圖示）。
    *   在「版本」下拉選單中選擇「新版本」。
    *   點擊「部署」。
4.  **同步至 GitHub**（可選，用於版本控制和 GitHub Pages）：
    *   將變更 commit 並 push 到 GitHub 儲存庫。

---
**文件撰寫者：** Manus AI (2026-07-24)
