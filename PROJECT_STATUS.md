## 【Project Status - 2026-01-26】

> 產出時間：2026-01-26  
> 產出者：Cursor AI（自動掃描）

---

### A. 專案在做什麼（以「現在」為主）

- **核心用途：** 這是一個在瀏覽器中運行的表單編輯器，標題顯示為 "iPVMS Form"。使用者可以輸入公司名稱（程式碼中稱為 `studentId`），然後編輯兩種類型的資料：Model Data 和 Period Data。

- **使用者在畫面上可以做的事情：**
  - 輸入公司名稱（首次使用時會彈出 modal）
  - 在 Model Data 和 Period Data 之間切換
  - 在 Period Data 模式下，可以選擇或建立月份（Period Month，格式為 YYYY-MM）
  - 透過分頁標籤切換不同的工作表（sheets）
  - 在表格中輸入資料（每個儲存格是一個 `<input>` 元素）
  - 選取儲存格（單選、拖曳框選、Shift+點選範圍、Ctrl+點選多選）
  - 編輯儲存格（雙擊、F2、或直接打字）
  - 複製/剪下/貼上（Ctrl+C/X/V，使用 TSV 格式）
  - 新增列（+ Add Row 按鈕）
  - 刪除列（每列右側的 × 按鈕）
  - 上傳 .xlsx 檔案（覆寫或合併資料）
  - 下載 .xlsx 檔案（檔名包含公司名稱和時間戳）
  - 重置資料（Reset 按鈕，會顯示確認 modal）

- **它「像不像 Excel / 表單 / 教學工具 / 資料編輯器」：**
  - 觀察到的行為：表格使用 `<table>` 結構，每個儲存格是 `<input>`，支援類似 Excel 的選取、編輯、複製貼上操作
  - 有必填欄位驗證（紅色邊框、placeholder 提示）
  - 有 ChangeLog 功能（記錄所有變更）
  - 有 Undo/Redo 功能（Ctrl+Z、Ctrl+Shift+Z、Ctrl+Y）
  - 資料儲存在瀏覽器的 localStorage（依公司名稱區分）
  - 整體操作流程類似 Excel，但沒有公式、圖表等功能

---

### B. 目前的整體結構（高層）

- **是否為單頁應用：** 是。只有一個 `index.html` 檔案，所有功能都在這個頁面中完成。

- **是否有 Model / Period / Workbook / Sheet 這類概念：**
  - 有。程式碼中明確存在：
    - `ModelData` 和 `PeriodData` 兩個 workbook（工作簿）
    - 每個 workbook 下有多個 sheet（工作表）
    - `state.activeGroup` 儲存當前 workbook（"ModelData" 或 "PeriodData"）
    - `state.activeSheet` 儲存當前工作表名稱（例如 "Company"、"Exchange Rate"）
    - `state.activePeriod` 儲存當前月份（Period Data 模式使用，格式 YYYY-MM）

- **資料大致怎麼流動：**
  - 初始化：`template-data.js` 定義所有工作表的結構（headers、data、required 欄位等）→ `app.js` 的 `initFromTemplate()` 建立初始資料 → 儲存到 `state.data`
  - 載入：從 localStorage 讀取（key: `excelForm_v1_{公司名稱}`）→ 解析 JSON → 寫入 `state.data` → `renderTable()` 動態產生表格 HTML
  - 編輯：使用者輸入 → `updateCell()` 更新 `state.data` → `autoSave()` debounce 500ms → `saveToStorage()` 寫入 localStorage
  - 下載：從 `state.data` 讀取 → 使用 XLSX 函式庫組裝 Excel 檔案 → 下載
  - 上傳：讀取 .xlsx 檔案 → 解析 → 覆寫 `state.data` → 儲存 → 重新渲染

---

### C. 檔案與職責對照表（非常重要）

| 檔名 | 目前看起來負責的功能 | 分類（核心/輔助/歷史遺留） |
|------|---------------------|---------------------------|
| `index.html` | UI 骨架、載入外部資源（xlsx CDN、template-data.js、app.js）、定義所有 DOM 元素（header、toolbar、表格容器、modal 等） | 核心 |
| `css/style.css` | 全站樣式定義（顏色變數、按鈕、表格、input、選取狀態、modal、分頁等） | 核心 |
| `js/template-data.js` | 定義所有工作表的結構（TEMPLATE_DATA 物件，包含 workbooks、sheets、sheetOrder）、提供工具函式（getSheetConfig、isRequired、getExcelSheetName 等） | 核心 |
| `js/app.js` | 應用主邏輯（state 管理、初始化、事件綁定、表格渲染、選取/編輯/剪貼、Undo/Redo、驗證、上傳/下載、localStorage 操作等） | 核心 |
| `_archive/legacy_grid_system/selection_events.js` | 舊架構的選取事件處理（檔案內有 `console.log("✅ [selection_events.js] loaded")`） | 歷史遺留 |
| `_archive/legacy_grid_system/table_core.js` | 舊架構的表格核心（檔案內有 `console.log("✅ [table_core.js] loaded")`） | 歷史遺留 |
| `_archive/legacy_grid_system/table_render_core.js` | 舊架構的表格渲染核心（檔案內有 `console.log("✅ [table_render_core.js] loaded")`） | 歷史遺留 |
| `.gitattributes` | Git 設定檔 | 輔助 |
| `PROJECT_STATUS.md` | 專案狀態文件（本檔案） | 輔助 |

⚠️ **觀察到的行為：**
- `index.html` 中**沒有載入** `_archive/legacy_grid_system/` 下的任何檔案
- `_archive/legacy_grid_system/` 下的檔案內容只有 console.log，沒有實際功能程式碼（或功能已被移除）
- `app.js` 使用 `#tableHead` 和 `#tableBody` 作為表格容器，而舊檔案可能使用 `#gridHead` 和 `#gridBody`（從檔名推測）

---

### D. 已實作的功能（只列出確定存在的）

- **表格動態產生：** `renderTable()` 函式會根據 `state.data[state.activeSheet]` 動態產生 `<thead>` 和 `<tbody>`，每個儲存格是 `<input>` 元素，具有 `data-sheet`、`data-row`、`data-col` 屬性

- **選取功能：**
  - 單選：點擊儲存格會設定 `state.activeCell`
  - 拖曳框選：mousedown 後拖曳會設定 `state.selection`（矩形範圍）
  - Shift+點選：會以當前選取起點或 activeCell 為錨點，擴展選取範圍
  - Ctrl+點選：會切換 `state.multiSelection`（不連續多選）
  - 全選：Ctrl+A 會選取整個表格

- **輸入功能：**
  - 單擊：Select 模式，input 為 readonly，不會進入編輯
  - 雙擊：會進入 Edit 模式，input 移除 readonly 並 focus
  - F2：在 Select 模式下按 F2 會進入 Edit 模式
  - 直接打字：在 Select 模式下直接輸入字元會進入 Edit 模式並填入該字元
  - Enter：在 Edit 模式下會儲存並下移，在 Select 模式下僅下移
  - Tab：會右移，到最右欄則下一列第 0 欄
  - Delete/Backspace：在 Select 模式下會清空所有選取的儲存格

- **複製/剪下/貼上：**
  - 複製（Ctrl+C）：會將選取的儲存格轉換為 TSV 格式（\t 分隔欄、\n 分隔列）並寫入剪貼簿
  - 剪下（Ctrl+X）：同上，但會設定 `state.cutCells` 並顯示 `.cell-cut` 樣式
  - 貼上（Ctrl+V）：會讀取剪貼簿內容，解析 TSV，從 `activeCell` 開始貼上，必要時會自動新增列；如果上次是 cut 操作，會清空來源儲存格

- **新增/刪除列：**
  - 新增列：`#btnAddRow` 按鈕會呼叫 `addRow()`，在 `state.data[state.activeSheet].data` 中 push 一個空陣列
  - 刪除列：每列右側有 × 按鈕，會呼叫 `deleteRow()`，從 `state.data[state.activeSheet].data` 中移除該列

- **必填驗證：**
  - `validateCell()` 會檢查欄位是否必填（透過 `isRequired()` 函式）
  - 必填欄位為空時會加上 `.cell-error` class 和 placeholder 提示
  - `updateValidationSummary()` 會顯示錯誤總數

- **自動儲存：**
  - `updateCell()` 會呼叫 `autoSave()`，使用 `setTimeout` debounce 500ms 後呼叫 `saveToStorage()`
  - `saveToStorage()` 會將 `state.data`、`state.changeLog` 等寫入 localStorage（key: `excelForm_v1_{公司名稱}`）

- **上傳功能：**
  - `#inputUpload` 接受 .xlsx 檔案
  - `uploadBackup()` 會使用 XLSX 函式庫解析檔案
  - `detectWorkbook()` 會根據工作表名稱判斷是 ModelData 還是 PeriodData
  - 解析後會覆寫 `state.data` 對應的工作表資料
  - 如果有 ChangeLog 工作表，會合併到 `state.changeLog`

- **下載功能：**
  - `#btnDownload` 會呼叫 `downloadExcel()`
  - `downloadWorkbook()` 會根據 `state.activeGroup` 組裝對應的 workbook
  - 會包含所有工作表、ChangeLog、以及 Hidden 規則
  - 檔名格式：`{filename}_{公司名稱}_{時間戳}.xlsx`

- **Undo/Redo：**
  - `beginAction()`、`recordChange()`、`commitAction()` 會記錄變更到 `state.undoStack`
  - Ctrl+Z 會呼叫 `undo()`，Ctrl+Shift+Z 或 Ctrl+Y 會呼叫 `redo()`

- **Period 管理（Period Data 模式）：**
  - 可以選擇月份（`#periodSelect`）
  - 可以建立新月份（`#btnCreatePeriod`）
  - 每個月份的資料會分別儲存在 localStorage 的 `periods` 物件中

- **特殊工作表功能：**
  - `Resource Driver(Actvity Center)` 的表頭（第 3 欄起）可以編輯（使用 `.th-input`）
  - `Resource Driver(M. A. C.)` 和 `Resource Driver(S. A. C.)` 有 3 行表頭（headers、headers2、headers3）
  - `TableMapping` 工作表是唯讀的（系統對照表）

---

### E. 明顯的限制與設計原則（如果能從 code 看出）

- **避免模組化：** 所有邏輯都集中在 `app.js`（2648 行），沒有使用 ES6 modules 或任何模組系統

- **避免 framework：** 沒有使用 React、Vue、Angular 等框架，純原生 JavaScript

- **單一 HTML 檔案：** 只有一個 `index.html`，沒有路由或多頁面

- **localStorage 為唯一儲存：** 沒有後端、沒有 API、沒有資料庫，所有資料都存在瀏覽器的 localStorage

- **依賴 CDN：** XLSX 函式庫從 CDN 載入（`https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js`）

- **硬編碼的 sheet 名稱：** `app.js` 中有多處硬編碼的 sheet 名稱（例如 `DEFAULT_ROWS_MODEL_MAP`、`PERIOD_DIM_SHEETS`、`SYSTEM_EXPORT_SHEETS`），如果 `template-data.js` 中新增或更名 sheet，這些地方可能不會自動更新

- **UI 表格為單一真實來源：** 註解中明確寫道「UI table data = Source of Truth. Download rebuilds from headers + data arrays.」

---

### F. 新接手者「在改 code 前一定要知道的事」

- **不要亂動的檔案：**
  - `_archive/legacy_grid_system/` 下的檔案：雖然沒有被載入，但可能保留作歷史記錄或合規用途
  - `template-data.js` 的 `TEMPLATE_DATA` 結構：被 `app.js` 廣泛使用，改動會影響整個系統

- **核心假設：**
  - `state.data` 的結構：`state.data[sheetName] = { headers: [], data: [] }`
  - 表格的 DOM 結構：`#tableHead` 和 `#tableBody`，每個儲存格是 `<input>` 並有 `data-sheet`、`data-row`、`data-col` 屬性
  - localStorage 的 key 格式：`excelForm_v1_{公司名稱}`
  - Period Data 的資料結構：`periods[YYYY-MM] = { data: {}, changeLog: [], activeSheet: "" }`

- **改錯會整個壞掉的地方：**
  - `state` 物件的結構：被多處讀寫，增刪屬性需要全局搜尋
  - `getInputAt()` 函式依賴 `data-sheet`、`data-row`、`data-col` 屬性，如果改動表格 DOM 結構會失效
  - `renderTable()` 會清空並重建整個表格，如果改動 DOM 結構或事件綁定方式，選取/編輯功能可能失效
  - `updateCell()` 是資料更新的唯一入口，如果改動會影響自動儲存、ChangeLog、Undo/Redo
  - `template-data.js` 的 `getSheetConfig()`、`isRequired()` 等函式被廣泛使用，改動會影響驗證、上傳/下載

- **script 載入順序很重要：** `index.html` 中的 script 順序必須是：xlsx CDN → template-data.js → app.js

---

### G. 不確定或需要向原作者確認的問題（列清單）

⚠️ **這一段非常重要**  
以下問題無法從 code 確定，需要向原作者確認：

1. **`_archive/legacy_grid_system/` 下的檔案是否可以刪除？** 還是必須保留作合規/稽核用途？

2. **`Resource Driver(Actvity Center)` 表頭編輯不進 ChangeLog/Undo 是刻意的嗎？** 還是之後要納入？

3. **`DEFAULT_ROWS_MODEL_MAP`、`PERIOD_DIM_SHEETS`、`SYSTEM_EXPORT_SHEETS`、`nav-pill-muted` 的 sheet 名稱是寫死的，** 如果之後在 `template-data.js` 中新增或更名 sheet，是否有流程要同步更新 `app.js`，還是現階段手動即可？

4. **Download 功能是否只下載當前 workbook？** 還是之後需要支援一次下載 Model+Period 兩本？

5. **`app.js` 中有多處 `console.log`（例如 `[headers]`、`[keydown] Ctrl+X`、`[doPaste] error`），** 這些是除錯用的嗎？是否需要移除或管制？

6. **`#studentModal` 在 HTML 中沒有 `hidden` class，** 首屏可能會短暫顯示再隱藏，這是刻意的嗎？還是需要加上初始 `hidden` class？

7. **Period Data 模式下，如果切換月份時沒有選擇月份，** 系統會如何處理？從 code 看會初始化，但這是否符合預期？

8. **`TableMapping` 工作表是系統對照表，** 是否永遠不允許使用者修改？還是之後會開放部分欄位？

9. **`Resource Driver(Actvity Center)` 可以動態新增/刪除欄位（Driver Code），** 這個功能是否也適用於其他工作表？還是只有這個工作表有這個功能？

10. **`migrateMACHeaderNames()`、`migrateSACHeaderNames()` 等遷移函式，** 這些是為了向下相容舊資料嗎？之後是否還需要這些函式？
