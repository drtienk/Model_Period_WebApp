# PROJECT_STATUS.md

## 【Project Status - 2026-02-07】

本文件為 **完整交接文件**，供從未接觸本專案的新開發者或 AI Agent 快速理解系統、架構與修改入口。所有功能描述均包含：目的、實作位置、相關檔案、安全修改方式。

---

## 1. Project Overview

### Business Purpose
- **產品名稱：** iPVMS Form（瀏覽器端表單/試算表編輯器）
- **業務目的：** 為公司/模型與期間（月份）提供結構化資料輸入；透過 Excel（.xlsx）匯入/匯出，支援 Model Data（基本設定）與 Period Data（月度資料）兩大工作區。

### Primary Users
- 需依公司名稱區分資料的業務或教學使用者
- 需依 YYYY-MM 管理多期別資料的使用者

### Core Workflows
1. **進入系統：** 首次使用彈出公司名稱輸入（modal）→ 公司名稱存入 sessionStorage，資料以公司為範圍存於 localStorage。
2. **切換工作區：** Model Data / Period Data 切換；Period Data 下可選擇或建立月份（YYYY-MM）。
3. **編輯：** 分組分頁切換工作表 → 在 input 型網格中編輯儲存格；支援點選、拖曳、Shift、Ctrl 多選；複製/剪下/貼上（TSV）；新增/刪除列。
4. **上傳/下載：** 上傳 .xlsx（覆蓋/合併現有資料）、下載 .xlsx（檔名含公司名 + 時間戳）、重置為模板。

### High-Level Capability Summary
- 配置驅動的分頁（tabs）：由 template-data 定義 workbooks、uiGroups、sheetOrder。
- 特定工作表支援動態欄位：Resource Driver(Actvity Center)、Resource Driver(M. A. C.)、Resource Driver(S. A. C.) — 可新增/刪除/重新命名「使用者新增欄位」。
- Undo/Redo、ChangeLog、必填欄位驗證、TableMapping 唯讀系統表。
- Period Data 可選「Optional Tabs」與「Required Fields」管理（Optional Tabs 控制分頁是否顯示為 dim；Required Fields 可覆寫必填設定）。

---

## 2. Technology Stack

| 類別 | 技術 |
|------|------|
| **Frontend** | 單一 HTML 頁面，純 Vanilla JS（無框架） |
| **Storage** | `localStorage` 鍵：`excelForm_v1_{companyName}`；`sessionStorage` 存公司名稱 |
| **Auth** | 無；僅以公司名稱識別（sessionStorage + UI 顯示） |
| **Cloud** | 無；僅本地 |
| **Import/Export** | XLSX（CDN：cdnjs xlsx 0.18.5） |
| **Styles** | `css/style.css` |

---

## 3. Application Architecture Overview

### 系統分層與職責

| 層級 | 職責 | 主要檔案 / 位置 |
|------|------|------------------|
| **UI Rendering Layer** | DOM 結構（header、toolbar、workbook 切換、period 選擇器、分頁容器、`#tableHead`/`#tableBody`、modals、狀態列）與樣式 | `index.html`, `css/style.css` |
| **Configuration Layer** | 單一真相來源：workbooks、sheets、sheetOrder、TABLE_MAPPING、必填/選填、Excel 工作表名稱對應 | `js/template-data.js` |
| **State / Cache Layer** | 單一 in-memory 狀態：studentId、activeGroup、activeSheet、activePeriod、data、changeLog、選取、剪貼、Undo/Redo | `js/app.js` 的 `state` 物件 |
| **Table/Grid Engine** | 表格渲染、選取、編輯、複製/剪下/貼上、新增刪除列、動態欄位 UI、驗證 | `js/app.js`（renderTable、選取與剪貼相關函數） |
| **Tab Management System** | 依 workbook 與 uiGroups 渲染分組分頁、dim 樣式、切換 sheet | `js/app.js`：renderWorkbookToggle、renderGroupedNav |
| **Upload / Download Engine** | Excel 解析與寫入、欄位對齊、自動擴欄、覆蓋/合併、檔名與時間戳 | `js/app.js`：uploadBackup、downloadExcel、downloadWorkbook、detectWorkbook |
| **Data Persistence Layer** | 依 company 與 period 讀寫 localStorage、自動儲存（debounce 500ms） | `js/app.js`：saveToStorage、autoSave、loadStudentData |
| **Authentication / Presence** | 無；無雲端、無即時在線狀態 | — |

### 資料流概要
- **載入：** `loadStudentData(studentId)` 讀取 localStorage → 依 activeGroup 載入 root `data`（Model）或 `periods[activePeriod]`（Period）→ `ensureAllSheets()` 補齊 sheet、TableMapping 強制覆寫 → `renderAll()`。
- **儲存：** 儲存格或欄位變更 → `updateCell` / 欄位操作 → `autoSave()` → `saveToStorage()` 寫入對應 key（Model 寫根層；Period 寫 `periods[activePeriod]`）。
- **上傳：** FileReader → XLSX.read → `detectWorkbook` 辨識 Model/Period → 逐 sheet 解析、欄位對齊與擴欄、寫入 `state.data` → TableMapping 還原、migrations → `saveToStorage()`、`renderAll()`。
- **下載：** 依當前 workbook 從 `state.data` 與 template 組出所有 sheet + ChangeLog → XLSX.writeFile（檔名含公司 + 時間戳）。

---

## 4. Folder / File Responsibility Map

### 根目錄
- **index.html** — 單頁骨架：header、toolbar、workbook 按鈕、period 選擇器、分頁容器、表格、modals（公司名稱、Reset、Required Fields、Optional Tabs）、狀態列。腳本載入順序：XLSX → template-data.js → app.js。
- **PROJECT_STATUS.md** — 本交接文件。

### css/
- **style.css** — 全站樣式：layout、表格、按鈕、modal、驗證、分頁、dim 等。

### js/
- **template-data.js**  
  - **責任：** 工作表結構與對照表之單一真相來源；sheet 名稱須與原始 Excel 完全一致（含空格、拼字）。  
  - **主要內容：** `TABLE_MAPPING_HEADERS` / `TABLE_MAPPING_DATA`、`TEMPLATE_DATA`（workbooks、sheets、sheetOrder）、`getSheetsForWorkbook`、`getSheetConfig`、`norm`、`getRequiredOverride` / `setRequiredOverride`、`isRequired`、`getExcelSheetName`、`getInternalSheetName`。  
  - **觸發時機：** 載入時即執行；app.js 在 init、載入、渲染、上傳、下載時呼叫上述函數。  
  - **依賴：** 所有 sheet 列表、欄位、必填、Excel 名稱對應皆由此取得；修改 sheet 名稱或結構時須同步 app.js 內常數。  
  - **修改風險：** 中高；更動 sheet 名稱或 workbook 會影響 app.js 多處常數與條件。

- **app.js**  
  - **責任：** 應用程式全部邏輯：初始化、載入/儲存、渲染（分頁 + 表格）、選取、編輯、複製/剪下/貼上、新增刪除列、動態欄位（新增/刪除/重新命名）、Undo/Redo、驗證、上傳/下載、期間切換、migrations、Required/Optional UI。  
  - **主要函數（節錄）：**  
    - Init：`init`、`checkStudentId`、`setStudentId`、`loadStudentData`、`initFromTemplate`、`ensureAllSheets`。  
    - Storage：`saveToStorage`、`autoSave`。  
    - Cell：`updateCell`、`updateValidation`、`validateCell`、`updateValidationSummary`。  
    - Undo/Redo：`beginAction`、`recordChange`、`recordChangeSpecial`、`commitAction`、`undo`、`redo`、`clearRedo`。  
    - Selection：`getInputAt`、`setActiveCell`、`getSelectedCells`、`getSelectedCellInputs`、`updateSelectionUI`、`rectFromTwoPoints`、`getCellFromPoint`。  
    - Clipboard：`doCopy`、`doCut`、`doPaste`、`buildClipboardTextFromCells`、`writeClipboardText`、`readClipboardText`。  
    - Table：`renderTable`、`renderAll`、`renderWorkbookToggle`、`renderGroupedNav`、`addRow`、`deleteRow`。  
    - Dynamic columns：`addDriverCodeColumn`、`deleteDriverCodeColumn`、`addDriverCodeColumnMAC`/`deleteDriverCodeColumnMAC`、`addDriverCodeColumnSAC`/`deleteDriverCodeColumnSAC`、`insertColumnRight`、`insertColumnAt`、`removeColumn`、`deleteUserColumn`、`renameColumn`、`makeUniqueHeaderName`、`ensureUserAddedColIds`、`isUserAddedColumn`。  
    - Download/Upload：`downloadExcel`、`downloadWorkbook`、`uploadBackup`、`detectWorkbook`。  
    - Period：`getAllPeriodsFromStorage`、`getCurrentPeriodData`、`switchPeriod`、`updatePeriodSelector`、`createNewPeriod`、`deleteCurrentPeriod`。  
    - Required/Optional：`showRequiredFieldsModal`、`saveRequiredFieldsOverride`、`getPeriodTabDimPrefs`、`setPeriodTabDimPrefs`、`showOptionalTabsModal`、`saveOptionalTabsFromUI`。  
  - **觸發時機：** 頁面載入後 `init()` → `checkStudentId()`、`bindEvents()`；其餘由使用者操作與定時 autoSave 觸發。  
  - **修改風險：** 高；單檔約 3.6k+ 行，選取與表格 DOM 契約（data-sheet/data-row/data-col）被多處依賴。

### _archive/
- **_archive/legacy_grid_system/** — 舊版網格實作，**未被載入**；僅供參考或合規保留。

---

## 5. Tab & Sheet Configuration System

### 定義位置
- **檔案：** `js/template-data.js`
- **結構：** `TEMPLATE_DATA.workbooks` 有兩個 workbook：
  - **ModelData** — name、description、color、filename、`uiGroups`（key、label、sheets 陣列）。
  - **PeriodData** — 同上。
- **分頁順序：** `TEMPLATE_DATA.sheetOrder.ModelData` 與 `sheetOrder.PeriodData` 陣列決定每個 workbook 的 sheet 顯示順序。

### 如何新增/移除/調整分頁
- **新增 sheet：** 在 `template-data.js` 的 `TEMPLATE_DATA.sheets` 新增一筆（workbook、headers、data、必要時 required、hidden、sheetNameInExcel 等），並在對應的 `sheetOrder[workbookKey]` 陣列中加入該 sheet 的 key。
- **移除 sheet：** 從 `sheets` 與 `sheetOrder` 中移除；若為 Model 且有用到預設行數，須同步修改 app.js 的 `DEFAULT_ROWS_MODEL_MAP`。
- **修改分頁標籤或分組：** 修改 `workbooks[].uiGroups` 的 label 或 sheets 陣列；分頁由 `getSheetsForWorkbook()` 取得順序，再依 uiGroups 分組顯示（實際顯示仍以 sheetOrder 為準，group 主要影響視覺分組）。

### Required / Optional 樣式
- **必填：** `template-data.js` 各 sheet 的 `required` 陣列；另可透過「Required Fields」管理員 modal 覆寫，存於 `localStorage` 鍵 `period_required_override_v1`。`isRequired(sheetName, columnName)` 會先查覆寫再查 template。
- **Optional Tabs（dim）：** Period Data 下，哪些分頁顯示為「選用/dim」由 `getPeriodTabDimPrefs(studentId)` 決定，預設來自 app.js 的 `PERIOD_DIM_SHEETS` Set；使用者可透過「Optional Tabs」modal 修改，存於 `excelForm_v1_{studentId}.uiPrefs.periodTabDim`。`renderGroupedNav()` 會對在 dim set 中的 sheet 加上 `nav-pill-muted` class。

### 重要常數（app.js）
- **DEFAULT_ROWS_MODEL_MAP** — Model Data 各 sheet 的預設空白行數（若 template 未給 data 時使用）。
- **PERIOD_DIM_SHEETS** — 預設視為 Optional（dim）的 Period 工作表名稱 Set。
- **SYSTEM_EXPORT_SHEETS** — Model 下載時使用 template 內容的系統 sheet 名稱 Set。
- 隱藏/系統 sheet 在 `renderGroupedNav` 中會依 `config.hidden` 等條件處理；dim 分頁為 `nav-pill-muted`。

---

## 6. Table / Grid Engine Behavior

### Cell 渲染
- **位置：** `app.js` 的 `renderTable()`。
- **邏輯：** 依 `state.activeSheet` 從 `state.data[sheetName]` 取 `headers` 與 `data`；表頭依 sheet 類型分三種：一般（單行 th）、RDAC（前兩欄固定文字，其後可編輯 + 加/刪欄按鈕）、MAC/SAC（三行表頭：headers、headers2、headers3）。每個儲存格為 `<input>`，帶 `data-sheet`、`data-row`、`data-col`；TableMapping 的 input 為 readonly。

### 選取系統
- **單格：** `state.activeCell`；點擊設定、鍵盤導覽更新。
- **矩形選取：** Shift + 點擊 → `state.selection`（startRow/startCol、endRow/endCol）。
- **不連續多格：** Ctrl/Cmd + 點擊 → `state.multiSelection` 陣列。
- **依賴：** `getInputAt(row, col)` 依 `data-sheet/data-row/data-col` 查詢；變更表格 DOM 或未在 `renderTable()` 後正確綁定會導致選取/貼上異常。

### 複製/貼上規則
- **格式：** TSV（Tab 分隔）；複製/剪下寫入系統剪貼簿，貼上時解析 TSV、以 activeCell 或 selection 為目標範圍；剪下後貼上會清空來源格（若 lastClipboardOp === 'cut' 且 cutCells 存在）。
- **實作：** `doCopy`、`doCut`、`doPaste`、`buildClipboardTextFromCells`、`writeClipboardText`、`readClipboardText`。

### 新增/刪除列
- **新增列：** `addRow()` — 在當前 sheet 的 `data` 尾端 push 一列空白，欄數與 headers 一致；`renderTable()` 後更新 UI。
- **刪除列：** 每列右側有刪除按鈕，呼叫 `deleteRow(sheetName, rowIndex)`；若刪到 0 列會補回一列空白。TableMapping 不顯示刪除按鈕。

### 驗證與格式
- **驗證：** `validateCell(sheetName, rowIndex, colIndex)` 檢查必填（透過 `isRequired`）；空行不驗證。不合時加 `.cell-error`、placeholder 顯示錯誤訊息；`updateValidationSummary()` 彙總顯示於 `#validationSummary`。
- **欄位定義來源：** 從 `state.data[sheetName].headers`（及 MAC/SAC 的 headers2、headers3）讀取；預設來自 template-data，執行期僅透過 app.js 修改。

---

## 7. Dynamic Column System

### 支援的工作表
- **Resource Driver(Actvity Center)（RDAC）：** 前 2 欄固定，第 3 欄起可為「Driver Code N」、可新增/刪除欄位。
- **Resource Driver(M. A. C.)（MAC）、Resource Driver(S. A. C.)（SAC）：** 前 3 欄固定，第 4 欄起為「Driver N」、三行表頭（headers/headers2/headers3）、可新增/刪除欄位。
- **其餘 Period Data 工作表：** 可透過右鍵或表頭「+」呼叫 `insertColumnRight` 在右側插入使用者欄位；使用者欄位可刪除、可重新命名。

### 預設欄 vs 使用者新增欄
- **區分方式：** `state.data[sheetName].userAddedColIds` 陣列與 headers 同長；若 `userAddedColIds[colIndex]` 有值則視為使用者新增欄。RDAC/MAC/SAC 若無 userAddedColIds，則以 colIndex ≥ 4（RDAC 為 ≥ 2 的「非前兩欄」邏輯在 isUserAddedColumn 中為 colIndex >= 4）作為後備判斷。
- **維護：** `ensureUserAddedColIds(sheet)` 確保陣列長度與 headers 一致；上傳時對「超出 template 欄數」的欄位寫入 `userAddedColIds`；新增欄時 push 新 id（如 `u_` + Date.now()）。

### 新增欄
- **RDAC：** 表頭第 4 欄（colIndex 3）的「+」→ `addDriverCodeColumn()`，新增「Driver Code N」、更新 headers 與 data 列、userAddedColIds。
- **MAC/SAC：** 第 4 欄（colIndex 3）的「+」→ `addDriverCodeColumnMAC()` / `addDriverCodeColumnSAC()`，新增「Driver N」、同步 headers2/headers3（`ensureMACHeaderRows`/`ensureSACHeaderRows`）。
- **其他 Period 表：** `insertColumnRight(sheetName, colIndex)` — 在 colIndex+1 插入新欄、名稱由 `makeUniqueHeaderName` 產生、記錄 Undo、ChangeLog、重新渲染。

### 刪除欄
- **僅允許「使用者新增欄」：** `deleteUserColumn(sheetName, colIndex)` 會檢查 `isUserAddedColumn`；RDAC/MAC/SAC 的「×」按鈕也是呼叫 `deleteUserColumn` 或對應的 `deleteDriverCodeColumn*`。
- **實作：** `removeColumn` 從 headers、headers2/3、userAddedColIds、data 每列 splice 掉該欄；刪除前會 recordChangeSpecial 供 Undo。

### 重新命名欄
- **邏輯：** `renameColumn(sheetName, colIndex, newName, opts)` — 僅對 `isUserAddedColumn` 為 true 的欄位生效；新名稱經 `makeUniqueHeaderName` 去重；可選 skipUndo/skipLog。
- **觸發：** RDAC/MAC/SAC 表頭 input blur 時若值有變更即呼叫；一般 Period 表頭可雙擊 prompt 或透過既有 input blur。

### 欄位 metadata 儲存位置
- **state.data[sheetName].headers**（必備）
- **state.data[sheetName].headers2 / headers3**（僅 MAC/SAC）
- **state.data[sheetName].userAddedColIds**（可選，與 headers 同長）

### UI 表頭按鈕位置
- 均在 `renderTable()` 內：RDAC 在 colIndex >= 2 的 th 中組出 th-input、+、×；MAC/SAC 在 colIndex >= 3 的 th 中組出三行 input 與 +、×；一般 Period 表為 `th-has-insert` 的 + 與使用者欄的 ×、contextmenu 呼叫 `insertColumnRight`。

### 安全修改入口
- 新增「可動態欄位」的 sheet：在 template-data 定義 sheet（必要時設 `headerRows`），在 app.js 的 `renderTable()` 中為該 sheet 增加一組與 RDAC/MAC/SAC 類似的分支（或共用抽象），並在 `uploadBackup` 中對該 sheet 做欄位擴充與 userAddedColIds 寫入。
- 調整「預設欄數」：改 template-data 的該 sheet 的 `headers`/`data` 欄數；上傳邏輯會以 template 為基準擴欄。

---

## 8. Upload / Import Logic

### 流程
1. **觸發：** 工具列「Upload」、`#inputUpload` change 事件 → `uploadBackup(file)`。
2. **讀檔：** FileReader readAsArrayBuffer → `XLSX.read(data, { type: "array" })`。
3. **辨識 workbook：** `detectWorkbook(wb.SheetNames)` — 依系統 sheet 名稱（如 Item<IT使用> vs 工作表2/Sheet2）或唯一 sheet 名稱判斷 ModelData 或 PeriodData；無法辨識則報錯。
4. **逐 sheet 處理：** 跳過「ChangeLog」；對每個 Excel 工作表名以 `getInternalSheetName(excelSheetName, workbookKey)` 對應到內部 sheet 名；hidden sheet 跳過。
5. **欄位對齊與擴欄：**
   - **MAC/SAC：** 若為三行表頭，取前 3 列為 headers/headers2/headers3，其餘為 data；欄數取 max(表頭列長度, 資料列長度, template 欄數)；多出的欄標記為 user-added。
   - **RDAC：** 前 2 欄用 template，其餘用上傳表頭或 `buildExtraHeaderName`；可擴欄。
   - **一般 sheet：** 欄數 = max(template 欄數, 上傳表頭長度, 上傳資料最大列長)；不足用 template，多出用上傳表頭或 `pickHeaderName`/`buildExtraHeaderName`；多出的欄標記為 user-added。
6. **覆蓋規則：** 上傳為「覆蓋」— 每個處理到的 sheet 直接 `state.data[internalName] = sheet`；沒有「合併」選項，但 ChangeLog 會與現有 changeLog 合併（依 timestamp 去重）。
7. **TableMapping：** 不論上傳內容，Period 上傳後一律以 `TABLE_MAPPING_HEADERS`/`TABLE_MAPPING_DATA` 覆寫 `state.data["TableMapping"]`。
8. **Migrations：** 上傳後執行 migrateMACHeaderNames、migrateSACHeaderNames、migrateACAPDescriptionColumn、migrateADriverDescriptionColumn、migrateADriverActivityDescriptionColumn。
9. **儲存與 UI：** `ensureAllSheets()`、`saveToStorage()`、`renderAll()`。Period 上傳僅影響當前 `activePeriod`。

### 對應與欄名
- **Excel 表名 → 內部表名：** `getInternalSheetName(excelName, workbookKey)` 在 template-data.js；依 `sheetNameInExcel` 或 key 對應。
- **欄名來源：** 上傳第一列（或 MAC/SAC 前三列）為表頭；空白時用 `buildExtraHeaderName(idx)`（Column 1, Column 2, …）。

---

## 9. Download / Export Logic

### 觸發
- 工具列「Download」→ `downloadExcel()`；若未選公司或（Period 模式下）未選期間會提示錯誤。

### 組檔與命名
- **範圍：** 僅匯出「當前 workbook」（ModelData 或 PeriodData）。
- **函數：** `downloadWorkbook(workbookKey, timestamp, periodSuffix)`。
  - **timestamp：** `new Date().toISOString().slice(0,16).replace(/[-:T]/g,"")`（YYYYMMDDHHmm）。
  - **periodSuffix：** Period 時為 `_YYYYMM`（例如 `_202401`）；Model 為空。
- **檔名：** `{workbookConfig.filename}{periodSuffix}_{state.studentId}_{timestamp}.xlsx`，例如 `PeriodData_202401_Company1_202602071200.xlsx`。

### Sheet 收集與順序
- 依 `TEMPLATE_DATA.sheets` 篩選 `workbook === workbookKey` 的 entry，依此順序產出 Excel sheet；每個 sheet 的 Excel 名為 `config.sheetNameInExcel || internalName`（超過 31 字會擋下匯出）。

### 內容規則
- **SYSTEM_EXPORT_SHEETS（Model）：** Item<IT使用>、TableMapping<IT使用>、List_Item<IT使用>、Remark 使用 template 的 headers/data，不從 state 讀。
- **exportUseTemplate：** 若 config 有 `exportUseTemplate`，用 template 的 headers/data。
- **exportAsBlank：** 若 hidden 且 exportAsBlank，寫入空白單格。
- **一般：** 使用 `state.data[internalName].headers` 與 `state.data[internalName].data`；MAC/SAC 會組三行表頭再接 data。
- **ChangeLog：** 篩選屬於當前 workbook 的 Excel 表名的 changeLog 項目，寫入「ChangeLog」工作表。
- **PeriodData 專用：** 會覆寫「Item」與「TableMapping」為固定內容；並設定部分 sheet 為 Hidden（工作表2、Sheet2、Item、TableMapping）。
- **ModelData 專用：** 部分 sheet 設為 Hidden（Remark、Item<IT使用>、TableMapping<IT使用> 等），並確保至少有一張可見。

---

## 10. Data Storage & Scoping

### localStorage 鍵名
- **格式：** `excelForm_v1_{studentId}`，其中 studentId 為公司名稱（使用者輸入、存於 sessionStorage）。

### 儲存結構（JSON）
- **根層級：** `version`、`studentId`、`lastModified`、`activeGroup`、`activeSheet`、`data`、`changeLog`；Period 時還有 `activePeriod`、`periods`。
- **Model Data：** `data`、`changeLog` 在根；一個公司一份。
- **Period Data：** 每月份一筆 `periods[YYYY-MM]` = `{ data, changeLog, activeSheet }`；根層 `activePeriod`、`activeGroup` 表示目前月份與模式。
- **Optional Tabs：** `uiPrefs.periodTabDim` 陣列存「視為 dim 的 sheet 名稱」。
- **Required 覆寫：** 存於獨立鍵 `period_required_override_v1`（JSON 物件，sheet 名 → 欄名正規化 → boolean）。

### Company / Period 隔離
- 一個公司一個 key；切換 workbook 會先 `saveToStorage()` 再切換 activeGroup；切換 period 會 `switchPeriod(period)` 載入 `periods[period]` 到 state.data/changeLog 並 `renderAll()`。

### 載入順序
1. `loadStudentData(studentId)` 讀取 key。
2. 若無資料 → `initFromTemplate()` 再 `renderAll()`。
3. 若有：解析 JSON；若無 `periods` 則做遷移（舊 Period 資料遷到預設 2024-01）；依 activeGroup 載入 root 或 periods[activePeriod]；`ensureAllSheets()`；TableMapping 強制覆寫；migrations；`renderAll()`。

### Reset 模板行為
- **Reset：** 確認後呼叫 `resetData()`。Period 且已選月份時：`initFromTemplate()` 後只保留 PeriodData 的 sheet、清空 changeLog/undo/redo、`ensureAllSheets()`、`saveToStorage()`、`renderAll()`。Model 時：全量 `initFromTemplate()`、`saveToStorage()`、`renderAll()`。

---

## 11. Authentication / Cloud Features (If Present)

- **無。** 無 Supabase、無雲端同步、無 presence/heartbeat；僅本地 localStorage + sessionStorage。

---

## 12. UI Behavior Mapping Guide

| 功能 | UI 位置 | 邏輯檔案 | 配置/來源 |
|------|----------|----------|-----------|
| 分頁渲染 | `#groupedNav` | app.js：renderGroupedNav | template-data：workbooks[*].uiGroups、sheetOrder |
| 表格渲染 | `#tableHead` / `#tableBody` | app.js：renderTable | template-data：getSheetConfig、isRequired、getExcelSheetName |
| 欄位管理（+ / × / 重新命名） | renderTable 內 th | app.js：addDriverCodeColumn*、deleteUserColumn、insertColumnRight、renameColumn | state.data[sheet].headers、userAddedColIds |
| Upload | 工具列 Upload、#inputUpload | app.js：uploadBackup、bindEvents | — |
| Download | 工具列 Download、#btnDownload | app.js：downloadExcel、downloadWorkbook | template-data：workbooks[*].filename、sheets |
| 公司選擇 | header 顯示、#btnChangeStudent、#studentModal | app.js：setStudentId、checkStudentId、loadStudentData | sessionStorage.studentId |
| Period 選擇 | #periodSelectorContainer、#periodSelect、#btnCreatePeriod、#btnDeletePeriod | app.js：switchPeriod、updatePeriodSelector、createNewPeriod、deleteCurrentPeriod | localStorage periods、activePeriod |
| Workbook 切換 | .workbook-btn（Model Data / Period Data） | app.js：bindEvents、renderWorkbookToggle | state.activeGroup |
| 必填驗證 | .cell-error、#validationSummary、* 星號 | app.js：validateCell、updateValidationSummary、isRequired | template-data：required、getRequiredOverride |
| Optional Tabs | #btnOptionalTabs、#optionalTabsModal | app.js：getPeriodTabDimPrefs、setPeriodTabDimPrefs、saveOptionalTabsFromUI | PERIOD_DIM_SHEETS、uiPrefs.periodTabDim |
| Required Fields 管理 | #btnRequiredFields、#requiredFieldsModal | app.js：showRequiredFieldsModal、saveRequiredFieldsOverride | period_required_override_v1、template required |
| 新增列 | #btnAddRow | app.js：addRow | — |
| Reset | #btnReset、#resetModal、#btnResetConfirm | app.js：resetData | — |

---

## 13. Known Extension Points

- **新增 sheet：** 在 template-data.js 的 `sheets` 與 `sheetOrder` 新增；若為 Model 且需預設行數，在 app.js 的 `DEFAULT_ROWS_MODEL_MAP` 加一筆；若需在 export 中隱藏，在 downloadWorkbook 的 HIDE_SHEETS_MODEL 或 PERIOD_HIDDEN_SHEETS 加入。
- **新增動態欄位行為：** 在 `renderTable()` 為新 sheet 加一組類似 RDAC/MAC/SAC 的 thead 邏輯；在 `uploadBackup` 中對該 sheet 做欄數擴充與 userAddedColIds；必要時實作 `ensure*HeaderRows` 與 migration。
- **新增驗證規則：** 在 `validateCell` 中增加條件；或擴充 template 的 required 與 Required Fields 覆寫邏輯。
- **管理員/進階控制：** 已有 Required Fields、Optional Tabs 兩組 modal；可仿照在 index.html 加 modal、在 bindEvents 綁定、在 state 或 localStorage 存偏好。
- **新 dashboard / 報表：** 可新增獨立頁或區塊，從 `state.data` 或 localStorage 讀取同一 key 的資料後自行渲染。

---

## 14. Risk / Sensitive Areas

| 區域 | 原因 |
|------|------|
| **app.js `state` 物件** | 全域共用；新增/刪除/改名屬性需全文搜尋並確認所有讀寫處。 |
| **app.js 選取 + getInputAt** | 依賴每個 cell input 的 `data-sheet`、`data-row`、`data-col`；表格 DOM 或綁定時機變更會導致選取、複製貼上、鍵盤導覽錯誤。 |
| **app.js renderTable()** | 重建整張表與表頭；改 thead/tbody 結構或事件綁定會影響選取、編輯、動態欄按鈕與 contextmenu。 |
| **app.js updateCell()** | 儲存格變更、autoSave、ChangeLog、Undo 皆經此；繞過或改變語義會影響儲存與還原。 |
| **template-data.js** | getSheetConfig、isRequired、sheet 名稱被 app.js 大量引用；改名或結構變更須同步 app.js 常數與條件（如 DEFAULT_ROWS_MODEL_MAP、PERIOD_DIM_SHEETS、SYSTEM_EXPORT_SHEETS、hidden、TableMapping、RDAC/MAC/SAC 名稱）。 |
| **app.js 內硬編碼 sheet 名稱** | 如 "TableMapping"、"Resource Driver(Actvity Center)"、"Resource Driver(M. A. C.)"、"Resource Driver(S. A. C.)" 等；易與 template-data 不一致，修改時需一併搜尋。 |
| **上傳合併與 migrations** | uploadBackup 內欄位對齊、userAddedColIds、TableMapping 覆寫、各 migrate* 呼叫順序若被改動，可能導致舊檔匯入後結構錯誤或遺失欄位。 |

---

## 15. Current Feature Status Snapshot

- **已實作：** 公司名稱識別、Model/Period 雙 workbook、分組分頁、表格編輯、選取（單格/矩形/Ctrl 多選）、複製/剪下/貼上（TSV）、Undo/Redo、ChangeLog、必填驗證與摘要、動態欄位（RDAC/MAC/SAC 與一般 Period 表）、上傳 Excel（欄位對齊與擴欄）、下載 Excel（含 ChangeLog、Item/TableMapping 覆寫）、期間切換與建立/刪除、Required Fields 覆寫、Optional Tabs（dim）、TableMapping 唯讀、Reset 模板、自動儲存。
- **部分實作：** 必填/選填的 UI 與儲存完整，但 sheet 名稱集合在 app.js 與 template 間仍有多處常數需手動同步。
- **未實作 / 待確認：** 雲端同步、多人在線、合併上傳選項（目前僅覆蓋）；TableMapping 是否需任何使用者可編輯欄位；是否需「整份 Model+Period 一起匯出」；_archive 是否長期保留。

---

## 16. Suggested Safe Modification Workflow

1. **定位功能：** 用本文件的「UI Behavior Mapping」與「Folder / File Responsibility Map」找到對應的檔案與函數名；必要時在 app.js 或 template-data.js 搜尋關鍵字（sheet 名、函數名）。
2. **本地測試：** 以瀏覽器直接開啟 index.html（或本機 server）；準備一組測試公司與至少一個 Period，測試：切換 workbook/period、切換 sheet、編輯、複製貼上、新增刪除列、動態欄位、上傳/下載、Reset、Required/Optional 設定。
3. **避免破壞其他表：** 改 renderTable 時只改目標 sheet 的分支；改選取或剪貼時跑一輪點選、拖曳、Shift、Ctrl、Ctrl+C/V；改 template 或 storage 結構時確認 loadStudentData、saveToStorage、ensureAllSheets、uploadBackup、downloadWorkbook 都仍一致。
4. **建議提交策略：** 一次改動限於單一功能或單一層（例如只改 template 結構、或只改某個 sheet 的渲染）；提交前在兩大 workbook 下各選幾張表做快速回歸。

---

## 17. Glossary of Internal Terms

| 術語 | 說明 |
|------|------|
| **Model Data** | 第一個 workbook：基本設定（公司、事業單位、資源、作業中心等）；資料存於 localStorage 根層的 `data`、`changeLog`。 |
| **Period Data** | 第二個 workbook：月度資料（匯率、資源、動因、製令等）；資料存於 `periods[YYYY-MM].data` / `changeLog`。 |
| **Sheet** | 單一工作表；對應一筆 `state.data[sheetName]`，含 `headers`、`data`（二維陣列）；MAC/SAC 另有 `headers2`、`headers3`。 |
| **Dynamic Column** | 可由使用者新增/刪除/重新命名的欄位；以 `userAddedColIds[colIndex]` 有值標記。 |
| **Default Column** | 來自 template 的固定欄位；無 userAddedColIds 或對應為空。 |
| **Company Scope** | 以公司名稱（studentId）為範圍；一個公司一個 localStorage key（`excelForm_v1_{studentId}`）。 |
| **Period Scope** | Period Data 下以月份（YYYY-MM）為範圍；每月份一筆 `periods[YYYY-MM]`。 |
| **TableMapping** | Period Data 中的系統對照表（工作表名稱、Table、欄位對應等）；唯讀、永遠從 template-data 的 TABLE_MAPPING_* 覆寫。 |
| **RDAC / MAC / SAC** | Resource Driver(Actvity Center)、Resource Driver(M. A. C.)、Resource Driver(S. A. C.)；三張支援多行表頭與動態 Driver 欄的 Period 表。 |
| **Optional Tabs / dim** | Period 分頁中標記為「選用」的 sheet；在 nav 上以 `nav-pill-muted` 顯示；清單存於 `uiPrefs.periodTabDim`。 |
| **Required Fields Override** | 管理員可覆寫各 sheet 各欄是否必填；存於 `period_required_override_v1`。 |

---

*文件結束。若新增重要功能或重構，請更新本文件對應章節與日期標題。*
