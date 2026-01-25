# 專案狀態 PROJECT_STATUS（自動產生）
> 最後更新：2026-01-25
> 產生者：Cursor（接手者模式）

## 1) 這個 WebApp 是做什麼的（1~8 行）
- **目標使用者：** 學生（教學情境，需輸入學號以區分資料）。
- **使用情境：** Excel 教學用表單；在瀏覽器內像操作 Excel 一樣編輯兩本工作簿（Model Data、Period Data）底下的多張工作表，做資料填寫、練習、繳交。
- **核心流程（從打開網頁到完成操作）：** 開啟頁面 → 輸入學號（或沿用 session 的學號）→ 選擇 Model Data / Period Data → 點分頁切換工作表 → 在表格中編輯、新增/刪除列、複製/剪下/貼上 → 可上傳 .xlsx 覆寫/合併、下載 .xlsx、或 Reset 還原 template；資料依學號存於 localStorage。
- **主要輸入/輸出：** 輸入：學號（sessionStorage）、上傳 .xlsx；輸出：下載 .xlsx（檔名含學號與時間戳）、寫入 localStorage（key: `excelForm_v1_{學號}`，內容含 data、changeLog、activeGroup、activeSheet 等）。純前端，無後端、無 API、無資料庫；依賴 CDN 的 xlsx 函式庫。

## 2) 目前已完成的功能清單（只列「現在真的有」）
- [x] 學號輸入 modal（Enter / Continue 寫入 sessionStorage，以 excelForm_v1_{學號} 讀寫 localStorage）— `js/app.js`: `showStudentIdModal`, `hideStudentIdModal`, `setStudentId`, `checkStudentId`
- [x] Model Data / Period Data 切換（`state.activeGroup`，`renderAll`）— `js/app.js`: `bindEvents`（`.workbook-btn`）, `renderWorkbookToggle`
- [x] 分組 sheet 導覽（`uiGroups`，點分頁切換 `activeSheet`，hidden 表不顯示）— `js/app.js`: `renderGroupedNav`; `js/template-data.js`: `TEMPLATE_DATA.workbooks[].uiGroups`
- [x] 表格動態產生（`#tableHead`、`#tableBody`，`<td><input>`，`data-sheet`/`data-row`/`data-col`）— `js/app.js`: `renderTable`
- [x] 必填驗證（`cell-error`、placeholder、`updateValidationSummary`「⚠️ N required field(s) need attention」）— `js/app.js`: `validateCell`, `updateValidation`, `updateValidationSummary`; `js/template-data.js`: `isRequired`
- [x] 單格選取、拖曳框選、Shift+點範圍、Ctrl+點多選；`cell-active`、`cell-selected`、`cell-cut`、`cell-editing` — `js/app.js`: `#tableBody` mousedown、document mousemove/mouseup、`updateSelectionUI`, `setActiveCell`, `rectFromTwoPoints`, `getSelectedCells`
- [x] Select / Edit 模式（Select 時 readonly、單擊不進編輯；雙擊、F2、直接打字進 Edit）— `js/app.js`: `#tableBody` dblclick、document keydown（F2、isPrintableKey）、`focusActiveCell`, `updateSelectionUI`
- [x] 鍵盤：方向鍵、Enter、Tab 移動；Ctrl+A 全選；Ctrl+C/X/V 複製/剪下/貼上（TSV）；Esc；Delete/Backspace 清空選取格 — `js/app.js`: document keydown、`doCopy`, `doCut`, `doPaste`, `selectAll`, `clearCutState`, `clearCutStateIfEditingIntent`
- [x] Cut 視覺與搬移（`cell-cut`，貼上成功且上次為 cut 時清空來源；切 sheet/workbook、Esc、輸入等 `clearCutState`）— `js/app.js`: `doCut`, `doPaste`, `clearCutState`, `clearCutStateIfEditingIntent`
- [x] Undo / Redo（`beginAction`/`recordChange`/`commitAction`，Ctrl+Z、Ctrl+Shift+Z、Ctrl+Y）— `js/app.js`: `undo`, `redo`, `recordChange`, `beginAction`, `commitAction`; document keydown
- [x] 新增列（+ Add Row）、刪除列（每列 ×）— `js/app.js`: `addRow`, `deleteRow`; `renderTable` 內 delete 鈕 click、`#btnAddRow` click
- [x] 自動儲存（`updateCell` 後 debounce 500ms 寫 localStorage）— `js/app.js`: `autoSave`, `saveToStorage`
- [x] 上傳 .xlsx（解析後依工作表名稱判斷 Model/Period，對應 sheet 覆寫到 `state.data`，ChangeLog 合併，`ensureAllSheets`、`saveToStorage`、`renderAll`）— `js/app.js`: `uploadBackup`, `detectWorkbook`; `#inputUpload` change
- [x] 下載 .xlsx（依 `activeGroup` 組 xlsx，含 Hidden 規則、ChangeLog，檔名含學號與時間戳）— `js/app.js`: `downloadExcel`, `downloadWorkbook`; `#btnDownload` click
- [x] Reset（確認 modal 後 `initFromTemplate`、`saveToStorage`、`renderAll`）— `js/app.js`: `resetData`; `#btnReset`, `#btnResetConfirm`, `#resetModal`
- [x] Change 學號（清 sessionStorage，再開學號 modal，以新學號載入或初始化）— `js/app.js`: `#btnChangeStudent` click、`showStudentIdModal`
- [x] Resource Driver(Actvity Center) 表頭 C 欄起可編輯（`th-input`，改動寫入 `state.data[].headers`、`autoSave`；不經 `updateCell`/ChangeLog）— `js/app.js`: `renderTable` 內 `isRDAC` 分支、`th-input` input
- [ ] 一次下載「兩本」工作簿（Model + Period）— 待確認：目前 Download 只下載 `activeGroup` 對應的那一本

## 3) 專案結構與檔案責任分工（最重要）
| 檔案 | 角色/責任 | 重要函式/物件 | 會影響哪些功能 | 風險/注意事項 |
|---|---|---|---|---|
| index.html | UI 骨架、載入 xlsx CDN、template-data.js、app.js；header、toolbar、#groupedNav、#tableHead/#tableBody、學號 modal、Reset modal | #app, #dataTable, #tableHead, #tableBody, #studentModal, #resetModal, #groupedNav, #btnAddRow, #inputUpload, #btnDownload, #btnReset | 全站 | script 順序不可亂改：xlsx → template-data.js → app.js；id/class 被 app.js、style.css 依賴 |
| css/style.css | 全站樣式：:root 變數、按鈕、表格、input、.cell-active/.cell-selected/.cell-cut/.cell-editing/.cell-error、modal、分頁、狀態列 | .header, .toolbar, .workbook-btn, .nav-pill, .data-table, .modal-overlay, .status-bar | 全站 | 改 class 名須同步 app.js 與 index.html |
| js/template-data.js | 表單結構單一來源：TEMPLATE_DATA（workbooks、sheets、sheetOrder）、getSheetsForWorkbook、getSheetConfig、isRequired、getExcelSheetName、getInternalSheetName、norm | TEMPLATE_DATA, getSheetConfig, isRequired, getExcelSheetName, getInternalSheetName, getSheetsForWorkbook | 分頁、表格欄位、驗證、上傳/下載、匯出 Hidden/檔名 | 被 app.js 與上傳/下載廣用；改 sheets/workbooks、get*、isRequired 會影響畫面與匯入匯出 |
| js/app.js | 應用主邏輯：state、init、學號 modal、localStorage、renderAll/renderTable/renderGroupedNav、選取/編輯/剪貼、Undo/Redo、驗證、上傳/下載、Reset、addRow/deleteRow、bindEvents | state, init, checkStudentId, loadStudentData, bindEvents, renderAll, renderTable, updateCell, doCopy/doCut/doPaste, undo/redo, uploadBackup, downloadExcel, saveToStorage | 全站 | 高度耦合；state、getInputAt 依賴 data-sheet/row/col；改選取/編輯易牽動 updateSelectionUI、keydown、updateCell、renderTable |
| _archive/legacy_grid_system/*.js | 舊架構（contentEditable、#gridHead/#gridBody），已歸檔；index.html 未載入 | — | 無（現行未使用） | 僅供查閱；現行開發勿改、勿載入 |

## 4) 執行流程（從載入到可操作）
- 瀏覽器載入 `index.html`，解析到 `<script>` 時依序載入：  
  1) xlsx CDN（`xlsx.full.min.js`）→ 全域 `XLSX`；  
  2) `js/template-data.js` → 全域 `TEMPLATE_DATA`、`getSheetsForWorkbook`、`getSheetConfig`、`isRequired`、`getExcelSheetName`、`getInternalSheetName`；  
  3) `js/app.js` → 定義 `state`、各函式，最後若 `typeof XLSX !== "undefined"` 則直接 `init()`，否則等 `DOMContentLoaded` 再 `init()`。  
- **Entry point：** `init()`（`js/app.js` 尾段）。  
- `init()` 內：`checkStudentId()` → 有 `sessionStorage.studentId` 則 `setStudentId(id)` → `loadStudentData(trimmed)`、`hideStudentIdModal()`、更新 `#studentDisplay`；沒有則 `showStudentIdModal()`；最後 `bindEvents()`。  
- `loadStudentData(studentId)`：用 `excelForm_v1_{學號}` 讀 `localStorage`；有則 `JSON.parse` 取出 `data`、`changeLog`、`activeGroup`、`activeSheet`，`ensureAllSheets()`，沒有則 `initFromTemplate()`；然後 `renderAll()`。  
- `renderAll()` 順序：`renderWorkbookToggle()` → `renderGroupedNav()` → `renderTable()`。  
- `bindEvents()` 在 `init()` 最後執行，只綁一次；之後表格由 `renderTable` 動態重建，`renderTable` 內對每個 `input` 綁 `input`、對 delete 鈕綁 `click`；`#tableBody` 的 `mousedown`、`dblclick` 與 `document` 的 `mousemove`、`mouseup`、`keydown` 在 `bindEvents` 綁定（事件委派到 `#tableBody` 或依 `activeCell` 處理）。

## 5) 資料模型與狀態來源
- **in-memory 全域：** `state`（`studentId`, `activeGroup`, `activeSheet`, `data`, `changeLog`, `activeCell`, `selection`, `multiSelection`, `isSelecting`, `editMode`, `cutCells`, `lastClipboardOp`, `undoStack`, `redoStack`, `_tx`）；`_dragAnchor`、`saveTimeout`（`js/app.js`）。`TEMPLATE_DATA`、`getSheetsForWorkbook`、`getSheetConfig`、`isRequired`、`getExcelSheetName`、`getInternalSheetName`（`js/template-data.js`）。
- **localStorage：** key 為 `excelForm_v1_{學號}`；存 `{ version, studentId, lastModified, activeGroup, activeSheet, data, changeLog }`。寫入：`saveToStorage()`（被 `autoSave` debounce 500ms 或 `loadStudentData` 分支、`uploadBackup`、`resetData` 後呼叫）；讀取：`loadStudentData`。
- **sessionStorage：** key `studentId`；`setStudentId` 寫入、`checkStudentId` 讀取、`#btnChangeStudent` 點擊時 `removeItem`。
- **上傳 Excel：** `#inputUpload` change → `uploadBackup(file)`：`FileReader` 讀檔 → `XLSX.read` → `detectWorkbook(uploadedSheets)` 判斷 Model/Period → 對每個 `excelSheetName` 用 `getInternalSheetName(excelSheetName, workbookKey)` 對回 internal 名，非 hidden 則覆寫 `state.data[internalName]`；若有 `ChangeLog` 工作表則合併進 `state.changeLog`；`ensureAllSheets`、`saveToStorage`、`renderAll`。
- **下載 Excel：** `#btnDownload` click → `downloadExcel()` → `downloadWorkbook(workbookKey, timestamp)`：`XLSX.utils.book_new()`，依 `TEMPLATE_DATA.sheets` 與 `state.data`、`getExcelSheetName`、Hidden/exportAsBlank/exportUseTemplate 等規則組 ws，追加 ChangeLog 工作表，`XLSX.writeFile(wb, filename)`；檔名 `{filename}_{studentId}_{timestamp}.xlsx`。
- **雲端 / 登入 / 權限：** 無。

## 6) UI 組成（重要 DOM id/class，只列實際存在）
- **表格容器：** `#dataTable`、`#tableHead`、`#tableBody`、`.table-wrapper`、`.data-table`
- **分頁/工作表切換：** `#sheetTabsWrapper`、`#groupedNav`、`.grouped-nav`、`.nav-group-row`、`.nav-group-badge`、`.nav-pills`、`.nav-pill`（.active、.nav-pill-muted、.tab-dim）
- **Toolbar：** `.toolbar`、`.workbook-toggle`、`.workbook-btn`（#btnModelData、#btnPeriodData、data-workbook）、`.toolbar-actions`；`#inputUpload`（accept .xlsx）、`#btnDownload`、`#btnReset`；`#btnAddRow`（.btn.btn-small）
- **Header：** `.header`、`#studentDisplay`、`#btnChangeStudent`
- **主內容：** `.main`、`.sheet-header`、`#sheetTitle`、`.validation-hint`、`#validationSummary`（.hidden）
- **狀態列：** `.status-bar`、`#statusText`
- **Modal：** `#studentModal`（.modal-overlay；無 .hidden 時顯示）、`#studentIdInput`、`#btnSetStudent`；`#resetModal`（.modal-overlay.hidden）、`#btnResetCancel`、`#btnResetConfirm`、`.modal`、`.modal-actions`
- **儲存格 / 列：** `input` 具 `data-sheet`、`data-row`、`data-col`；`.row-num`、`.row-actions`、`.btn-delete-row`；`.cell-active`、`.cell-selected`、`.cell-cut`、`.cell-editing`、`.cell-error`；`th.required`、`.req-star`、`.th-input`（Resource Driver(Actvity Center) 表頭）

## 7) 表格互動（像 Excel 的行為）目前做到哪裡
- **選取：**  
  - 單選：`#tableBody` mousedown → 設 `activeCell`，`updateSelectionUI` 加 `cell-active`；`js/app.js`：mousedown handler。  
  - 拖曳框選：mousedown 設 `_dragAnchor`、`isSelecting`，`document` mousemove 以 `getCellFromPoint`、`rectFromTwoPoints` 更新 `selection`，mouseup 結束；`updateSelectionUI` 對範圍加 `cell-selected`。  
  - Shift+點：mousedown 時以 `selection` 起點或 `activeCell` 為錨、`rectFromTwoPoints(anchor, cell)` 設 `selection`。  
  - Ctrl+點：mousedown 時切換 `multiSelection`，可多格；`updateSelectionUI` 對 `multiSelection` 加 `cell-selected`。  
  - 全選：document keydown Ctrl+A → `selectAll()`，`selection` 覆蓋整表。  
- **輸入：**  
  - 單擊：Select 模式，`input` readonly，不進編輯。  
  - 雙擊：`#tableBody` dblclick → `editMode=true`，`input` 去 readonly、focus、`setSelectionRange` 到尾；`updateSelectionUI` 加 `cell-editing`。  
  - F2：keydown F2 且非 Edit → `editMode=true`，同上。  
  - Select 下直接打字：keydown `isPrintableKey` → `clearCutStateIfEditingIntent`、`editMode=true`、`updateCell(..., e.key)`、`setSelectionRange(1,1)`。  
  - Enter：Edit 下 keydown Enter → `updateCell` 寫回、`editMode=false`、`setActiveCell(r+1, c)`；Select 下 Enter → 僅下移。  
  - Tab：keydown Tab → 右移，到最右欄則下一列、第 0 欄；逾最後列則回第 0 列；`setActiveCell(nr, nc)`。  
- **複製/剪下/貼上：**  
  - 複製：keydown Ctrl+C → `doCopy()`：`getSelectedCells` → `buildClipboardTextFromCells`（TSV，\t 欄、\n 列）→ `writeClipboardText`；`js/app.js`：`doCopy`, `buildClipboardTextFromCells`, `writeClipboardText`。  
  - 剪下：Ctrl+X → `doCut()`：同上寫 TSV、設 `state.cutCells`、`lastClipboardOp='cut'`；`updateSelectionUI` 對 `cutCells.cells` 加 `cell-cut`。  
  - 貼上：Ctrl+V → `doPaste()`：`readClipboardText` 讀 TSV，自 `activeCell` 填起，必要時 `push` 新列；若 `lastClipboardOp==='cut'` 且 `cutCells` 存在則清空剪下來源並 `clearCutState`。  
- **刪除：** keydown Delete/Backspace（Select 下）→ `clearCutStateIfEditingIntent`、`beginAction('Clear')`、`getSelectedCells` 逐格 `updateCell(..., '')`、`commitAction`；`js/app.js`：document keydown。列刪除：`renderTable` 內每列 `.btn-delete-row` click → `deleteRow(sheet, rowIndex)`。

## 8) 已知問題 / 待辦（從程式碼與現況推得出來的）
- **問題/限制：** Download 只下載 `activeGroup` 那一本，無法一次產出 Model+Period 兩檔；`#studentModal` 在 HTML 無 `hidden`，首屏會短暫顯示再依 `init`/`checkStudentId` 決定隱藏，可能有閃爍；`app.js` 內有 `console.log`（如 `[headers]`、`[keydown] Ctrl+X`、`[doPaste] error`），疑似除錯用；Resource Driver(Actvity Center) 表頭 `th-input` 編輯不經 `updateCell`/ChangeLog/Undo，與一般儲存格行為不一致。
- **待辦：** 無從程式碼直接推得；若需「一次下載兩本」、表頭納入 ChangeLog/Undo、移除或管制 `console.log`，需另定規格。
- **風險：** `app.js` 內 `DEFAULT_ROWS_MODEL_MAP`、`PERIOD_DIM_SHEETS`、`SYSTEM_EXPORT_SHEETS`、`nav-pill-muted` 的 sheet 名、`detectWorkbook` 的 system sheet 名均寫死，若 `template-data.js` 新增/更名 sheet 或改 `uiGroups`，易不同步；`state` 形狀被多處讀寫，增刪屬性需全局搜尋；`_archive/legacy_grid_system` 為舊架構，index.html 未載入，但若誤載入會與現行 `#tableHead`/`#tableBody` 架構衝突。

## 9) 接下來我（使用者）給你新需求時，你應該先問的 5 個問題
1. 這個需求會動到「表格的選取/編輯/剪貼/Undo」嗎？若是，要一起檢查 `updateSelectionUI`、keydown、`updateCell`、`renderTable` 的互動，避免改一壞三。  
2. 會改到 `template-data.js` 的 `sheets`、`workbooks`、`sheetOrder` 或 `get*`/`isRequired` 嗎？若是，上傳/下載、分頁、驗證都要一併確認。  
3. 新功能需不需要新的 DOM id/class？若有，`index.html`、`app.js`、`style.css` 要一起改，且不能與現有 id 衝突。  
4. 需求的「邊界」：例如貼上是否支援跨 sheet、Cut 要不要支援跨 workbook；Download 是否維持「只下載當前 workbook」；表頭編輯要不要進 ChangeLog/Undo。  
5. `_archive/legacy_grid_system` 的檔案是否可以刪除或移動？還是必須保留作合規/稽核？

## 10) 待我回覆的問題清單（如果有）
- Q1: Resource Driver(Actvity Center) 表頭（`th-input`）刻意不進 ChangeLog/Undo，還是之後要納入？
- Q2: `DEFAULT_ROWS_MODEL_MAP`、`PERIOD_DIM_SHEETS`、`SYSTEM_EXPORT_SHEETS`、`nav-pill-muted` 的 sheet 名，之後若在 template-data 增刪/更名，是否有流程要同步更新 `app.js`，還是現階段手動即可？
