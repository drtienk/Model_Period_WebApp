# PROJECT_STATUS.md

- **專案狀態日期**：0125 2026  
- **專案性質**：純前端 Web App（無後端、無 API、無資料庫）

---

## 1. 專案一句話說明（給新手看）

這是一個「Excel 教學用表單」：使用者先輸入學號，接著在瀏覽器裡像操作 Excel 一樣，編輯兩大本「工作簿」（Model Data、Period Data）底下的多張工作表；每張表都是可輸入的表格，可新增／刪除列、複製／剪下／貼上、上傳或下載 xlsx 檔，資料會依學號自動存到瀏覽器的本機儲存（localStorage）。一切都在前端完成，不連後端也不連資料庫。

---

## 2. 專案整體結構概覽

- **類型**：單頁純前端 Web App，依賴一個 CDN 的 xlsx 函式庫，其餘都是本機的 HTML / CSS / JS。
- **主要畫面**：一進去會先跳出「輸入學號」的視窗，輸入後進入主畫面。主畫面有：
  - 頂部：標題、學號顯示、Change 按鈕  
  - 工具列：Model Data / Period Data 切換、Upload、Download、Reset  
  - 分頁導覽：依 workbooks 的 `uiGroups` 分組顯示各張工作表名稱，點擊可切換  
  - 主內容：目前工作表的標題、+ Add Row、表格（`#tableHead` + `#tableBody`）、必填提示、驗證摘要  
  - 底部狀態列  
- **表格**：每個儲存格是 `<td><input /></td>`，欄位結構與預設資料來自 `template-data.js`，實際資料存在 `app.js` 的 `state.data`，再從那裡 render 成畫面。

---

## 3. 檔案與職責對照表（非常重要）

| 檔名 | 主要負責的功能 | 是否為核心檔案 | 通常在什麼情況下才需要修改 |
|------|----------------|----------------|----------------------------|
| **index.html** | 整頁的 HTML 骨架：header、toolbar、sheet 分頁區、主表格容器（`#tableHead`、`#tableBody`）、學號 modal、Reset 確認 modal；以及載入 xlsx CDN、`template-data.js`、`app.js`。 | ✅ 是 | 要加/改 HTML 區塊、按鈕、id、或要加載新的 JS/CSS 時。 |
| **css/style.css** | 全站樣式：按鈕、表格、input、必填星號、選取/編輯/cut 的樣式、modal、分頁、狀態列等。 | ✅ 是 | 改按鈕長相、表格邊框、顏色、版面、modal 樣式等 UI 時。 |
| **js/template-data.js** | 單一真相來源：`TEMPLATE_DATA`（workbooks、sheets 的 headers / data / required / hidden / 匯出規則等），以及 `getSheetsForWorkbook`、`getSheetConfig`、`isRequired`、`getExcelSheetName`、`getInternalSheetName`。 | ✅ 是 | 要新增/刪除工作表、改欄位名稱、改預設資料、改必填欄位、改 Excel 匯出時的工作表名稱或隱藏/系統表規則時。 |
| **js/app.js** | 應用主邏輯：`state`、初始化、學號 modal、localStorage 讀寫、`renderAll` / `renderTable` / `renderGroupedNav`、選取（單格、框選、Ctrl+點多選、Shift+範圍）、編輯模式、剪貼（Copy/Cut/Paste）、Undo/Redo、儲存格更新與驗證、上傳 xlsx、下載 xlsx、Reset、Add/Delete 列、鍵盤與滑鼠事件綁定。 | ✅ 是 | 改表格行為、選取邏輯、剪貼、Undo/Redo、上傳/下載流程、驗證邏輯、自動儲存、或任何「做什麼」的流程時。 |
| **_archive/legacy_grid_system/**（含 selection_events.js、table_core.js、table_render_core.js） | 舊架構（contentEditable + #gridHead/#gridBody + DEFS.SELECTION_CORE），已歸檔。**移動日期**：2026-01-25。**新路徑**：`_archive/legacy_grid_system/`。**說明**：此三支檔案為舊架構，非現行系統使用，僅保留作參考。 | ❌ 否 | 僅在需要查閱舊實作時；現行開發勿改。 |

**補充**：真正在跑的只有 **index.html、css/style.css、js/template-data.js、js/app.js**，以及 CDN 的 xlsx。原 `selection_events.js`、`table_core.js`、`table_render_core.js` 已於 2026-01-25 移至 `_archive/legacy_grid_system/`，屬舊架構，非現行系統使用。

---

## 4. 使用者操作流程（從 UI 角度）

1. **開啟頁面** → 若有 `sessionStorage` 的 `studentId`，直接載入該學號的 `localStorage` 資料並進入主畫面；沒有則跳出「Enter Student ID」modal。
2. **輸入學號並 Continue** → 寫入 `sessionStorage`、`state.studentId`，以 `excelForm_v1_{學號}` 為 key 讀寫 `localStorage`；沒有則用 `template-data` 初始化；接著 `renderAll()`（工具列、分頁、表格）。
3. **工具列**  
   - **Model Data / Period Data**：切換 `state.activeGroup`，`activeSheet` 改為該 workbook 的第一張；`renderAll()`。  
   - **Upload**：選 .xlsx，用 xlsx 解析，依檔內工作表名稱判斷是 Model 或 Period，把對應的 sheet 填進 `state.data`，ChangeLog 若有則併入，`ensureAllSheets`、`saveToStorage`、`renderAll`。  
   - **Download**：依目前 `activeGroup` 組出一個 xlsx（headers + `state.data` 的資料 + 該 workbook 的 ChangeLog），部分 Model 表會設為 Excel 的 Hidden，檔名含學號與時間戳。  
   - **Reset**：開 Reset 確認 modal；確認後 `initFromTemplate`、`saveToStorage`、`renderAll`。
4. **分頁**：點分組裡的 sheet 名稱 → `state.activeSheet`、`renderAll()`；`activeGroup` 會跟著該 sheet 所屬 workbook。
5. **表格**  
   - 資料來源：`state.data[state.activeSheet]`（headers + data 二維陣列），**UI 表格即真相**；Download 是從 headers + data 重新組 xlsx。  
   - **選取**：單擊→`activeCell`、框選→`selection`、Ctrl+點→`multiSelection`；`updateSelectionUI` 會對 `input` 加 `cell-active`、`cell-selected`、`cell-cut`、`cell-editing`。  
   - **編輯**：雙擊或 F2 或 Select 模式下直接打字 → `editMode=true`，`input` 去 readonly、focus；Enter 寫回並下移，Tab 右移/換行；Esc 離開編輯。  
   - **剪貼**：Ctrl+C 複製選取格（TSV 寫入 clipboard）、Ctrl+X 剪下（寫 TSV 並設 `cutCells`，畫面上 `cell-cut`）、Ctrl+V 貼上（從 clipboard 讀 TSV，從 `activeCell` 填起，必要時自動加列；若上次是 Cut 則清空剪下來源並 `clearCutState`）。  
   - **Add Row**：`sheet.data.push(空白列)`，`autoSave`，`renderTable`。  
   - **刪除列**：每列右側有 ×，`deleteRow` 做 `splice`，若變 0 列會補一空白列。  
   - 每次 `updateCell` 會 `autoSave`（debounce 500ms 寫入 `localStorage`）、`updateValidation`（必填紅框與 placeholder）、`updateValidationSummary`。
6. **Change 學號**：清 `sessionStorage.studentId`，再開學號 modal；重新輸入後以新學號為 key 載入/初始化並 `renderAll`。

---

## 5. 現階段已存在的功能（只列事實）

- 學號輸入 modal，Enter / 按 Continue 寫入 `sessionStorage`，並以 `excelForm_v1_{學號}` 存到 `localStorage`。  
- Model Data / Period Data 切換，分組 sheet 導覽（`uiGroups`），點分頁切換 `activeSheet`。  
- 表格由 `renderTable` 動態產生：`#tableHead` 一列 th（#、欄名、最後一欄空白），`#tableBody` 每列為 `# + 多個 <td><input></td> + 刪除鈕`；input 有 `data-sheet`、`data-row`、`data-col`。  
- 必填驗證：`template-data` 的 `required` + `isRequired`，不合者 `input` 加 `cell-error`、placeholder；`updateValidationSummary` 顯示「⚠️ N required field(s) need attention」。  
- 單格選取、拖曳框選、Shift+點範圍、Ctrl+點多選；`cell-active`、`cell-selected`、`cell-cut`、`cell-editing` 樣式。  
- Select / Edit 模式：Select 時 input readonly、單擊不進編輯；雙擊、F2、或直接打字進 Edit。  
- 鍵盤：方向鍵、Enter、Tab 移動；Ctrl+A 全選；Ctrl+C / X / V 複製 / 剪下 / 貼上（TSV）；Esc 離開編輯或清除 cut；Delete/Backspace 清空選取格；Edit 下打字、Delete/Backspace 會 `clearCutState`。  
- Cut：Ctrl+X 設 `cutCells` 並畫 `cell-cut`；貼上成功且上次為 cut 時清空剪下來源；切 sheet/workbook、Esc、開始輸入等會 `clearCutState`。  
- Undo / Redo：`beginAction`/`recordChange`/`commitAction`，`undoStack`/`redoStack`；Ctrl+Z、Ctrl+Shift+Z / Ctrl+Y。  
- 新增列（+ Add Row）、刪除列（×）。  
- 自動儲存：`updateCell` 後 `autoSave`，debounce 500ms 呼叫 `saveToStorage` 寫 `localStorage`。  
- Upload：選 .xlsx，解析後依工作表名稱判斷 Model/Period，對應 sheet 覆寫/合併到 `state.data`，ChangeLog 合併，`ensureAllSheets`、`saveToStorage`、`renderAll`。  
- Download：依 `activeGroup` 組 xlsx（含部分 Hidden sheet、ChangeLog），檔名含學號與時間戳。  
- Reset：確認 modal 後 `initFromTemplate`、`saveToStorage`、`renderAll`。  
- Change 學號：清 `sessionStorage` 再開 modal，重新以新學號載入或初始化。

---

## 6.「之後要怎麼改 code」的導航說明（關鍵章節）

### 表格行為 / Excel-like 操作

- **選取、編輯、鍵盤、剪貼**：幾乎都在 **`js/app.js`**。  
  - 選取與 UI：`getDataBounds`、`getInputAt`、`updateSelectionUI`、`setActiveCell`、`focusActiveCell`、`getCellFromPoint`、`rectFromTwoPoints`、`selectAll`、`getSelectedCells`、`getSelectedCellInputs`。  
  - 剪貼：`writeClipboardText`、`readClipboardText`、`buildClipboardTextFromCells`、`doCopy`、`doCut`、`doPaste`、`clearCutState`、`clearCutStateIfEditingIntent`。  
  - 事件：`bindEvents` 裡對 `#tableBody` 的 `mousedown`、`dblclick`，`document` 的 `mousemove`、`mouseup`、`keydown`。  
- **儲存格更新、驗證、Undo/Redo**：**`js/app.js`** 的 `updateCell`、`recordChange`/`beginAction`/`commitAction`、`updateValidation`、`validateCell`、`updateValidationSummary`、`undo`/`redo`。  
- **表格長相（欄、列、input 怎麼畫）**：**`js/app.js`** 的 `renderTable`（thead/th、tbody/tr/td/input、row-num、row-actions、delete 鈕）。  
- **欄位與預設結構**：**`js/template-data.js`** 的 `TEMPLATE_DATA.sheets`（headers、data、required、hidden、export 相關）、`getSheetConfig`、`isRequired`。

### UI（按鈕、版面、樣式）

- **版面與元件**：**`index.html`**（header、toolbar、sheet 區、main、modals、`#tableHead`/`#tableBody`）。  
- **樣式**：**`css/style.css`**（`:root` 變數、.btn、.workbook-btn、.nav-pill、.data-table、input、.cell-active、.cell-selected、.cell-cut、.cell-editing、.cell-error、.modal-overlay、.status-bar 等）。

### 高度耦合、風險高，不應隨便動

- **`js/app.js`**：同時負責 state、init、render、選取、編輯、剪貼、鍵盤、驗證、Undo/Redo、上傳、下載、Reset、儲存，且直接依賴 `state`、`getInputAt`、`#tableHead`/`#tableBody` 的 DOM 結構。改選取或編輯時，易牽動 `updateSelectionUI`、`bindEvents` 的 keydown、以及 `updateCell`/`renderTable`。  
- **`js/template-data.js`**：`getSheetConfig`、`getExcelSheetName`、`getInternalSheetName`、`isRequired` 被 app.js 與下載/上傳邏輯廣用；欄位名、`sheetNameInExcel`、`hidden`、`exportAsBlank`、`exportUseTemplate` 等一改就會影響畫面、驗證、匯入匯出。  
- **`state` 的形狀**：`activeCell`、`selection`、`multiSelection`、`editMode`、`cutCells`、`lastClipboardOp`、`undoStack`、`redoStack`、`_tx` 等被多處讀寫；增刪或改名要全局搜尋。

### 整個 App 的進入點（初始化流程）

- **`index.html`** 依序載入：xlsx CDN → `js/template-data.js` → `js/app.js`。  
- **`js/app.js`** 結尾：若已有 `XLSX` 則直接 `init()`，否則在 `DOMContentLoaded` 後再 `init()`。  
- **`init()`**：`checkStudentId()` → 有 `sessionStorage.studentId` 則 `setStudentId` → `loadStudentData`，沒有則 `showStudentIdModal`；最後 `bindEvents()`。  
- **`loadStudentData`**：用 `excelForm_v1_{學號}` 讀 `localStorage`，有則解析 `data`、`changeLog`、`activeGroup`、`activeSheet` 並 `ensureAllSheets`，沒有則 `initFromTemplate()`；然後 `renderAll()`。  
- **`renderAll`**：`renderWorkbookToggle`、`renderGroupedNav`、`renderTable`。

---

## 重要開發原則（必須遵守）

- 本專案為 **純前端 Web App**，無後端、無 API、無資料庫。  
- 後續修改時，**避免一直新增新的 JS 檔案**；優先在 **`app.js`、`template-data.js`、`style.css`、`index.html`** 既有結構內擴充或調整行為。  
- 此文件的目的，是讓 AI 或接手者知道 **「要去哪裡改」**，不是一般性教學文件。

---

## 待釐清問題清單

1. ~~**`js/selection_events.js`、`js/table_core.js`、`js/table_render_core.js` 未在 `index.html` 中載入**……~~ **【已處理】** 2026-01-25 已將三支檔案移至 `_archive/legacy_grid_system/`。說明：此三支檔案為舊架構，非現行系統使用，僅保留作參考。

2. ~~若 `selection_events.js` / `table_core.js` / `table_render_core.js` 未來確定不採用，是否要從專案中移除或移到 `_archive` 等目錄……~~ **【已處理】** 已移入 `_archive/legacy_grid_system/`（移動日期：2026-01-25；新路徑：`_archive/legacy_grid_system/`）。
