## 【Project Status - 2026-02-05】

> 產出時間：2026-02-05  
> 產出者：Cursor AI（全面掃描後差異更新）

---

## 🔄 2026-02-05 Delta Summary

- 掃描範圍：index.html、css/style.css、js/app.js、js/template-data.js、_archive/；無其他 js utilities 目錄。
- app.js 行數為 2647（與前版記載 2648 微調）。
- Undo/Redo 僅支援「儲存格變更」一種 action 型別（changes[].sheet/row/col/oldValue/newValue）；無 INSERT_COL、RENAME_COL。
- 動態欄位／表頭編輯僅存在於 RDAC、MAC、SAC 三張 sheet（addDriverCodeColumn / deleteDriverCodeColumn + .th-input）；其餘 Period 表無「插入欄位右側」或「重新命名欄位」之程式碼。
- 上傳 XLSX：僅 Resource Driver(Actvity Center) 與 MAC/SAC 會依檔案第一列或三列表頭寫入 headers；其餘 Period sheet 使用 template headers，上傳後多出的欄位不會保留。
- _archive/legacy_grid_system/ 三支檔案內含完整程式碼（selection_events、table_core、table_render_core），但 index.html 未引用，目前為未使用之舊架構。
- state、localStorage schema、renderTable 之 DOM 結構與前版描述一致，無重大變更。
- 技術債：app.js 單檔過大、多處硬編碼 sheet 名稱、undo 僅處理 cell 型別，擴充欄位操作需改 action 結構。

---

### A. 專案在做什麼（以「現在」為主）

**無重大變更。**

- **核心用途：** 這是一個在瀏覽器中運行的表單編輯器，標題顯示為 "iPVMS Form"。使用者可以輸入公司名稱（程式碼中稱為 `studentId`），然後編輯兩種類型的資料：Model Data 和 Period Data。

- **使用者在畫面上可以做的事情：**（同上）
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

- **它「像不像 Excel / 表單 / 教學工具 / 資料編輯器」：**（同上，無變更）

---

### B. 目前的整體結構（高層）

**無重大變更。**

- **是否為單頁應用：** 是。只有一個 `index.html` 檔案，所有功能都在這個頁面中完成。

- **是否有 Model / Period / Workbook / Sheet 這類概念：**（同上，無變更）

- **資料大致怎麼流動：**（同上，無變更）

---

### C. 檔案與職責對照表（非常重要）

**精準補強：**

| 檔名 | 目前看起來負責的功能 | 分類（核心/輔助/歷史遺留） |
|------|---------------------|---------------------------|
| `index.html` | UI 骨架、載入外部資源（xlsx CDN、template-data.js、app.js）、定義所有 DOM 元素（header、toolbar、表格容器、modal 等） | 核心 |
| `css/style.css` | 全站樣式定義（顏色變數、按鈕、表格、input、選取狀態、modal、分頁、.th-input、.btn-add-column、.row-actions 等） | 核心 |
| `js/template-data.js` | 定義所有工作表的結構（TEMPLATE_DATA 物件，包含 workbooks、sheets、sheetOrder）、提供工具函式（getSheetConfig、isRequired、getExcelSheetName、getInternalSheetName 等） | 核心 |
| `js/app.js` | 應用主邏輯（state 管理、初始化、事件綁定、表格渲染、選取/編輯/剪貼、Undo/Redo、驗證、上傳/下載、localStorage 操作等）；約 2647 行 | 核心 |
| `_archive/legacy_grid_system/selection_events.js` | 舊架構的選取事件與視覺覆蓋層（DEFS.SELECTION_EVENTS），含實際程式碼；**index.html 未載入** | 歷史遺留 |
| `_archive/legacy_grid_system/table_core.js` | 舊架構的表格核心（DEFS.TABLE_CORE），含 ensureSize、parseClipboardGrid 等；**index.html 未載入** | 歷史遺留 |
| `_archive/legacy_grid_system/table_render_core.js` | 舊架構的表格渲染（DEFS.TABLE_RENDER），含 makeHeaderInput、renderHeaderDefault 等；**index.html 未載入** | 歷史遺留 |
| `.gitattributes` | Git 設定檔 | 輔助 |
| `PROJECT_STATUS.md` | 專案狀態文件（本檔案） | 輔助 |

⚠️ **觀察到的行為：**
- `index.html` 中**沒有載入** `_archive/legacy_grid_system/` 下的任何檔案。
- `_archive/` 下三支檔案**內含完整邏輯**（非僅 console.log），但未被現有頁面使用；與現行 app.js 的 `#tableHead` / `#tableBody` 架構為兩套並存之舊版實作。

---

### D. 已實作的功能（只列出確定存在的）

**差異更新與補強：**

- **表格動態產生：**（無變更）`renderTable()` 根據 `state.data[state.activeSheet]` 動態產生 `<thead>` 和 `<tbody>`，每個儲存格為 `<input>`，具 `data-sheet`、`data-row`、`data-col` 屬性。

- **選取功能：**（無變更）

- **輸入功能：**（無變更）

- **複製/剪下/貼上：**（無變更）

- **新增/刪除列：**（無變更）

- **必填驗證：**（無變更）

- **自動儲存：**（無變更）

- **上傳功能：**
  - 行為補強（從 code 確認）：`uploadBackup()` 使用 `detectWorkbook()` 辨識 ModelData / PeriodData。
  - **Resource Driver(Actvity Center)**：第一列作為 headers，且自第 3 欄起保留上傳的欄位名（動態欄位可經上傳保留）。
  - **Resource Driver(M. A. C.) / Resource Driver(S. A. C.)**：若檔案有 ≥3 列，前三列作為 headers / headers2 / headers3；否則第一列為 headers，headers2/3 填空。
  - **其餘 Period 工作表**：使用 `templateHeaders`（config.headers），上傳檔案中多出的欄位**不會**寫入 state，下載後再上傳會丟失多餘欄位。

- **下載功能：**（無變更）依 `state.data` 的 headers + data 組裝；MAC/SAC 會輸出三列表頭。

- **Undo/Redo：**
  - **從 code 確認：** 僅處理「儲存格變更」型別。`undo()` / `redo()` 皆假設 `action.changes[i]` 具 `sheet, row, col, oldValue, newValue`，並呼叫 `updateCell(..., { skipLog: true, skipUndo: true })`。**不存在** `insert_col`、`rename_col` 等 action 型別。

- **Period 管理（Period Data 模式）：**（無變更）

- **特殊工作表功能（精準補強）：**
  - **Resource Driver(Actvity Center)**：第 3 欄起為 `.th-input` 可編輯；第 3 欄有「+」按鈕可呼叫 `addDriverCodeColumn()`，第 4 欄起有「×」可呼叫 `deleteDriverCodeColumn(colIndex)`。僅 PeriodData 模式有效。
  - **Resource Driver(M. A. C.) / Resource Driver(S. A. C.)**：3 行表頭（headers、headers2、headers3），`ensureMACHeaderRows()` / `ensureSACHeaderRows()` 維持欄數一致；第 4 欄有「+」新增 Driver 欄、第 5 欄起可「×」刪除欄（`addDriverCodeColumnMAC/SAC`、`deleteDriverCodeColumnMAC/SAC`）。表頭為 `.th-input` 可編輯。
  - **TableMapping**：唯讀，系統對照表；`updateCell` 內若 `sheetName === "TableMapping"` 直接 return。
  - **其餘 Period 表**：表頭為靜態文字（`escapeHtml(h)`），**無**「插入欄位右側」或「雙擊重新命名欄位」之 UI 或函式。

---

### E. 明顯的限制與設計原則（如果能從 code 看出）

**無重大變更。**

- **避免模組化：** 所有邏輯集中在 `app.js`（約 2647 行），未使用 ES6 modules。
- **避免 framework：** 純原生 JavaScript。
- **單一 HTML 檔案、localStorage 為唯一儲存、依賴 CDN、硬編碼 sheet 名稱、UI 表格為單一真實來源**：同上。

---

### F. 新接手者「在改 code 前一定要知道的事」

**精準補強：**

- **不要亂動的檔案：**（同上）  
  - 另：`_archive/` 三支檔案雖未載入，但內含完整舊架構邏輯，刪除前建議確認無合規／稽核需求。

- **核心假設：**（同上）  
  - 補充：`state.data[sheetName]` 可含 `headers2`、`headers3`（MAC/SAC）；其餘同前。

- **改錯會整個壞掉的地方：**（同上）  
  - 補充：`undo()` / `redo()` 僅處理 `updateCell()` 路徑；若未來新增「插入欄／重新命名欄」等，需擴充 `action.changes` 的型別與還原邏輯，否則會錯用 `ch.row`/`ch.col` 導致異常。

- **script 載入順序：**（無變更）

---

### G. 不確定或需要向原作者確認的問題（列清單）

**無重大變更。** 以下維持原清單，僅將第 2、9 項依 code 事實微調表述：

1. **`_archive/legacy_grid_system/` 下的檔案是否可以刪除？** 還是必須保留作合規/稽核用途？（註：三支檔案內皆有完整程式碼，非空殼。）

2. **`Resource Driver(Actvity Center)`、MAC、SAC 的表頭編輯（.th-input）不經 ChangeLog/Undo 是刻意的嗎？** code 上表頭 input 的 `input` 事件僅寫入 `s.headers[idx]` 並 `autoSave()`，未呼叫 `recordChange()` 或 push changeLog。

3. **`DEFAULT_ROWS_MODEL_MAP`、`PERIOD_DIM_SHEETS`、`SYSTEM_EXPORT_SHEETS`、`nav-pill-muted` 的 sheet 名稱為硬編碼**，若在 `template-data.js` 新增或更名 sheet，是否有流程需同步更新 `app.js`？

4. **Download 是否只下載當前 workbook？** 目前 code 為是。

5. **`app.js` 中的 `console.log`**（如 `[headers]`、`[doPaste]` 等）是否為除錯用、需否移除或管制？

6. **`#studentModal` 在 HTML 中沒有初始 `hidden` class**，首屏是否會短暫顯示再隱藏，需向原作者確認。

7. **Period Data 模式下，若未選擇月份時切換月份**，系統會如何處理？從 code 看會初始化或使用第一個可用月份。

8. **TableMapping 是否永遠不允許使用者修改？** 目前 code 為唯讀。

9. **動態新增/刪除欄位（Driver Code）目前僅實作於 RDAC、MAC、SAC。** 其餘 Period 表在 code 中無「插入欄位右側」或「重新命名欄位」功能；若需求擴及全部 Period 表，需新增。

10. **遷移函式**（migrateMACHeaderNames、migrateSACHeaderNames、migrateACAPDescriptionColumn 等）是否為向下相容舊資料、之後是否仍需保留？

---

## 🧭 專案複雜度觀察（新增）

- **app.js 單檔過大（約 2647 行）**：state、事件、渲染、上傳/下載、Undo/Redo、遷移邏輯、特殊 sheet 分支全在同一檔案，耦合高，不利單元測試與局部修改。
- **renderTable() 分支多**：依 sheet 名稱分 isRDAC / isMAC / isSAC / 預設四類，表頭結構與事件綁定各自實作，新增一種表頭行為需改多處並注意 multi-header 同步（headers2/headers3）。
- **Undo/Redo 僅支援儲存格變更**：action 格式固定為 `{ sheet, row, col, oldValue, newValue }`，若未來支援插入欄、刪除欄、重新命名欄，須擴充 change 型別並在 undo/redo 內分支處理，否則易錯用欄位索引。
- **上傳邏輯依 sheet 名稱分岐**：Resource Driver(Actvity Center)、MAC、SAC 有專用解析；其餘 sheet 用 template headers，動態欄位無法經上傳還原，與「下載後再上傳」預期可能不一致。
- **硬編碼 sheet 名稱散落**：DEFAULT_ROWS_MODEL_MAP、PERIOD_DIM_SHEETS、SYSTEM_EXPORT_SHEETS、migrate*、ensure*、upload 內 internalName 比對等，與 template-data.js 的單一來源未統一，新增/更名 sheet 易遺漏。
- **TableMapping 與 Item 等系統表** 在儲存與載入時被強制覆寫為系統定義，邏輯分散在 loadStudentData、ensureAllSheets、uploadBackup、downloadWorkbook 等多處，修改時需全局搜尋。
- **_archive 與現行 app.js 並存**：兩套表格/選取架構並存，若未來要廢棄 _archive 或重啟舊架構，需明確決策與文件。
