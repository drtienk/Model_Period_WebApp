# PROJECT_STATUS.md

## 【Project Status - 2026-02-05】

---

## 1. Project Overview

- **What it does:** Browser-based form/spreadsheet editor titled "iPVMS Form". Users enter a company name and edit two workspaces: **Model Data** (basic settings) and **Period Data** (monthly data).
- **Business purpose:** Structured data entry for company/model and period-specific sheets; export/import via Excel (.xlsx).
- **Key workflows:**
  - Enter company name (modal on first use) → data scoped by company in localStorage.
  - Switch between Model Data and Period Data; in Period Data, select or create period month (YYYY-MM).
  - Switch sheets via grouped tab navigation; edit cells (input-based grid), select (click, drag, Shift, Ctrl), copy/cut/paste (TSV), add/delete rows.
  - Upload .xlsx (overwrite/merge), download .xlsx (company + timestamp), reset to template.
- **Major capabilities:** Config-driven tabs, dynamic columns on specific sheets (e.g. Resource Driver(Actvity Center)), Undo/Redo, ChangeLog, required-field validation, TableMapping read-only system table.

---

## 2. Tech Stack Summary

| Category | Technology |
|----------|------------|
| Frontend | Single HTML page, vanilla JS (no framework) |
| Storage | `localStorage` (key: `excelForm_v1_{companyName}`); `sessionStorage` for company name |
| Auth | None; company name only (sessionStorage + displayed in UI) |
| Cloud | None; local-only |
| Libraries | XLSX (CDN: cdnjs xlsx 0.18.5) |
| Styles | `css/style.css` |

---

## 3. System Architecture Map

| Layer | Files | Responsibilities |
|-------|--------|------------------|
| **UI Layer** | `index.html`, `css/style.css` | DOM skeleton (header, toolbar, workbook toggle, period selector, sheet tabs container, table `#tableHead`/`#tableBody`, modals, status bar). All styles. |
| **State / Cache Layer** | `js/app.js` (`state` object) | Single in-memory state: `studentId`, `activeGroup`, `activeSheet`, `activePeriod`, `data`, `changeLog`, selection (`activeCell`, `selection`, `multiSelection`), `editMode`, `cutCells`, `lastClipboardOp`, `undoStack`, `redoStack`, `_tx`. |
| **Business Logic Layer** | `js/app.js` | Init, load/save, render (tabs + table), selection, edit, copy/cut/paste, add/delete row, add/delete/rename column (where supported), undo/redo, validation, upload/download, period switch, migrations. |
| **Configuration Layer** | `js/template-data.js` | `TEMPLATE_DATA` (workbooks, sheets, sheetOrder), `TABLE_MAPPING_*`, helpers: `getSheetConfig`, `getSheetsForWorkbook`, `isRequired`, `getExcelSheetName`, `getInternalSheetName`. |
| **Cloud Sync Layer** | — | None; persistence is localStorage only. |

- **Note:** `_archive/legacy_grid_system/*` is not loaded; legacy only.

---

## 4. Core Data Model & Storage

### Local Storage

- **Key:** `excelForm_v1_{studentId}` (studentId = company name).
- **Value (JSON):** Root: `version`, `studentId`, `lastModified`, `activeGroup`, `activeSheet`, `data`, `changeLog`; for Period Data also `activePeriod`, `periods`.
- **Model Data:** `data` and `changeLog` at root; one set per company.
- **Period Data:** Stored under `periods[YYYY-MM]`: `{ data, changeLog, activeSheet }`. Root `activePeriod` and `activeGroup` track current period and mode.

### Company / Period Separation

- One company → one localStorage key. Model Data is one blob; Period Data is keyed by YYYY-MM inside `periods`.
- Switching workbook saves current state then loads the other; switching period loads `periods[YYYY-MM]` into `state.data` and re-renders.

### Tab Configuration

- `template-data.js`: `TEMPLATE_DATA.workbooks` (ModelData, PeriodData), each with `uiGroups` (key, label, sheets). `sheetOrder` defines order per workbook.
- `app.js`: `renderGroupedNav()` builds tab pills from `getSheetsForWorkbook(state.activeGroup)` and template; muted pills for hidden/system sheets (e.g. PERIOD_DIM_SHEETS).

### Presence / Heartbeat

- None; no real-time or presence logic.

---

## 5. Feature Ownership Map

| Feature | Main Files | Supporting Files | Notes |
|--------|------------|------------------|--------|
| Tab rendering | `app.js`: `renderGroupedNav()`, `renderWorkbookToggle()` | `template-data.js`: `TEMPLATE_DATA.workbooks`, `sheetOrder`, `getSheetsForWorkbook()` | Grouped pills by workbook; active sheet + muted logic in app.js |
| Sheet loading | `app.js`: `loadStudentData()`, `initFromTemplate()`, `ensureAllSheets()` | `template-data.js`: `TEMPLATE_DATA.sheets`, `getSheetConfig()` | TableMapping always overwritten from template-data |
| Toolbar operations | `app.js`: `bindEvents()` (Upload, Download, Reset, workbook buttons, period select, Add Row) | `index.html`: toolbar DOM | Upload → `uploadBackup()`; Download → `downloadExcel()`; Reset → confirm then reset + `initFromTemplate()` |
| Selection engine | `app.js`: `setActiveCell`, `getSelectedCells`, `getSelectedCellInputs`, `updateSelectionUI`, `getInputAt`, `getCellFromPoint`, `rectFromTwoPoints`, mouse/key handlers in `bindEvents()` | — | Depends on `data-sheet`/`data-row`/`data-col` on inputs; avoid changing DOM contract |
| Undo/Redo | `app.js`: `beginAction`, `recordChange`, `recordChangeSpecial`, `commitAction`, `undo()`, `redo()`, `clearRedo` | — | Cell edits and insert_col/rename_col recorded; undo replays changes and may switch sheet |
| Cloud save | — | — | No cloud; "save" = `saveToStorage()` → localStorage |
| Auto-save (local) | `app.js`: `autoSave()`, `saveToStorage()` | — | `updateCell` and other mutators call `autoSave()`; 500ms debounce |
| Presence tracking | — | — | Not implemented |
| Dynamic column add | `app.js`: `addDriverCodeColumn()`, `addDriverCodeColumnMAC()`, `addDriverCodeColumnSAC()`; RDAC header + button in `renderTable()` | `template-data.js`: sheet config for RDAC, MAC, SAC | RDAC: "Driver Code N"; MAC/SAC: "Driver N" + headers2/headers3 |
| Dynamic column delete | `app.js`: `deleteDriverCodeColumn()`, `deleteDriverCodeColumnMAC()`, `deleteDriverCodeColumnSAC()`; delete buttons in `renderTable()` | — | RDAC: colIndex ≥ 4; MAC/SAC: `canDeleteMacColumn`/`canDeleteSacColumn` (≥ 4) |
| Column rename | `app.js`: `renameColumn()`, `makeUniqueHeaderName()`; RDAC/MAC/SAC header `th-input` blur in `renderTable()` | — | Period Data only; optional insert column via context menu → `insertColumnRight()` |
| Period data handling | `app.js`: `getAllPeriodsFromStorage()`, `getCurrentPeriodData()`, `switchPeriod()`, `loadStudentData()` (periods branch), `saveToStorage()` (periods branch) | — | Create period UI → persists new key in `periods` and switches |
| Table rendering | `app.js`: `renderTable()` | `template-data.js`: `getSheetConfig()`, `isRequired()`, `getExcelSheetName()` | Builds `#tableHead` / `#tableBody`; special header rows for RDAC, MAC, SAC |
| Validation | `app.js`: `validateCell()`, `updateValidationSummary()` | `template-data.js`: `isRequired()` | Required columns from config; `.cell-error` and validation summary |
| Upload/Download | `app.js`: `uploadBackup()`, `downloadExcel()`, `downloadWorkbook()`, `detectWorkbook()`, XLSX parse/build | — | Upload can merge/overwrite; download current workbook only |
| ChangeLog | `app.js`: `updateCell` (and column ops) push to `state.changeLog`; export in download | — | Stored in same JSON as data |
| Copy/Cut/Paste | `app.js`: `doCopy()`, `doCut()`, `doPaste()`, `buildClipboardTextFromCells`, `writeClipboardText`, `readClipboardText` | — | TSV; cut clears source on paste when `lastClipboardOp === 'cut'` |

---

## 6. Period Data Table Editing Rules

- **Dynamic columns (add):**
  - **Resource Driver(Actvity Center):** Headers from col 2 are editable; col 3 has "+" to add "Driver Code N". Logic in `addDriverCodeColumn()` (app.js). Column definitions live in `state.data[sheetName].headers` (and data rows extended in sync).
  - **Resource Driver(M. A. C.) / (S. A. C.):** Same idea; "Driver N", with `headers2`/`headers3` kept in sync by `ensureMACHeaderRows()` / `ensureSACHeaderRows()`.
- **Column definitions:** Stored in `state.data[sheetName].headers` (and for MAC/SAC also `headers2`, `headers3`). Template default in `template-data.js`; runtime changes only in app.js.
- **Delete column:** Implemented only for RDAC, MAC, SAC: `deleteDriverCodeColumn(colIndex)`, `deleteDriverCodeColumnMAC(colIndex)`, `deleteDriverCodeColumnSAC(colIndex)`. Must keep fixed columns (e.g. colIndex ≥ 4 for deletable). UI: "×" on header in `renderTable()`.
- **Rename column:** `renameColumn(sheetName, colIndex, newName, opts)` in app.js. Used from RDAC/MAC/SAC header `th-input` blur and optionally elsewhere. Uniqueness via `makeUniqueHeaderName()`.
- **Column header UI:** Rendered in `renderTable()`: normal sheets = single row of `<th>`; RDAC = first two columns fixed text, rest `th-input` + add/delete buttons; MAC/SAC = three header rows from `headers` / `headers2` / `headers3`. Right-click header (Period Data) → `insertColumnRight(sheetName, colIndex)`.

---

## 7. Application Boot Sequence

1. **Load order (index.html):** XLSX (CDN) → `js/template-data.js` → `js/app.js`.
2. **Execution:** If `XLSX` already defined, `init()` runs immediately; else on `DOMContentLoaded`.
3. **init():** `checkStudentId()` → `bindEvents()`.
4. **checkStudentId():** Reads `sessionStorage.studentId`; if set → `setStudentId(id)` else `showStudentIdModal()`.
5. **setStudentId(id):** Save to sessionStorage + state, then `loadStudentData(studentId)`.
6. **loadStudentData():** Read localStorage `excelForm_v1_{studentId}`. If present: parse; migrate if no `periods`; set `activeGroup`/`activePeriod`; load either root `data`/`changeLog` (Model) or `periods[activePeriod]` (Period); `ensureAllSheets()`; TableMapping overwrite; migration helpers; then `renderAll()`. If absent: `initFromTemplate()` then `renderAll()`.
7. **ensureAllSheets():** For each template sheet, if missing in `state.data`, create from template (or DEFAULT_ROWS_MODEL_MAP for Model); TableMapping always from template; then `ensureMACHeaderRows()` / `ensureSACHeaderRows()`.
8. **renderAll():** `renderWorkbookToggle()` → `renderGroupedNav()` → `renderTable()`.

---

## 8. Known Constraints & Design Philosophy

- **Config-driven tabs:** Sheet list and grouping come from `template-data.js`; app.js only renders. Adding/renaming sheets in template may require syncing hardcoded names in app.js (e.g. DEFAULT_ROWS_MODEL_MAP, PERIOD_DIM_SHEETS, SYSTEM_EXPORT_SHEETS, nav-pill-muted).
- **Model vs Period workspace:** Strict separation; one state.data at a time; persistence shape differs (root vs periods[YYYY-MM]).
- **Local-first, no cloud:** All persistence is localStorage (and sessionStorage for company name).
- **Single app.js:** No modules; all logic in one file (~2.8k lines). Toolbar and table behavior are in the same file.
- **UI table as source of truth:** Comment in app.js: "UI table data = Source of Truth. Download rebuilds from headers + data arrays."
- **TableMapping read-only:** Always re-applied from template-data; not user-editable.

---

## 9. Current High-Risk Areas

- **app.js `state` object:** Shared by all features; adding/removing or renaming properties needs global search and care.
- **app.js selection + `getInputAt()`:** Relies on `data-sheet`, `data-row`, `data-col` on inputs. Changing table DOM or not re-binding after `renderTable()` can break selection/copy/paste.
- **app.js `renderTable()`:** Rebuilds entire table; any change to thead/tbody structure or event attachment affects selection, editing, and column add/delete/rename.
- **app.js `updateCell()`:** Central for cell updates, autoSave, ChangeLog, undo recording. Bypassing or changing semantics affects save and undo.
- **template-data.js:** `getSheetConfig()`, `isRequired()`, sheet names used across app.js. Renames or structure changes must be reflected in app.js constants and conditionals.
- **Hardcoded sheet names in app.js:** DEFAULT_ROWS_MODEL_MAP, PERIOD_DIM_SHEETS, SYSTEM_EXPORT_SHEETS, "TableMapping", "Resource Driver(Actvity Center)", etc. Easy to drift from template-data.

---

## 10. Recommended Safe Editing Strategy for AI Developers

- Prefer changing **template-data.js** (structure, sheet list, required fields) before touching UI or table logic in app.js; then update any app.js constants that reference sheet names.
- Avoid editing the **selection engine** (getInputAt, setActiveCell, getSelectedCells, updateSelectionUI, mouse/key handlers) unless the change is necessary; test click, drag, Shift, Ctrl, copy/paste after any table or event change.
- After **renderTable()** changes, verify: header buttons (+/×), context menu, input blur (rename), and that every cell input has correct `data-sheet`/`data-row`/`data-col`.
- Use **updateCell()** as the single path for cell value changes so autoSave, ChangeLog, and undo stay consistent.
- When adding or renaming sheets, search app.js for **DEFAULT_ROWS_MODEL_MAP**, **PERIOD_DIM_SHEETS**, **SYSTEM_EXPORT_SHEETS**, and **nav-pill-muted** and update as needed.
- Test **localStorage** compatibility: key format `excelForm_v1_{companyName}`; structure with `data`, `changeLog`, `periods` (for Period Data). Preserve existing fields when changing save/load.
- Keep **script load order**: XLSX → template-data.js → app.js.

---

## 11. Open TODO / Future Enhancement Hooks

- Sync hardcoded sheet-name sets in app.js with template-data (or derive from config) to avoid drift.
- Optional: Centralize or remove debug `console.log` in app.js (e.g. [headers], [doCopy], [doPaste]).
- Consider initial `hidden` on `#studentModal` to avoid brief flash before hide.
- TableMapping: confirm whether any fields will ever be user-editable.
- Resource Driver(Actvity Center) header edits: confirm if they should enter ChangeLog/Undo (currently blur → renameColumn does log and undo).
- Download: currently per-workbook; consider combined Model+Period export if needed.
- _archive/legacy_grid_system: confirm whether to keep for compliance or remove.
