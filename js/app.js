/**
 * Excel Teaching Form - Application Logic
 * UI table data = Source of Truth. Download rebuilds from headers + data arrays.
 */

const state = {
  studentId: null,
  activeGroup: "ModelData",
  activeSheet: "Company",
  data: {},
  changeLog: [],
  activeCell: null,
  selection: null,
  multiSelection: [],
  isSelecting: false,
  editMode: false, // Select(false): 單擊只選取、readonly；Edit(true): 雙擊/F2/直接打字後可編輯
  cutCells: null,     // { sheet, cells, token, timestamp } 或 null；貼上成功後、Esc、新選取、開始輸入時清除
  lastClipboardOp: null, // "cut" | "copy" | null；doPaste 只有當 lastClipboardOp==="cut" 且 cutCells 存在時才清空來源
  // --- Undo/Redo (START) ---
  undoStack: [],
  redoStack: [],
  _tx: null  // { id, label, changes: [] } 或 null
  // --- Undo/Redo (END) ---
};

let _dragAnchor = null;

let saveTimeout = null;

// --- Selection helpers ---

function getDataBounds() {
  const sheet = state.data[state.activeSheet];
  if (!sheet) return { rowCount: 0, colCount: 0 };
  return { rowCount: sheet.data.length, colCount: sheet.headers.length };
}

function getInputAt(row, col) {
  const tbody = document.getElementById("tableBody");
  if (!tbody) return null;
  return tbody.querySelector(
    'input[data-sheet="' + state.activeSheet + '"][data-row="' + row + '"][data-col="' + col + '"]'
  );
}

function getActiveInput() {
  if (!state.activeCell) return null;
  return getInputAt(state.activeCell.row, state.activeCell.col);
}

function isPrintableKey(e) {
  return (e.key || "").length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey;
}

function rectFromTwoPoints(a, b) {
  return {
    startRow: Math.min(a.row, b.row),
    startCol: Math.min(a.col, b.col),
    endRow: Math.max(a.row, b.row),
    endCol: Math.max(a.col, b.col)
  };
}

function updateSelectionUI() {
  const tbody = document.getElementById("tableBody");
  if (!tbody) return;
  const inputs = tbody.querySelectorAll("input[data-sheet][data-row][data-col]");
  inputs.forEach(function (inp) {
    inp.classList.remove("cell-active", "cell-selected", "cell-cut", "cell-editing");
    var isActive = state.activeCell && inp.dataset.row === String(state.activeCell.row) && inp.dataset.col === String(state.activeCell.col);
    if (isActive && state.editMode) {
      inp.removeAttribute("readonly");
      inp.classList.add("cell-active", "cell-editing");
    } else {
      inp.setAttribute("readonly", "readonly");
      if (isActive) inp.classList.add("cell-active");
    }
  });
  if (state.selection) {
    var r, c;
    for (r = state.selection.startRow; r <= state.selection.endRow; r++) {
      for (c = state.selection.startCol; c <= state.selection.endCol; c++) {
        const el = getInputAt(r, c);
        if (el) el.classList.add("cell-selected");
      }
    }
  } else if (state.multiSelection && state.multiSelection.length > 0) {
    state.multiSelection.forEach(function (cell) {
      const el = getInputAt(cell.row, cell.col);
      if (el) el.classList.add("cell-selected");
    });
  }
  // applyCutVisual: 僅在 cutCells.sheet === activeSheet 時加 .cell-cut（可與 .cell-selected 共存）
  if (state.cutCells && state.cutCells.cells && state.cutCells.sheet === state.activeSheet) {
    state.cutCells.cells.forEach(function (cell) {
      const el = getInputAt(cell.row, cell.col);
      if (el) el.classList.add("cell-cut");
    });
  }
}

function setActiveCell(row, col, o) {
  const opt = o || {};
  const doFocus = opt.focus !== false;
  const bounds = getDataBounds();
  if (bounds.rowCount === 0 || bounds.colCount === 0) return;
  const r = Math.max(0, Math.min(row, bounds.rowCount - 1));
  const c = Math.max(0, Math.min(col, bounds.colCount - 1));
  state.activeCell = { row: r, col: c };
  state.selection = null;
  state.multiSelection = [];
  if (opt.editMode === true) state.editMode = true;
  else if (opt.editMode === false) state.editMode = false;
  updateSelectionUI();
  if (doFocus) focusActiveCell();
}

function focusActiveCell() {
  if (!state.activeCell) return;
  const inp = getActiveInput();
  if (!inp) return;
  if (state.editMode) {
    inp.removeAttribute("readonly");
    inp.focus();
  } else {
    inp.setAttribute("readonly", "readonly");
    var ae = document.activeElement;
    if (ae && ae.closest && ae.closest("#tableBody")) ae.blur();
  }
}

function getCellFromPoint(x, y) {
  const el = document.elementFromPoint(x, y);
  const input = el && el.matches && el.matches("input[data-row][data-col]")
    ? el
    : (el && el.closest && el.closest("td") && el.closest("td").querySelector("input[data-row][data-col]"));
  if (!input || !input.dataset.row || !input.dataset.col) return null;
  return { row: parseInt(input.dataset.row, 10), col: parseInt(input.dataset.col, 10) };
}

function selectAll() {
  const b = getDataBounds();
  if (b.rowCount === 0 || b.colCount === 0) return;
  state.selection = { startRow: 0, startCol: 0, endRow: b.rowCount - 1, endCol: b.colCount - 1 };
  state.multiSelection = [];
  state.activeCell = { row: 0, col: 0 };
  state.editMode = false;
  updateSelectionUI();
  focusActiveCell();
}

function getSelectedCells() {
  if (state.multiSelection && state.multiSelection.length > 0) {
    return state.multiSelection.slice().sort(function (a, b) {
      return a.row !== b.row ? a.row - b.row : a.col - b.col;
    });
  }
  if (state.selection) {
    var out = [], r, c;
    for (r = state.selection.startRow; r <= state.selection.endRow; r++) {
      for (c = state.selection.startCol; c <= state.selection.endCol; c++) {
        out.push({ row: r, col: c });
      }
    }
    return out;
  }
  if (state.activeCell) return [state.activeCell];
  return [];
}

// getSelectedCellInputs: 單一真實來源，回傳已選取 cells 的 input 元素（含 active）；若無則 []，有選取時至少會回傳對應的 input
function getSelectedCellInputs() {
  var cells = getSelectedCells();
  var out = [];
  for (var i = 0; i < cells.length; i++) {
    var inp = getInputAt(cells[i].row, cells[i].col);
    if (inp) out.push(inp);
  }
  if (out.length === 0 && state.activeCell) {
    var inp = getActiveInput();
    if (inp) out.push(inp);
  }
  return out;
}

// clearCutState: 清除 cut 狀態與 .cell-cut；在貼上後、Esc、再次 Ctrl+X、Ctrl+C、開始輸入/Delete/Backspace 時呼叫
function clearCutState() {
  state.cutCells = null;
  updateSelectionUI();
}

// 只有「會改變內容」或「取消」時才清除 cut；點擊、導覽(Arrow/Enter/Tab)、F2、Ctrl+A 不 clear
function clearCutStateIfEditingIntent(e) {
  if (!e || !e.key) return;
  if (e.key === "Escape") { clearCutState(); return; }
  if (e.key === "Enter" && state.editMode) { clearCutState(); return; }
  if (!state.editMode && (isPrintableKey(e) || e.key === "Backspace" || e.key === "Delete")) { clearCutState(); return; }
}

// --- 剪貼簿 helper：優先 clipboard API，失敗時 fallback execCommand / 提示 ---
function writeClipboardText(text) {
  if (!navigator.clipboard || typeof navigator.clipboard.writeText !== "function") {
    return fallbackCopyText(text);
  }
  return navigator.clipboard.writeText(text).then(function () { return true; }).catch(function () {
    return fallbackCopyText(text);
  });
}
function fallbackCopyText(text) {
  try {
    var ta = document.createElement("textarea");
    ta.value = text;
    ta.style.cssText = "position:fixed;left:-9999px;top:0;";
    document.body.appendChild(ta);
    ta.select();
    var ok = document.execCommand("copy");
    document.body.removeChild(ta);
    return Promise.resolve(!!ok);
  } catch (e) {
    return Promise.resolve(false);
  }
}
function readClipboardText() {
  if (!navigator.clipboard || typeof navigator.clipboard.readText !== "function") {
    showStatus("Paste denied. Click a cell then use browser paste (Ctrl+V) / allow clipboard permission", "error");
    return Promise.resolve(null);
  }
  return navigator.clipboard.readText().then(function (t) { return t; }).catch(function () {
    showStatus("Paste denied. Click a cell then use browser paste (Ctrl+V) / allow clipboard permission", "error");
    return Promise.resolve(null);
  });
}

function buildClipboardTextFromCells(cells) {
  var sheet = state.data[state.activeSheet];
  if (!sheet) return "";
  var byRow = {};
  cells.forEach(function (c) {
    if (!byRow[c.row]) byRow[c.row] = [];
    byRow[c.row].push(c.col);
  });
  Object.keys(byRow).forEach(function (r) { byRow[r].sort(function (a, b) { return a - b; }); });
  var rows = Object.keys(byRow).map(Number).sort(function (a, b) { return a - b; });
  return rows.map(function (row) {
    return byRow[row].map(function (col) {
      var v = (sheet.data[row] || [])[col];
      return v != null ? String(v) : "";
    }).join("\t");
  }).join("\n");
}

function doCopy() {
  var cells = getSelectedCells();
  if (cells.length === 0) return;
  var sheet = state.data[state.activeSheet];
  if (!sheet) return;
  var text = buildClipboardTextFromCells(cells);
  console.log("[doCopy] cells=", cells.length);
  writeClipboardText(text).then(function (ok) {
    if (ok) {
      state.lastClipboardOp = "copy";
      console.log("[doCopy] write ok");
    } else {
      showStatus("Copy failed", "error");
    }
  });
}

// handleCut: 先清舊 cut，寫入剪貼簿，成功才設定 cutCells + lastClipboardOp 並套 .cell-cut
function doCut() {
  if (getSelectedCellInputs().length === 0) {
    console.log("[doCut] no inputs, skip");
    return;
  }
  clearCutState(); // 清掉舊 cut
  var cells = getSelectedCells();
  var text = buildClipboardTextFromCells(cells);
  console.log("[doCut] cells=", cells.length, "sheet=", state.activeSheet);
  writeClipboardText(text).then(function (ok) {
    if (ok) {
      state.cutCells = { sheet: state.activeSheet, cells: cells.map(function (c) { return { row: c.row, col: c.col }; }), token: Date.now(), timestamp: Date.now() };
      state.lastClipboardOp = "cut";
      updateSelectionUI();
      console.log("[doCut] write ok, cut state set");
    } else {
      showStatus("Cut/Copy failed. Try again or allow clipboard permission.", "error");
      console.log("[doCut] write failed");
    }
  });
}

function doPaste() {
  readClipboardText().then(function (text) {
    if (text === null) {
      console.log("[doPaste] readText null/failed");
      return;
    }
    if (!state.activeCell) return;
    var sheet = state.data[state.activeSheet];
    if (!sheet) return;
    if (!text || String(text).trim() === "") return;
    console.log("[doPaste] start lastOp=", state.lastClipboardOp, "hasCut=", !!state.cutCells);

    var rows = text.split(/\r?\n/);
    var data = rows.map(function (line) { return line.split("\t"); });
    if (data.length === 0) return;

    beginAction("Paste");
    var colCount = sheet.headers.length;
    var startRow = state.activeCell.row;
    var startCol = state.activeCell.col;
    var addedRows = 0;

    for (var i = 0; i < data.length; i++) {
      var targetRow = startRow + i;
      while (sheet.data.length <= targetRow) {
        sheet.data.push(Array(colCount).fill(""));
        addedRows++;
      }
      var rowData = data[i] || [];
      for (var j = 0; j < rowData.length; j++) {
        var targetCol = startCol + j;
        if (targetCol >= colCount) break;
        var val = rowData[j] != null ? String(rowData[j]) : "";
        updateCell(state.activeSheet, targetRow, targetCol, val);
      }
    }

    // hook: 只有「貼上成功」且 lastClipboardOp==="cut" 且 cutCells 存在時，才清空剪下來源（並避免與貼上目標重疊）
    if (state.lastClipboardOp === "cut" && state.cutCells && state.cutCells.sheet && state.cutCells.cells) {
      var maxCol = data.reduce(function (m, r) { return Math.max(m, r.length); }, 0) || 0;
      var endRow = startRow + data.length - 1;
      var endCol = startCol + Math.max(0, maxCol - 1);
      console.log("[doPaste] clearing cut sources, count=", state.cutCells.cells.length, "avoid overlap", startRow, startCol, endRow, endCol);
      state.cutCells.cells.forEach(function (c) {
        if (state.cutCells.sheet !== state.activeSheet) {
          updateCell(state.cutCells.sheet, c.row, c.col, "");
        } else if (c.row < startRow || c.row > endRow || c.col < startCol || c.col > endCol) {
          updateCell(state.cutCells.sheet, c.row, c.col, "");
        }
      });
      clearCutState();
    }
    commitAction();

    if (addedRows > 0) {
      var saved = { active: state.activeCell, sel: state.selection, multi: state.multiSelection ? state.multiSelection.slice() : [] };
      renderTable();
      state.activeCell = saved.active;
      state.selection = saved.sel;
      state.multiSelection = saved.multi;
      updateSelectionUI();
      focusActiveCell();
    }
  }).catch(function (e) {
    console.log("[doPaste] error", e);
    state._tx = null;
    showStatus("Paste failed or denied", "error");
  });
}

// --- Initialization ---

function init() {
  checkStudentId();
  bindEvents();
}

function checkStudentId() {
  const id = sessionStorage.getItem("studentId");
  if (id) {
    setStudentId(id);
  } else {
    showStudentIdModal();
  }
}

function setStudentId(id) {
  const trimmed = String(id).trim();
  if (!trimmed) return;
  sessionStorage.setItem("studentId", trimmed);
  state.studentId = trimmed;
  loadStudentData(trimmed);
  hideStudentIdModal();
  document.getElementById("studentDisplay").textContent = trimmed;
}

function loadStudentData(studentId) {
  state.studentId = studentId;
  const key = "excelForm_v1_" + studentId;
  const saved = localStorage.getItem(key);

  if (saved) {
    try {
      const parsed = JSON.parse(saved);
      state.data = parsed.data || {};
      state.changeLog = parsed.changeLog || [];
      state.undoStack = [];
      state.redoStack = [];
      state._tx = null;
      state.activeGroup = parsed.activeGroup || "ModelData";
      const order = getSheetsForWorkbook(state.activeGroup);
      state.activeSheet = order.includes(parsed.activeSheet) ? parsed.activeSheet : (order[0] || "Company");
      ensureAllSheets();
      showStatus("Welcome back, " + studentId);
    } catch (e) {
      console.error("Load error:", e);
      initFromTemplate();
      showStatus("New session for " + studentId);
    }
  } else {
    initFromTemplate();
    showStatus("New session for " + studentId);
  }
  renderAll();
}

function makeBlankRows(colCount, rowCount) {
  return Array(rowCount).fill(0).map(function () { return Array(colCount).fill(""); });
}

var DEFAULT_ROWS_MODEL_MAP = {
  "Company": 1,
  "Business Unit": 1,
  "Company Resource": 20,
  "Activity Center": 10,
  " Normal Capacity": 5,
  "Activity": 5,
  "Driver and Allocation Formula": 10,
  "Customer": 100,
  "Service Driver": 5
};

var PERIOD_DIM_SHEETS = new Set([
  "Resource Driver(Value Object)",
  "Resource Driver(Machine)",
  "Activity Center Driver(N. Cap.)",
  "ProductProject Driver",
  "Manufacture Order",
  "Manufacture Material",
  "Purchased Material and WIP",
  "Expected Project Value",
  "Revenue(InternalTransaction)NA",
  "Std. Workhour",
  "Std. Material(BOM)"
]);

function initFromTemplate() {
  state.data = {};
  state.changeLog = [];
  state.undoStack = [];
  state.redoStack = [];
  state._tx = null;
  for (const [sheetName, config] of Object.entries(TEMPLATE_DATA.sheets)) {
    if (!config.headers || config.headers.length === 0) {
      state.data[sheetName] = { headers: [], data: [] };
      continue;
    }
    var headers = [...config.headers];
    var data;
    if (config.workbook === "ModelData" && DEFAULT_ROWS_MODEL_MAP[sheetName] != null) {
      data = makeBlankRows(headers.length, DEFAULT_ROWS_MODEL_MAP[sheetName]);
    } else {
      data = config.data.map(function (row) { return [...row]; });
    }
    state.data[sheetName] = { headers: headers, data: data };
  }
  state.activeGroup = "ModelData";
  state.activeSheet = "Company";
}

function ensureAllSheets() {
  for (const [sheetName, config] of Object.entries(TEMPLATE_DATA.sheets)) {
    if (state.data[sheetName]) continue;
    if (!config.headers || config.headers.length === 0) {
      state.data[sheetName] = { headers: [], data: [] };
    } else {
      var headers = [...config.headers];
      var data;
      if (config.workbook === "ModelData" && DEFAULT_ROWS_MODEL_MAP[sheetName] != null) {
        data = makeBlankRows(headers.length, DEFAULT_ROWS_MODEL_MAP[sheetName]);
      } else {
        data = config.data.map(function (row) { return [...row]; });
      }
      state.data[sheetName] = { headers: headers, data: data };
    }
  }
}

// --- Student Modal ---

function showStudentIdModal() {
  document.getElementById("studentModal").classList.remove("hidden");
  document.getElementById("studentIdInput").value = "";
  document.getElementById("studentIdInput").focus();
}

function hideStudentIdModal() {
  document.getElementById("studentModal").classList.add("hidden");
}

// --- Storage ---

function saveToStorage() {
  if (!state.studentId) return;
  const key = "excelForm_v1_" + state.studentId;
  const saveData = {
    version: 1,
    studentId: state.studentId,
    lastModified: new Date().toISOString(),
    activeGroup: state.activeGroup,
    activeSheet: state.activeSheet,
    data: state.data,
    changeLog: state.changeLog
  };
  try {
    localStorage.setItem(key, JSON.stringify(saveData));
    showStatus("Saved (" + state.changeLog.length + " changes)");
  } catch (e) {
    console.error("Storage error:", e);
    showStatus("Storage full. Please Download backup first.", "error");
  }
}

function autoSave() {
  clearTimeout(saveTimeout);
  saveTimeout = setTimeout(saveToStorage, 500);
}

// --- Undo/Redo helpers (START) ---
function beginAction(label) {
  state._tx = { id: Date.now(), label: label || "Edit", changes: [] };
}
function recordChange(sheetName, rowIndex, colIndex, oldValue, newValue) {
  if (state._tx) {
    state._tx.changes.push({ sheet: sheetName, row: rowIndex, col: colIndex, oldValue: oldValue, newValue: newValue });
  } else {
    var act = { id: Date.now(), label: "Edit", changes: [{ sheet: sheetName, row: rowIndex, col: colIndex, oldValue: oldValue, newValue: newValue }] };
    state.undoStack.push(act);
    clearRedo();
  }
}
function commitAction() {
  if (state._tx && state._tx.changes.length > 0) {
    state.undoStack.push(state._tx);
    clearRedo();
  }
  state._tx = null;
}
function clearRedo() {
  state.redoStack = [];
}
function undo() {
  if (state.undoStack.length === 0) { showStatus("Nothing to undo", ""); return; }
  var action = state.undoStack.pop();
  state.redoStack.push(action);
  var last = action.changes[action.changes.length - 1];
  for (var i = 0; i < action.changes.length; i++) {
    var ch = action.changes[i];
    if (ch.sheet !== state.activeSheet) {
      state.activeSheet = ch.sheet;
      var cfg = getSheetConfig(ch.sheet);
      if (cfg && cfg.workbook) state.activeGroup = cfg.workbook;
      renderAll();
    }
    updateCell(ch.sheet, ch.row, ch.col, ch.oldValue, { skipLog: true, skipUndo: true });
    last = ch;
  }
  state.activeCell = { row: last.row, col: last.col };
  state.selection = null;
  state.multiSelection = [];
  updateSelectionUI();
  focusActiveCell();
  showStatus("Undo", "");
}
function redo() {
  if (state.redoStack.length === 0) { showStatus("Nothing to redo", ""); return; }
  var action = state.redoStack.pop();
  state.undoStack.push(action);
  var last = action.changes[action.changes.length - 1];
  for (var i = 0; i < action.changes.length; i++) {
    var ch = action.changes[i];
    if (ch.sheet !== state.activeSheet) {
      state.activeSheet = ch.sheet;
      var cfg = getSheetConfig(ch.sheet);
      if (cfg && cfg.workbook) state.activeGroup = cfg.workbook;
      renderAll();
    }
    updateCell(ch.sheet, ch.row, ch.col, ch.newValue, { skipLog: true, skipUndo: true });
    last = ch;
  }
  state.activeCell = { row: last.row, col: last.col };
  state.selection = null;
  state.multiSelection = [];
  updateSelectionUI();
  focusActiveCell();
  showStatus("Redo", "");
}
// --- Undo/Redo helpers (END) ---

// --- Cell update & ChangeLog ---

function updateCell(sheetName, rowIndex, colIndex, newValue, opts) {
  opts = opts || {};
  var skipLog = opts.skipLog;
  var skipUndo = opts.skipUndo;
  const sheet = state.data[sheetName];
  if (!sheet || !sheet.data[rowIndex]) return;
  const oldValue = sheet.data[rowIndex][colIndex];
  if (oldValue === newValue) return;

  if (!skipUndo) recordChange(sheetName, rowIndex, colIndex, oldValue, newValue);
  sheet.data[rowIndex][colIndex] = newValue;
  if (!skipLog) {
    const excelSheetName = getExcelSheetName(sheetName);
    state.changeLog.push({
      timestamp: new Date().toISOString(),
      sheet: excelSheetName,
      row: rowIndex + 2,
      column: sheet.headers[colIndex],
      oldValue: oldValue,
      newValue: newValue
    });
  }
  autoSave();
  updateValidation(sheetName, rowIndex, colIndex);
  var tbody = document.getElementById("tableBody");
  if (tbody) {
    var inp = tbody.querySelector('input[data-sheet="' + sheetName + '"][data-row="' + rowIndex + '"][data-col="' + colIndex + '"]');
    if (inp) inp.value = newValue;
  }
}

function updateValidation(sheetName, rowIndex, colIndex) {
  const input = document.querySelector(
    'input[data-sheet="' + sheetName + '"][data-row="' + rowIndex + '"][data-col="' + colIndex + '"]'
  );
  if (!input) return;
  const validation = validateCell(sheetName, rowIndex, colIndex);
  if (validation.valid) {
    input.classList.remove("cell-error");
    input.placeholder = "";
  } else {
    input.classList.add("cell-error");
    input.placeholder = validation.error;
  }
  updateValidationSummary();
}

function updateValidationSummary() {
  const summaryEl = document.getElementById("validationSummary");
  if (!summaryEl) return;
  const sheet = state.data[state.activeSheet];
  if (!sheet) return;
  let errorCount = 0;
  sheet.data.forEach((row, rowIndex) => {
    if (row.every(cell => cell === "")) return;
    sheet.headers.forEach((_, colIndex) => {
      if (!validateCell(state.activeSheet, rowIndex, colIndex).valid) errorCount++;
    });
  });
  if (errorCount > 0) {
    summaryEl.textContent = "\u26A0\uFE0F " + errorCount + " required field(s) need attention";
    summaryEl.classList.remove("hidden");
  } else {
    summaryEl.classList.add("hidden");
  }
}

function validateCell(sheetName, rowIndex, colIndex) {
  const sheet = state.data[sheetName];
  if (!sheet) return { valid: true };
  const columnName = sheet.headers[colIndex];
  const value = (sheet.data[rowIndex] || [])[colIndex];
  const rowData = sheet.data[rowIndex] || [];
  const isRowEmpty = rowData.every(cell => cell === "");
  if (isRowEmpty) return { valid: true };
  if (isRequired(sheetName, columnName) && (value === "" || value == null)) {
    return { valid: false, error: "Required" };
  }
  return { valid: true };
}

// --- Download (TWO files) ---

var SYSTEM_EXPORT_SHEETS = new Set(["Item<IT使用>", "TableMapping<IT使用>", "List_Item<IT使用>", "Remark"]);

function downloadExcel() {
  if (!state.studentId) {
    showStatus("Enter Student ID first", "error");
    return;
  }
  const timestamp = new Date().toISOString().slice(0, 16).replace(/[-:T]/g, "");
  const wbKey = state.activeGroup;
  downloadWorkbook(wbKey, timestamp);
  if (wbKey === "ModelData") {
    showStatus("Downloaded ModelData file");
  } else if (wbKey === "PeriodData") {
    showStatus("Downloaded PeriodData file");
  } else {
    showStatus("Downloaded " + wbKey + " file");
  }
}

function downloadWorkbook(workbookKey, timestamp) {
  const wb = XLSX.utils.book_new();
  const workbookConfig = TEMPLATE_DATA.workbooks[workbookKey];
  const allSheets = Object.entries(TEMPLATE_DATA.sheets).filter(function (entry) {
    return entry[1].workbook === workbookKey;
  });

  for (var i = 0; i < allSheets.length; i++) {
    var config = allSheets[i][1];
    var excelSheetName = config.sheetNameInExcel || allSheets[i][0];
    if (excelSheetName.length > 31) {
      showStatus('Error: Sheet name "' + excelSheetName + '" exceeds 31 characters. Export blocked.', "error");
      return;
    }
  }

  allSheets.forEach(function (entry) {
    const internalName = entry[0];
    const config = entry[1];
    const sheet = state.data[internalName];
    const excelSheetName = config.sheetNameInExcel || internalName;

    if (config.hidden && config.exportAsBlank) {
      const ws = XLSX.utils.aoa_to_sheet([[]]);
      XLSX.utils.book_append_sheet(wb, ws, excelSheetName);
      return;
    }

    var headers, data;
    if (SYSTEM_EXPORT_SHEETS.has(excelSheetName)) {
      headers = config.headers;
      data = config.data || [];
    } else if (config.exportUseTemplate && config.headers) {
      headers = config.headers;
      data = config.data || [];
    } else {
      if (!sheet) {
        if ((internalName === "List setting" || internalName === "Label") && config.headers && config.headers.length > 0) {
          headers = config.headers;
          data = (config.data && config.data.length > 0) ? config.data : [Array(config.headers.length).fill("")];
        } else {
          return;
        }
      } else {
        headers = sheet.headers;
        data = sheet.data;
        if (internalName === "List setting") {
          if (!headers || !Array.isArray(headers) || headers.length === 0) headers = config.headers || [];
          if (!data || !Array.isArray(data) || data.length === 0) data = [Array((headers || config.headers || []).length).fill("")];
        }
        if (internalName === "Label") {
          if (!headers || !Array.isArray(headers) || headers.length === 0 || (config.headers && headers.length !== config.headers.length)) headers = config.headers || [];
          if (!data || !Array.isArray(data) || data.length === 0) data = [Array((config.headers || headers || []).length).fill("")];
        }
      }
    }
    const wsData = [headers].concat(data);
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    XLSX.utils.book_append_sheet(wb, ws, excelSheetName);
  });

  const workbookExcelNames = allSheets.map(function (entry) {
    return (entry[1].sheetNameInExcel) ? entry[1].sheetNameInExcel : entry[0];
  });
  const relevantChanges = state.changeLog.filter(function (c) {
    return workbookExcelNames.indexOf(c.sheet) !== -1;
  });

  if (relevantChanges.length > 0) {
    const logHeaders = ["Timestamp", "Sheet", "Row", "Column", "Old Value", "New Value"];
    const logData = relevantChanges.map(function (c) {
      return [c.timestamp, c.sheet, c.row, c.column, c.oldValue, c.newValue];
    });
    const logWs = XLSX.utils.aoa_to_sheet([logHeaders].concat(logData));
    XLSX.utils.book_append_sheet(wb, logWs, "ChangeLog");
  }

  // PeriodData：覆蓋 "Item" 工作表為固定內容（headers + 17 列），確保下載時永遠是正確表格
  if (workbookKey === "PeriodData") {
    var ITEM_HEADERS = ["No.", "中文", "英文", "工作表名稱", "需求", "條件說明"];
    var ITEM_DATA = [
      ["1", "匯率", "Exchange Rate", "Exchange Rate", "Required", ""],
      ["2", "資源", "Resource", "Resource", "Required", ""],
      ["3", "資源動因(作業中心)", "Resource Driver(Activity Center)", "Resource Driver(Activity Center)", "Required", ""],
      ["4", "資源動因(標的)", "Resource Driver(Value Object)", "Resource Driver(Value Object)", "Required", ""],
      ["5", "資源動因(機台)", "Resource Driver(Machine)", "Resource Driver(Machine)", "Required", "製造業 Required；其它產業 Optional"],
      ["6", "資源動因(管理作業中心)", "Resource Driver(Management Activity Center)", "Resource Driver(M. A. C.)", "Required", ""],
      ["7", "資源動因(支援作業中心)", "Resource Driver(Supporting Activity Center)", "Resource Driver(S. A. C.)", "Required", ""],
      ["8", "作業中心動因(正常產能)", "Activity Center Driver(Normal Capacity)", "Activity Center Driver(N. Cap.)", "Optional", "正常產能法"],
      ["9", "作業中心動因(實際產能)", "Activity Center Driver(Actual Capacity)", "Activity Center Driver(A. Cap.)", "Required", ""],
      ["10", "作業動因", "Activity Driver", "Activity Driver", "Required", ""],
      ["11", "產品專案動因", "ProductProject Driver", "ProductProject Driver", "Optional", "設計產品專案為價值標的"],
      ["12", "製令單", "Manufacture Order", "Manufacture Order", "Required", "製造業 Required；其它產業 Optional"],
      ["13", "製令用料", "Manufacture Material", "Manufacture Material", "Required", "製造業 Required；其它產業 Optional"],
      ["14", "原料與半成品採購", "Purchased Material and Work-in-process Product", "Purchased Material and WIP", "Required", ""],
      ["15", "預期專案價值", "Expected Project Value", "Expected Project Value", "Optional", "設計專案為價值標的，且分攤至產品成本"],
      ["16", "銷貨收入", "Sales Revenue", "Sales Revenue", "Required", ""],
      ["17", "服務動因", "Service Driver", "Service Driver", "Required", ""]
    ];
    var itemWs = XLSX.utils.aoa_to_sheet([ITEM_HEADERS].concat(ITEM_DATA));
    if (wb.SheetNames.indexOf("Item") !== -1) {
      wb.Sheets["Item"] = itemWs;
    } else {
      XLSX.utils.book_append_sheet(wb, itemWs, "Item");
    }

    // PeriodData：覆蓋 "TableMapping" 工作表為固定系統對照表（headers + 74 列），確保下載時永遠是正確內容
    var TM_HEADERS = ["Table", "Table欄位名稱", "Excel 中文欄位名稱", "Excel 英文欄位名稱"];
    var TM_DATA = [
      ["ExExchangeRate", "BusinessUnitCurrency", "事業單位幣別", "Business Unit Currency"],
      ["ExExchangeRate", "CompanyCurrency", "公司幣別", "Company Currency"],
      ["ExExchangeRate", "ExchangeRate", "匯率值", "Exchange Rate"],
      ["ExResource", "BusinessUnitNo", "事業單位", "Business Unit"],
      ["ExResource", "ResourceNo", "資源代碼", "Resource Code"],
      ["ExResource", "ActivityCenterNo", "作業中心代碼", "Activity Center Code"],
      ["ExResource", "Amount", "金額", "Amount"],
      ["ExResource", "ValueObjectType", "價值標的類別", "Value Object Type"],
      ["ExResource", "ValueObjectNo", "價值標的代碼", "Value Object Code"],
      ["ExResource", "MachineGroupNo", "機台代碼", "Machine Code"],
      ["ExResource", "ProductNo", "產品代碼", "Product Code"],
      ["ExResourceDriverAC", "ActivityCenterNo", "作業中心代碼", "Activity Center Code"],
      ["ExResourceDriverValueObject", "BusinessUnitNo", "事業單位", "Business Unit"],
      ["ExResourceDriverValueObject", "ValueObjectType", "價值標的類別", "Value Object Type"],
      ["ExResourceDriverValueObject", "ValueObjectNo", "價值標的代碼", "Value Object Code"],
      ["ExResourceDriverMachine", "ActivityCenterNo", "作業中心代碼", "Activity Center Code"],
      ["ExResourceDriverMachine", "MachineGroupNo", "機台代碼", "Machine Code"],
      ["ExResourceDriverManagementAC", "ActivityCenterNo", "作業中心代碼", "Activity Center Code"],
      ["ExResourceDriverManagementAC", "MachineGroupNo", "機台代碼", "Machine Code"],
      ["ExResourceDriverSupportingAC", "ActivityCenterNo", "作業中心代碼", "Activity Center Code"],
      ["ExResourceDriverSupportingAC", "MachineGroupNo", "機台代碼", "Machine Code"],
      ["ExACDriverNormalCapacity", "ActivityCenterNo", "作業中心代碼", "Activity Center Code"],
      ["ExACDriverNormalCapacity", "MachineGroupNo", "機台代碼", "Machine Code"],
      ["ExACDriverNormalCapacity", "ActivityNo", "作業代碼", "Activity Code"],
      ["ExACDriverNormalCapacity", "NormalCapacityHours", "正常產能時間", "Normal Capacity Hours"],
      ["ExACDriverActualCapacity", "ActivityCenterNo", "作業中心代碼", "Activity Center Code"],
      ["ExACDriverActualCapacity", "MachineGroupNo", "機台代碼", "Machine Code"],
      ["ExACDriverActualCapacity", "SupportedActivityCenterNo", "受援作業中心代碼", "Supported Activity Center Code"],
      ["ExACDriverActualCapacity", "ActivityNo", "作業代碼", "Activity Code"],
      ["ExACDriverActualCapacity", "ActualCapacityHours", "實際產能時間", "Actual Capacity Hours"],
      ["ExACDriverActualCapacity", "ValueObjectNo", "價值標的代碼", "Value Object Code"],
      ["ExACDriverActualCapacity", "ValueObjectType", "價值標的類別", "Value Object Type"],
      ["ExACDriverActualCapacity", "ProductNo", "產品代碼", "Product Code"],
      ["ExActivityDriver", "ActivityCenterNo", "作業中心代碼", "Activity Center Code"],
      ["ExActivityDriver", "MachineGroupNo", "機台代碼", "Machine Code"],
      ["ExActivityDriver", "ActivityNo", "作業代碼", "Activity Code"],
      ["ExActivityDriver", "ActivityDriverNo", "作業動因", "Activity Driver"],
      ["ExActivityDriver", "ActivityDriverValue", "作業動因值", "Activity Driver Value"],
      ["ExActivityDriver", "ValueObjectNo", "價值標的代碼", "Value Object Code"],
      ["ExActivityDriver", "ValueObjectType", "價值標的類別", "Value Object Type"],
      ["ExActivityDriver", "ProductNo", "產品代碼", "Product Code"],
      ["ExProductProjectDriver", "ProductNo", "產品代碼", "Product Code"],
      ["ExProductProjectDriver", "ProjectDriverNo", "專案動因", "Project Driver"],
      ["ExProductProjectDriver", "ProjectDriverValue", "專案動因值", "Project Driver Value"],
      ["ExManufactureOrder", "BusinessUnitNo", "事業單位", "Business Unit"],
      ["ExManufactureOrder", "MO", "製令", "MO"],
      ["ExManufactureOrder", "ProductNo", "產品", "Product Code"],
      ["ExManufactureOrder", "Quantity", "完工數量(PC)", "Quantity"],
      ["ExManufactureOrder", "Closed", "製令關閉", "Closed"],
      ["ExManufactureMaterial", "BusinessUnitNo", "事業單位", "Business Unit"],
      ["ExManufactureMaterial", "MO", "製令", "MO"],
      ["ExManufactureMaterial", "MaterialNo", "材料", "Material Code"],
      ["ExManufactureMaterial", "Quantity", "用料數量", "Quantity"],
      ["ExManufactureMaterial", "Amount", "購入金額", "Amount"],
      ["ExPurchasedMaterialAndWIP", "BusinessUnitNo", "事業單位", "Business Unit"],
      ["ExPurchasedMaterialAndWIP", "MaterialNo", "料號", "Material Code"],
      ["ExPurchasedMaterialAndWIP", "Quantity", "本期數量", "Quantity"],
      ["ExPurchasedMaterialAndWIP", "Amount", "本期總金額", "Amount"],
      ["ExPurchasedMaterialAndWIP", "EndInventoryQty", "期末庫存數量", "End Inventory Qty"],
      ["ExPurchasedMaterialAndWIP", "Unit", "單位", "Unit"],
      ["ExPurchasedMaterialAndWIP", "EndInventoryAmount", "期末庫存金額", "End Inventory Amount"],
      ["ExExpectedProjectValue", "ProjectNo", "專案代碼", "Project Code"],
      ["ExExpectedProjectValue", "TotalProjectDriverValue", "預估專案動因總值", "Total Project Driver Value"],
      ["ExSalesRevenue", "OrderNo", "訂單編號", "Order No"],
      ["ExSalesRevenue", "CustomerNo", "顧客代碼", "Customer Code"],
      ["ExSalesRevenue", "ProductNo", "產品代碼", "Product Code"],
      ["ExSalesRevenue", "Quantity", "數量", "Quantity"],
      ["ExSalesRevenue", "Amount", "收入金額", "Amount"],
      ["ExSalesRevenue", "SalesActivityCenterNo", "銷售作業中心代碼", "Sales Activity Center Code"],
      ["ExSalesRevenue", "ShipmentBusinessUnitNo", "出貨事業單位", "Shipment Business Unit"],
      ["ExServiceDriver", "BusinessUnitNo", "事業單位", "Business Unit"],
      ["ExServiceDriver", "CustomerNo", "顧客代碼", "Customer Code"],
      ["ExServiceDriver", "ProductNo", "產品代碼", "Product Code"]
    ];
    var tmAoA = [TM_HEADERS].concat(TM_DATA).map(function (row) {
      return row.map(function (c) { return String(c == null ? "" : c).trim(); });
    });
    var tmWs = XLSX.utils.aoa_to_sheet(tmAoA);
    if (wb.SheetNames.indexOf("TableMapping") !== -1) {
      wb.Sheets["TableMapping"] = tmWs;
    } else {
      XLSX.utils.book_append_sheet(wb, tmWs, "TableMapping");
    }
  }

  // ModelData：指定分頁在輸出的 Excel 中預設為 Hidden（Hidden: 1 = hidden，非 veryHidden）
  if (workbookKey === "ModelData") {
    var HIDE_SHEETS_MODEL = new Set([
      "Remark",
      "Item<IT使用>",
      "TableMapping<IT使用>",
      "List_Item<IT使用>",
      "List setting",
      "Label"
    ]);
    wb.Workbook = wb.Workbook || {};
    wb.Workbook.Sheets = wb.SheetNames.map(function (name) {
      var o = {};
      if (HIDE_SHEETS_MODEL.has(name)) o.Hidden = 1;
      return o;
    });
    // 至少保留一張可見（避免 Excel 開啟時全部隱藏）
    var allHidden = wb.Workbook.Sheets.every(function (m) { return m.Hidden === 1; });
    if (allHidden && wb.Workbook.Sheets.length > 0) {
      wb.Workbook.Sheets[0].Hidden = 0;
    }
  }

  const filename = workbookConfig.filename + "_" + state.studentId + "_" + timestamp + ".xlsx";
  XLSX.writeFile(wb, filename);
}

// --- Upload (auto-detect workbook) ---

function uploadBackup(file) {
  if (!state.studentId) {
    showStatus("Enter Student ID first", "error");
    return;
  }
  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: "array" });
      const uploadedSheets = wb.SheetNames;
      const workbookKey = detectWorkbook(uploadedSheets);

      if (!workbookKey) {
        showStatus("Error: Could not identify file type", "error");
        return;
      }

      uploadedSheets.forEach(function (excelSheetName) {
        if (excelSheetName === "ChangeLog") return;
        const ws = wb.Sheets[excelSheetName];
        const jsonData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        const internalName = getInternalSheetName(excelSheetName, workbookKey);

        if (internalName && TEMPLATE_DATA.sheets[internalName]) {
          const config = TEMPLATE_DATA.sheets[internalName];
          if (config.hidden) return;

          const templateHeaders = config.headers;
          const uploadedData = jsonData.slice(1);
          const normalizedData = uploadedData.map(function (row) {
            const newRow = [].concat(row);
            while (newRow.length < templateHeaders.length) newRow.push("");
            return newRow.slice(0, templateHeaders.length);
          });

          state.data[internalName] = {
            headers: templateHeaders,
            data: normalizedData.length > 0 ? normalizedData : [Array(templateHeaders.length).fill("")]
          };
        }
      });

      if (wb.Sheets["ChangeLog"]) {
        const logData = XLSX.utils.sheet_to_json(wb.Sheets["ChangeLog"]);
        const restoredLogs = logData.map(function (row) {
          return {
            timestamp: row["Timestamp"],
            sheet: row["Sheet"],
            row: row["Row"],
            column: row["Column"],
            oldValue: row["Old Value"] || "",
            newValue: row["New Value"] || ""
          };
        });
        const existingTimestamps = {};
        state.changeLog.forEach(function (c) {
          existingTimestamps[c.timestamp] = true;
        });
        restoredLogs.forEach(function (log) {
          if (!existingTimestamps[log.timestamp]) {
            state.changeLog.push(log);
          }
        });
      }

      ensureAllSheets();
      saveToStorage();
      renderAll();
      showStatus("Restored " + workbookKey + " backup");
    } catch (err) {
      console.error(err);
      showStatus("Error reading file", "error");
    }
  };
  reader.readAsArrayBuffer(file);
}

function detectWorkbook(sheetNames) {
  const modelSystemSheets = ["Item<IT使用>", "TableMapping<IT使用>", "List_Item<IT使用>"];
  const periodSystemSheets = ["工作表2", "Sheet2"];
  const hasModelSystem = modelSystemSheets.some(function (s) {
    return sheetNames.indexOf(s) !== -1;
  });
  const hasPeriodSystem = periodSystemSheets.some(function (s) {
    return sheetNames.indexOf(s) !== -1;
  });
  if (hasModelSystem) return "ModelData";
  if (hasPeriodSystem) return "PeriodData";

  const modelUniqueSheets = [" Normal Capacity", "Company Resource", "Label"];
  const periodUniqueSheets = ["Exchange Rate", "Manufacture Order", "Sales Revenue "];
  const hasModelUnique = modelUniqueSheets.some(function (s) {
    return sheetNames.indexOf(s) !== -1;
  });
  const hasPeriodUnique = periodUniqueSheets.some(function (s) {
    return sheetNames.indexOf(s) !== -1;
  });
  if (hasModelUnique && !hasPeriodUnique) return "ModelData";
  if (hasPeriodUnique && !hasModelUnique) return "PeriodData";

  if (sheetNames.indexOf("Company") !== -1 && sheetNames.indexOf("Business Unit") !== -1) return "ModelData";
  if (sheetNames.indexOf("Resource") !== -1 && sheetNames.indexOf("Activity Driver") !== -1) return "PeriodData";
  return null;
}

// --- Reset ---

function resetData() {
  initFromTemplate();
  saveToStorage();
  renderAll();
  showStatus("Data reset to template");
}

// --- Rendering ---

function renderAll() {
  renderWorkbookToggle();
  renderGroupedNav();
  renderTable();
}

function renderWorkbookToggle() {
  document.querySelectorAll(".workbook-btn").forEach(function (btn) {
    btn.classList.toggle("active", btn.getAttribute("data-workbook") === state.activeGroup);
  });
  var btnDownload = document.getElementById("btnDownload");
  if (btnDownload) {
    btnDownload.textContent = state.activeGroup === "ModelData" ? "Download Model Data" : "Download Period Data";
  }
}

function renderGroupedNav() {
  const container = document.getElementById("groupedNav");
  if (!container) return;
  const wbKey = state.activeGroup;
  const wb = TEMPLATE_DATA.workbooks[wbKey];
  const groups = wb && wb.uiGroups;
  container.innerHTML = "";
  if (!groups || !groups.length) return;

  groups.forEach(function (grp) {
    const row = document.createElement("div");
    row.className = "nav-group-row";

    const badge = document.createElement("span");
    badge.className = "nav-group-badge";
    badge.textContent = grp.label;
    row.appendChild(badge);

    const pillsWrap = document.createElement("div");
    pillsWrap.className = "nav-pills";

    (grp.sheets || []).forEach(function (internalName) {
      const config = getSheetConfig(internalName);
      if (!config || config.hidden) return;

      const pill = document.createElement("button");
      pill.type = "button";
      pill.className = "nav-pill" + (internalName === state.activeSheet ? " active" : "");
      if (wbKey === "ModelData" && ["Machine(Activity Center Driver)", "Material", "ProductProject"].indexOf(internalName) !== -1) {
        pill.classList.add("nav-pill-muted");
      }
      if (wbKey === "PeriodData" && PERIOD_DIM_SHEETS.has(internalName)) {
        pill.classList.add("tab-dim");
      }
      pill.textContent = getExcelSheetName(internalName);
      pill.addEventListener("click", function () {
        state.activeGroup = wbKey;
        state.activeSheet = internalName;
        clearCutState(); // 切 sheet 時清除 cut（不支援跨 sheet 保留）
        renderAll();
        autoSave();
      });
      pillsWrap.appendChild(pill);
    });

    row.appendChild(pillsWrap);
    container.appendChild(row);
  });
}

function renderTable() {
  const sheet = state.data[state.activeSheet];
  const thead = document.getElementById("tableHead");
  const tbody = document.getElementById("tableBody");
  const config = getSheetConfig(state.activeSheet);

  state.activeCell = null;
  state.selection = null;
  state.multiSelection = [];
  state.isSelecting = false;
  state.editMode = false;
  // 同一 sheet 的 renderTable（新增列、貼上 addedRows）不清 cut；切 sheet 的清 cut 改由切換處處理
  // state.cutCells = null; 已移除
  _dragAnchor = null;

  document.getElementById("sheetTitle").textContent = getExcelSheetName(state.activeSheet);

  if (!sheet || !config || config.hidden) {
    thead.innerHTML = "";
    tbody.innerHTML = "<tr><td colspan='3'>No data</td></tr>";
    document.getElementById("btnAddRow").style.display = "none";
    document.getElementById("validationSummary").classList.add("hidden");
    return;
  }

  document.getElementById("btnAddRow").style.display = "";

  if (sheet.headers && sheet.headers.length > 0 && (!sheet.data || sheet.data.length === 0)) {
    sheet.data = [Array(sheet.headers.length).fill("")];
  }

  const headers = sheet.headers;
  console.log("[headers]", state.activeSheet, sheet.headers);
  let thHtml = "<tr><th class=\"row-num\">#</th>";
  headers.forEach(function (h) {
    const req = isRequired(state.activeSheet, h);
    const showStar = (state.activeGroup === "ModelData" && (state.activeSheet === "Company" || state.activeSheet === "Company Resource") && req);
    thHtml += "<th" + (req ? " class=\"required\"" : "") + ">" + escapeHtml(h) + (showStar ? "<span class=\"req-star\">*</span>" : "") + "</th>";
  });
  thHtml += "<th class=\"row-actions\"></th></tr>";
  thead.innerHTML = thHtml;

  tbody.innerHTML = "";
  sheet.data.forEach(function (row, rowIndex) {
    const tr = document.createElement("tr");
    tr.appendChild(createCell("td", "row-num", String(rowIndex + 2)));
    headers.forEach(function (_, colIndex) {
      const val = (row[colIndex] != null) ? String(row[colIndex]) : "";
      const validation = validateCell(state.activeSheet, rowIndex, colIndex);
      const input = document.createElement("input");
      input.type = "text";
      input.value = val;
      input.setAttribute("readonly", "readonly");
      input.dataset.sheet = state.activeSheet;
      input.dataset.row = String(rowIndex);
      input.dataset.col = String(colIndex);
      if (!validation.valid) {
        input.classList.add("cell-error");
        input.placeholder = validation.error;
      }
      input.addEventListener("input", function () {
        updateCell(state.activeSheet, rowIndex, colIndex, input.value);
      });
      const td = document.createElement("td");
      td.appendChild(input);
      tr.appendChild(td);
    });
    const actTd = document.createElement("td");
    actTd.className = "row-actions";
    const delBtn = document.createElement("button");
    delBtn.type = "button";
    delBtn.className = "btn-delete-row";
    delBtn.textContent = "\u00D7";
    delBtn.title = "Delete row";
    delBtn.addEventListener("click", function () {
      deleteRow(state.activeSheet, rowIndex);
    });
    actTd.appendChild(delBtn);
    tr.appendChild(actTd);
    tbody.appendChild(tr);
  });

  updateValidationSummary();
  updateSelectionUI();
}

function createCell(tag, className, text) {
  const el = document.createElement(tag);
  el.className = className;
  el.textContent = text;
  return el;
}

function escapeHtml(s) {
  const div = document.createElement("div");
  div.textContent = s;
  return div.innerHTML;
}

// --- Add / Delete row ---

function addRow() {
  const sheet = state.data[state.activeSheet];
  if (!sheet) return;
  const cols = sheet.headers.length;
  sheet.data.push(Array(cols).fill(""));
  autoSave();
  renderTable();
}

function deleteRow(sheetName, rowIndex) {
  const sheet = state.data[sheetName];
  if (!sheet) return;
  sheet.data.splice(rowIndex, 1);
  if (sheet.data.length === 0) {
    sheet.data.push(Array(sheet.headers.length).fill(""));
  }
  autoSave();
  renderTable();
}

// --- Status ---

function showStatus(text, type) {
  const el = document.getElementById("statusText");
  el.textContent = text;
  el.className = type === "error" ? "error" : "";
}

// --- Event binding ---

function bindEvents() {
  document.getElementById("btnSetStudent").addEventListener("click", function () {
    const v = document.getElementById("studentIdInput").value.trim();
    if (v) setStudentId(v);
  });
  document.getElementById("studentIdInput").addEventListener("keydown", function (e) {
    if (e.key === "Enter") {
      const v = document.getElementById("studentIdInput").value.trim();
      if (v) setStudentId(v);
    }
  });

  document.getElementById("btnChangeStudent").addEventListener("click", function () {
    sessionStorage.removeItem("studentId");
    showStudentIdModal();
  });

  document.querySelectorAll(".workbook-btn").forEach(function (btn) {
    btn.addEventListener("click", function () {
      const wb = btn.getAttribute("data-workbook");
      state.activeGroup = wb;
      const order = getSheetsForWorkbook(wb);
      state.activeSheet = order[0] || "Company";
      clearCutState(); // 切 workbook 時清除 cut
      renderAll();
      autoSave();
    });
  });

  document.getElementById("inputUpload").addEventListener("change", function () {
    const f = this.files && this.files[0];
    if (f) {
      uploadBackup(f);
      this.value = "";
    }
  });

  document.getElementById("btnDownload").addEventListener("click", function () {
    downloadExcel();
  });

  document.getElementById("btnReset").addEventListener("click", function () {
    document.getElementById("resetModal").classList.remove("hidden");
  });

  document.getElementById("btnResetCancel").addEventListener("click", function () {
    document.getElementById("resetModal").classList.add("hidden");
  });

  document.getElementById("btnResetConfirm").addEventListener("click", function () {
    document.getElementById("resetModal").classList.add("hidden");
    resetData();
  });

  document.getElementById("btnAddRow").addEventListener("click", function () {
    addRow();
  });

  // --- Selection: mousedown (delegation on tbody) ---
  // 點擊/框選/多選只換 active/selection，不 clear cut（Cut 後點目的地再貼上才可 move）
  document.getElementById("tableBody").addEventListener("mousedown", function (ev) {
    const input = ev.target.matches("input[data-row][data-col]")
      ? ev.target
      : (ev.target.closest && ev.target.closest("td") && ev.target.closest("td").querySelector("input[data-row][data-col]"));
    if (!input) return;
    ev.preventDefault();
    const row = parseInt(input.dataset.row, 10);
    const col = parseInt(input.dataset.col, 10);
    const cell = { row: row, col: col };

    if (ev.ctrlKey || ev.metaKey) {
      if (state.selection) {
        state.multiSelection = [];
        var r, c;
        for (r = state.selection.startRow; r <= state.selection.endRow; r++) {
          for (c = state.selection.startCol; c <= state.selection.endCol; c++) {
            state.multiSelection.push({ row: r, col: c });
          }
        }
        state.selection = null;
      } else if ((!state.multiSelection || state.multiSelection.length === 0) && state.activeCell) {
        state.multiSelection = [{ row: state.activeCell.row, col: state.activeCell.col }];
      }
      var idx = (state.multiSelection || []).findIndex(function (c) { return c.row === cell.row && c.col === cell.col; });
      if (idx >= 0) {
        state.multiSelection.splice(idx, 1);
        state.activeCell = state.multiSelection.length > 0 ? state.multiSelection[0] : cell;
      } else {
        state.multiSelection = state.multiSelection || [];
        state.multiSelection.push(cell);
        state.activeCell = cell;
      }
      state.selection = null;
      state.editMode = false;
      updateSelectionUI();
      focusActiveCell();
      return;
    }

    if (ev.shiftKey) {
      const anchor = state.selection
        ? { row: state.selection.startRow, col: state.selection.startCol }
        : state.activeCell;
      if (!anchor) {
        state.activeCell = cell;
        state.selection = null;
        state.multiSelection = [];
        state.editMode = false;
        updateSelectionUI();
        focusActiveCell();
        return;
      }
      const rect = rectFromTwoPoints(anchor, cell);
      state.selection = rect;
      state.multiSelection = [];
      state.activeCell = cell;
      state.editMode = false;
      updateSelectionUI();
      focusActiveCell();
      return;
    }

    state.activeCell = cell;
    state.selection = null;
    state.multiSelection = [];
    state.isSelecting = true;
    state.editMode = false;
    _dragAnchor = cell;
    updateSelectionUI();
    focusActiveCell(); // Select 模式：blur 表格內 input，單擊不出現文字游標
  });

  document.getElementById("tableBody").addEventListener("dblclick", function (ev) {
    const inp = ev.target.matches && ev.target.matches("input[data-row][data-col]")
      ? ev.target
      : (ev.target.closest && ev.target.closest("td") && ev.target.closest("td").querySelector("input[data-row][data-col]"));
    if (inp) {
      clearCutState(); // 確定要開始編輯，視為放棄 cut 搬移
      var row = parseInt(inp.dataset.row, 10);
      var col = parseInt(inp.dataset.col, 10);
      state.activeCell = { row: row, col: col };
      state.selection = null;
      state.multiSelection = [];
      state.editMode = true;
      inp.removeAttribute("readonly");
      inp.focus();
      inp.setSelectionRange(inp.value.length, inp.value.length);
      updateSelectionUI();
    }
  });

  document.addEventListener("mousemove", function (ev) {
    if (!state.isSelecting) return;
    const cell = getCellFromPoint(ev.clientX, ev.clientY);
    if (!cell) return;
    const rect = rectFromTwoPoints(_dragAnchor, cell);
    if (rect.startRow === rect.endRow && rect.startCol === rect.endCol) {
      state.selection = null;
    } else {
      state.selection = rect;
    }
    state.activeCell = { row: cell.row, col: cell.col };
    updateSelectionUI();
  });

  document.addEventListener("mouseup", function () {
    if (!state.isSelecting) return;
    state.isSelecting = false;
    focusActiveCell();
  });

  // handleKeydown: Ctrl+C/X/V、Esc、Enter、Tab、方向鍵、F2、Select 下打字/Delete/Backspace
  document.addEventListener("keydown", function (e) {
    const ae = document.activeElement;
    var inTable = ae && ae.matches && ae.matches("input[data-row][data-col]") && ae.closest && ae.closest("#tableBody");
    if (!state.activeCell && !inTable) return;
    if (!state.activeCell && inTable) {
      state.activeCell = { row: parseInt(ae.dataset.row, 10), col: parseInt(ae.dataset.col, 10) };
    }
    const r = state.activeCell.row;
    const c = state.activeCell.col;
    const bounds = getDataBounds();
    if (bounds.rowCount === 0 || bounds.colCount === 0) return;

    // hasTextSelection：只有 inTable（Edit 模式）且 input 有反白才 true，避免 ae 不在 table 時誤判
    var hasTextSelection = inTable && (ae && ae.selectionStart != null) && (ae.selectionStart !== ae.selectionEnd);
    var k = (e.key || "").toLowerCase();

    if ((e.ctrlKey || e.metaKey) && (k === "c" || k === "x" || k === "v")) {
      if (hasTextSelection) return; // 文字反白時讓瀏覽器原生剪下，不攔截
      if (k === "c") { e.preventDefault(); clearCutState(); doCopy(); return; }
      if (k === "x") {
        console.log("[keydown] Ctrl+X", { k: k, hasTextSelection: hasTextSelection, inTable: !!inTable });
        e.preventDefault();
        doCut();
        return;
      }
      if (k === "v") { e.preventDefault(); doPaste(); return; }
    }

    if ((e.ctrlKey || e.metaKey) && (e.key === "a" || e.key === "A")) {
      e.preventDefault();
      selectAll();
      return;
    }
    var isUndo = (e.ctrlKey || e.metaKey) && !e.shiftKey && (e.key === "z" || e.key === "Z");
    var isRedo = (e.ctrlKey && e.key === "y") || (e.ctrlKey && e.shiftKey && (e.key === "z" || e.key === "Z")) || (e.metaKey && e.shiftKey && (e.key === "z" || e.key === "Z"));
    if (isUndo || isRedo) {
      if (e.isComposing) return;
      if (inTable && state.editMode && hasTextSelection) return;
      e.preventDefault();
      if (isUndo) { undo(); return; }
      if (isRedo) { redo(); return; }
    }
    if (e.key === "Escape") {
      e.preventDefault();
      clearCutStateIfEditingIntent(e);
      if (state.editMode) {
        state.editMode = false;
        var inp = getActiveInput();
        if (inp) { inp.setAttribute("readonly", "readonly"); inp.blur(); }
        updateSelectionUI(); // exitEditMode：不改變選取
      }
      return;
    }
    if (e.key === "Enter" && state.editMode) {
      e.preventDefault();
      clearCutStateIfEditingIntent(e);
      var inp = getActiveInput();
      if (inp) updateCell(state.activeSheet, r, c, inp.value);
      state.editMode = false;
      updateSelectionUI();
      setActiveCell(Math.min(r + 1, bounds.rowCount - 1), c, { editMode: false });
      return;
    }
    if (e.key === "Tab") {
      e.preventDefault();
      var nr = r, nc = c + 1;
      if (nc >= bounds.colCount) { nr = r + 1; nc = 0; }
      if (nr >= bounds.rowCount) nr = 0;
      setActiveCell(nr, nc, { editMode: false });
      return;
    }
    if (e.key === "Enter") {
      e.preventDefault();
      var nr = Math.min(r + 1, bounds.rowCount - 1);
      setActiveCell(nr, c, { editMode: false });
      return;
    }
    if (e.key === "ArrowUp") {
      e.preventDefault();
      setActiveCell(Math.max(0, r - 1), c, { editMode: false });
      return;
    }
    if (e.key === "ArrowDown") {
      e.preventDefault();
      setActiveCell(Math.min(bounds.rowCount - 1, r + 1), c, { editMode: false });
      return;
    }
    if (e.key === "ArrowLeft") {
      if (state.editMode && ae && (ae.selectionStart || 0) > 0) return;
      e.preventDefault();
      setActiveCell(r, Math.max(0, c - 1), { editMode: false });
      return;
    }
    if (e.key === "ArrowRight") {
      if (state.editMode && ae && (ae.selectionStart || 0) < (ae.value || "").length) return;
      e.preventDefault();
      setActiveCell(r, Math.min(bounds.colCount - 1, c + 1), { editMode: false });
      return;
    }
    if (e.key === "F2" && !state.editMode) {
      e.preventDefault();
      var inp = getActiveInput();
      if (inp) {
        state.editMode = true; // enterEditMode
        inp.removeAttribute("readonly");
        inp.focus();
        inp.setSelectionRange(inp.value.length, inp.value.length);
        updateSelectionUI();
      }
      return;
    }
    // Select 模式：可輸入字元→清空 active 並輸入首字元進 Edit；Delete/Backspace→清空所有選取格；此為「會改變內容」→ clear cut
    if (!state.editMode && (isPrintableKey(e) || e.key === "Backspace" || e.key === "Delete")) {
      e.preventDefault();
      clearCutStateIfEditingIntent(e);
      var inp = getActiveInput();
      if (!inp) return;
      if (e.key === "Backspace" || e.key === "Delete") {
        beginAction("Clear");
        getSelectedCells().forEach(function (cell) {
          updateCell(state.activeSheet, cell.row, cell.col, "");
        });
        commitAction();
      } else {
        state.editMode = true;
        inp.removeAttribute("readonly");
        inp.focus();
        updateCell(state.activeSheet, r, c, e.key);
        inp.setSelectionRange(1, 1);
        updateSelectionUI();
      }
    }
  });
}

// --- Run ---

if (typeof XLSX !== "undefined") {
  init();
} else {
  document.addEventListener("DOMContentLoaded", function () {
    if (typeof XLSX !== "undefined") init();
  });
}
