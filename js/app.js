/**
 * iPVMS Form - Application Logic
 * UI table data = Source of Truth. Download rebuilds from headers + data arrays.
 */

const state = {
  studentId: null,
  activeGroup: "ModelData",
  activeSheet: "Company",
  activePeriod: null, // YYYY-MM 格式，Period Data 模式時使用
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

// --- Period Management helpers ---

function isValidPeriodFormat(period) {
  // 驗證格式 YYYY-MM
  var match = String(period).match(/^(\d{4})-(\d{2})$/);
  if (!match) return false;
  var year = parseInt(match[1], 10);
  var month = parseInt(match[2], 10);
  return year >= 2017 && year <= 2037 && month >= 1 && month <= 12;
}

function getAllPeriodsFromStorage() {
  // 從 localStorage 讀取所有已建立的月份
  if (!state.studentId) return [];
  var key = "excelForm_v1_" + state.studentId;
  var saved = localStorage.getItem(key);
  if (!saved) return [];
  try {
    var parsed = JSON.parse(saved);
    if (parsed.periods && typeof parsed.periods === "object") {
      return Object.keys(parsed.periods).sort();
    }
  } catch (e) {
    console.error("Error reading periods:", e);
  }
  return [];
}

function getCurrentPeriodData() {
  // 取得當前月份的資料（如果 activeGroup 是 PeriodData）
  if (state.activeGroup !== "PeriodData" || !state.activePeriod) return null;
  if (!state.studentId) return null;
  var key = "excelForm_v1_" + state.studentId;
  var saved = localStorage.getItem(key);
  if (!saved) return null;
  try {
    var parsed = JSON.parse(saved);
    if (parsed.periods && parsed.periods[state.activePeriod]) {
      return parsed.periods[state.activePeriod];
    }
  } catch (e) {
    console.error("Error reading period data:", e);
  }
  return null;
}

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
  const isTableMapping = (state.activeSheet === "TableMapping");
  inputs.forEach(function (inp) {
    inp.classList.remove("cell-active", "cell-selected", "cell-cut", "cell-editing");
    var isActive = state.activeCell && inp.dataset.row === String(state.activeCell.row) && inp.dataset.col === String(state.activeCell.col);
    // TableMapping 永遠唯讀
    if (isTableMapping) {
      inp.setAttribute("readonly", "readonly");
      if (isActive) inp.classList.add("cell-active");
    } else if (isActive && state.editMode) {
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

    // 解析剪貼簿內容為 2D 陣列
    var rows = text.split(/\r?\n/);
    var clipboardData = rows.map(function (line) { return line.split("\t"); });
    // 移除最後的空行（如果有的話）
    if (clipboardData.length > 0 && clipboardData[clipboardData.length - 1].length === 1 && clipboardData[clipboardData.length - 1][0] === "") {
      clipboardData.pop();
    }
    if (clipboardData.length === 0) return;

    var clipboardRows = clipboardData.length;
    var clipboardCols = clipboardData.reduce(function (max, row) { return Math.max(max, row.length); }, 0);
    var isSingleCell = (clipboardRows === 1 && clipboardCols === 1);

    beginAction("Paste");
    var colCount = sheet.headers.length;
    var addedRows = 0;

    // 決定目標範圍
    var targetRect = null;
    var targetCells = null; // 用於 multiSelection

    // 優先處理 multiSelection（Ctrl+點選的不連續多格）
    if (state.multiSelection && state.multiSelection.length > 0) {
      if (isSingleCell) {
        // 剪貼簿是 1x1，填滿所有 multiSelection 的格
        targetCells = state.multiSelection;
        var singleValue = (clipboardData[0] && clipboardData[0][0] != null) ? String(clipboardData[0][0]) : "";
        targetCells.forEach(function (cell) {
          if (cell.row >= 0 && cell.col >= 0 && cell.col < colCount) {
            while (sheet.data.length <= cell.row) {
              sheet.data.push(Array(colCount).fill(""));
              addedRows++;
            }
            updateCell(state.activeSheet, cell.row, cell.col, singleValue);
          }
        });
      } else {
        // 剪貼簿不是 1x1，以 activeCell 為起點貼上（multiSelection 不支援平鋪）
        targetRect = {
          startRow: state.activeCell.row,
          startCol: state.activeCell.col,
          endRow: state.activeCell.row + clipboardRows - 1,
          endCol: state.activeCell.col + clipboardCols - 1
        };
      }
    } else if (state.selection) {
      // 有矩形 selection，使用 selection 作為目標範圍
      targetRect = {
        startRow: state.selection.startRow,
        startCol: state.selection.startCol,
        endRow: state.selection.endRow,
        endCol: state.selection.endCol
      };
    } else {
      // 只有 activeCell，以 activeCell 為起點，大小 = clipboard 大小
      targetRect = {
        startRow: state.activeCell.row,
        startCol: state.activeCell.col,
        endRow: state.activeCell.row + clipboardRows - 1,
        endCol: state.activeCell.col + clipboardCols - 1
      };
    }

    // 處理矩形範圍的貼上（非 multiSelection 或 multiSelection 但剪貼簿不是 1x1）
    if (targetRect) {
      var targetRows = targetRect.endRow - targetRect.startRow + 1;
      var targetCols = targetRect.endCol - targetRect.startCol + 1;

      // 確保有足夠的列
      while (sheet.data.length <= targetRect.endRow) {
        sheet.data.push(Array(colCount).fill(""));
        addedRows++;
      }

      // 決定貼上策略
      if (isSingleCell) {
        // 剪貼簿是 1x1，填滿整個 targetRect
        var singleValue = (clipboardData[0] && clipboardData[0][0] != null) ? String(clipboardData[0][0]) : "";
        for (var r = targetRect.startRow; r <= targetRect.endRow; r++) {
          for (var c = targetRect.startCol; c <= targetRect.endCol; c++) {
            if (c >= 0 && c < colCount) {
              updateCell(state.activeSheet, r, c, singleValue);
            }
          }
        }
      } else {
        // 剪貼簿是矩形區塊，需要平鋪
        for (var r = targetRect.startRow; r <= targetRect.endRow; r++) {
          for (var c = targetRect.startCol; c <= targetRect.endCol; c++) {
            if (c >= 0 && c < colCount) {
              // 使用 modulo 計算對應的剪貼簿位置（平鋪）
              var srcRow = (r - targetRect.startRow) % clipboardRows;
              var srcCol = (c - targetRect.startCol) % clipboardCols;
              var val = (clipboardData[srcRow] && clipboardData[srcRow][srcCol] != null) ? String(clipboardData[srcRow][srcCol]) : "";
              updateCell(state.activeSheet, r, c, val);
            }
          }
        }
      }
    }

    // hook: 只有「貼上成功」且 lastClipboardOp==="cut" 且 cutCells 存在時，才清空剪下來源（並避免與貼上目標重疊）
    if (state.lastClipboardOp === "cut" && state.cutCells && state.cutCells.sheet && state.cutCells.cells) {
      var pasteEndRow, pasteEndCol;
      if (targetCells) {
        // multiSelection 且 1x1：計算實際貼上的範圍
        var minRow = Math.min.apply(null, targetCells.map(function (c) { return c.row; }));
        var maxRow = Math.max.apply(null, targetCells.map(function (c) { return c.row; }));
        var minCol = Math.min.apply(null, targetCells.map(function (c) { return c.col; }));
        var maxCol = Math.max.apply(null, targetCells.map(function (c) { return c.col; }));
        pasteEndRow = maxRow;
        pasteEndCol = maxCol;
        var pasteStartRow = minRow;
        var pasteStartCol = minCol;
        console.log("[doPaste] clearing cut sources (multiSelection), count=", state.cutCells.cells.length, "avoid overlap", pasteStartRow, pasteStartCol, pasteEndRow, pasteEndCol);
        state.cutCells.cells.forEach(function (c) {
          if (state.cutCells.sheet !== state.activeSheet) {
            updateCell(state.cutCells.sheet, c.row, c.col, "");
          } else if (c.row < pasteStartRow || c.row > pasteEndRow || c.col < pasteStartCol || c.col > pasteEndCol) {
            updateCell(state.cutCells.sheet, c.row, c.col, "");
          }
        });
      } else if (targetRect) {
        pasteEndRow = targetRect.endRow;
        pasteEndCol = targetRect.endCol;
        var pasteStartRow = targetRect.startRow;
        var pasteStartCol = targetRect.startCol;
        console.log("[doPaste] clearing cut sources, count=", state.cutCells.cells.length, "avoid overlap", pasteStartRow, pasteStartCol, pasteEndRow, pasteEndCol);
        state.cutCells.cells.forEach(function (c) {
          if (state.cutCells.sheet !== state.activeSheet) {
            updateCell(state.cutCells.sheet, c.row, c.col, "");
          } else if (c.row < pasteStartRow || c.row > pasteEndRow || c.col < pasteStartCol || c.col > pasteEndCol) {
            updateCell(state.cutCells.sheet, c.row, c.col, "");
          }
        });
      }
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
      
      // 向下相容：如果沒有 periods，將舊資料遷移到預設月份
      if (!parsed.periods || typeof parsed.periods !== "object") {
        // 舊版本資料，遷移到 2024-01
        var defaultPeriod = "2024-01";
        parsed.periods = {};
        if (parsed.data && Object.keys(parsed.data).length > 0) {
          // 檢查是否為 PeriodData
          var hasPeriodSheets = false;
          for (var sheetName in parsed.data) {
            var config = getSheetConfig(sheetName);
            if (config && config.workbook === "PeriodData") {
              hasPeriodSheets = true;
              break;
            }
          }
          if (hasPeriodSheets) {
            parsed.periods[defaultPeriod] = {
              data: parsed.data || {},
              changeLog: parsed.changeLog || [],
              activeSheet: parsed.activeSheet || "Exchange Rate"
            };
            // 清除舊的 data 和 changeLog（已遷移到 periods）
            parsed.data = {};
            parsed.changeLog = [];
            showStatus("Data structure upgraded. Period data migrated to " + defaultPeriod);
          }
        }
        // 儲存遷移後的資料
        localStorage.setItem(key, JSON.stringify(parsed));
      }
      
      // 載入當前 activeGroup 的資料
      state.activeGroup = parsed.activeGroup || "ModelData";
      
      if (state.activeGroup === "PeriodData") {
        // PeriodData 模式：從 periods 載入
        state.activePeriod = parsed.activePeriod || null;
        if (state.activePeriod && parsed.periods[state.activePeriod]) {
          var periodData = parsed.periods[state.activePeriod];
          state.data = periodData.data || {};
          state.changeLog = periodData.changeLog || [];
          state.activeSheet = periodData.activeSheet || "Exchange Rate";
        } else {
          // 沒有選擇月份或月份不存在，使用第一個可用月份或初始化
          var availablePeriods = Object.keys(parsed.periods || {}).sort();
          if (availablePeriods.length > 0) {
            state.activePeriod = availablePeriods[0];
            var periodData = parsed.periods[state.activePeriod];
            state.data = periodData.data || {};
            state.changeLog = periodData.changeLog || [];
            state.activeSheet = periodData.activeSheet || "Exchange Rate";
          } else {
            // 沒有任何月份，初始化
            state.activePeriod = null;
            initFromTemplate();
            state.activeSheet = "Exchange Rate";
          }
        }
      } else {
        // ModelData 模式：直接從根層級載入（向下相容）
        state.data = parsed.data || {};
        state.changeLog = parsed.changeLog || [];
        state.activePeriod = null;
        const order = getSheetsForWorkbook(state.activeGroup);
        state.activeSheet = order.includes(parsed.activeSheet) ? parsed.activeSheet : (order[0] || "Company");
      }
      
      state.undoStack = [];
      state.redoStack = [];
      state._tx = null;
      ensureAllSheets();
      // TableMapping 永遠使用系統定義，不從 localStorage 讀取
      var tmHeaders = TABLE_MAPPING_HEADERS.map(function (h) { return h; });
      var tmData = TABLE_MAPPING_DATA.map(function (row) { return row.map(function (cell) { return cell; }); });
      state.data["TableMapping"] = { headers: tmHeaders, data: tmData };
      migrateMACHeaderNames();
      migrateSACHeaderNames();
      migrateACAPDescriptionColumn();
      migrateADriverDescriptionColumn();
      migrateADriverActivityDescriptionColumn();
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
    if (sheetName === "TableMapping") {
      // TableMapping 永遠使用系統定義（深拷貝）
      headers = TABLE_MAPPING_HEADERS.map(function (h) { return h; });
      data = TABLE_MAPPING_DATA.map(function (row) { return row.map(function (cell) { return cell; }); });
    } else if (config.workbook === "ModelData" && DEFAULT_ROWS_MODEL_MAP[sheetName] != null) {
      data = makeBlankRows(headers.length, DEFAULT_ROWS_MODEL_MAP[sheetName]);
    } else {
      data = config.data.map(function (row) { return [...row]; });
    }
    state.data[sheetName] = { headers: headers, data: data };
  }
  ensureMACHeaderRows();
  ensureSACHeaderRows();
  state.activeGroup = "ModelData";
  state.activeSheet = "Company";
}

function ensureMACHeaderRows() {
  var s = state.data["Resource Driver(M. A. C.)"];
  if (!s || !s.headers) return;
  var L = s.headers.length;
  if (!s.headers2) s.headers2 = [];
  while (s.headers2.length < L) s.headers2.push("");
  if (s.headers2.length > L) s.headers2.splice(L, s.headers2.length - L);
  if (!s.headers3) s.headers3 = [];
  while (s.headers3.length < L) s.headers3.push("");
  if (s.headers3.length > L) s.headers3.splice(L, s.headers3.length - L);
}

function ensureSACHeaderRows() {
  var s = state.data["Resource Driver(S. A. C.)"];
  if (!s || !s.headers) return;
  var L = s.headers.length;
  if (!s.headers2) s.headers2 = [];
  while (s.headers2.length < L) s.headers2.push("");
  if (s.headers2.length > L) s.headers2.splice(L, s.headers2.length - L);
  if (!s.headers3) s.headers3 = [];
  while (s.headers3.length < L) s.headers3.push("");
  if (s.headers3.length > L) s.headers3.splice(L, s.headers3.length - L);
}

// MAC header rename: migrate old B=Machine Code, C=Driver 1 → B=(Activity Center ), C=Machine Code
function migrateMACHeaderNames() {
  var s = state.data["Resource Driver(M. A. C.)"];
  if (!s || !s.headers || s.headers.length < 3) return;
  if (s.headers[1] === "Machine Code" && s.headers[2] === "Driver 1") {
    s.headers[1] = "(Activity Center )";
    s.headers[2] = "Machine Code";
    autoSave();
  }
}

// SAC header rename start: migrate old B=Machine Code, C=Driver 1 → B=(Activity Center ), C=Machine Code
function migrateSACHeaderNames() {
  var s = state.data["Resource Driver(S. A. C.)"];
  if (!s || !s.headers || s.headers.length < 3) return;
  if (s.headers[1] === "Machine Code" && s.headers[2] === "Driver 1") {
    s.headers[1] = "(Activity Center )";
    s.headers[2] = "Machine Code";
    autoSave();
  }
}
// SAC header rename end

// ACAP insert description column start: 舊資料補第二欄 (Activity Center description)
function migrateACAPDescriptionColumn() {
  var s = state.data["Activity Center Driver(A. Cap.)"];
  if (!s || !s.headers) return;
  if (s.headers[1] === "(Activity Center description)") return;
  s.headers.splice(1, 0, "(Activity Center description)");
  if (s.data) { for (var i = 0; i < s.data.length; i++) { s.data[i].splice(1, 0, ""); } }
  if (s.headers2) s.headers2.splice(1, 0, "");
  if (s.headers3) s.headers3.splice(1, 0, "");
  autoSave();
}
// ACAP insert description column end

// ADriver insert description column start: 舊資料補第二欄 (Activity Center description)
function migrateADriverDescriptionColumn() {
  var s = state.data["Activity Driver"];
  if (!s || !s.headers) return;
  if (s.headers[1] === "(Activity Center description)") return;
  s.headers.splice(1, 0, "(Activity Center description)");
  if (s.data) { for (var i = 0; i < s.data.length; i++) { s.data[i].splice(1, 0, ""); } }
  if (s.headers2) s.headers2.splice(1, 0, "");
  if (s.headers3) s.headers3.splice(1, 0, "");
  autoSave();
}
// ADriver insert description column end

// ADriver insert activity description column start: 舊資料補第 5 欄 (Activity  Description)
function migrateADriverActivityDescriptionColumn() {
  var s = state.data["Activity Driver"];
  if (!s || !s.headers) return;
  if (s.headers.indexOf("(Activity  Description)") !== -1) return;
  s.headers.splice(4, 0, "(Activity  Description)");
  if (s.data) { for (var i = 0; i < s.data.length; i++) { s.data[i].splice(4, 0, ""); } }
  if (s.headers2) s.headers2.splice(4, 0, "");
  if (s.headers3) s.headers3.splice(4, 0, "");
  autoSave();
}
// ADriver insert activity description column end

function ensureAllSheets() {
  for (const [sheetName, config] of Object.entries(TEMPLATE_DATA.sheets)) {
    if (sheetName === "TableMapping") {
      // TableMapping 永遠使用系統定義（深拷貝），不從 localStorage 讀取
      var headers = TABLE_MAPPING_HEADERS.map(function (h) { return h; });
      var data = TABLE_MAPPING_DATA.map(function (row) { return row.map(function (cell) { return cell; }); });
      state.data[sheetName] = { headers: headers, data: data };
      continue;
    }
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
  ensureMACHeaderRows();
  ensureSACHeaderRows();
  ensureServiceDriverPeriodColumns();
}

function ensureServiceDriverPeriodColumns() {
  var sheetName = "Service Driver_Period";
  var sheet = state.data[sheetName];
  if (!sheet) return;
  var config = getSheetConfig(sheetName);
  if (!config || !config.headers) return;
  var templateColCount = config.headers.length;
  var currentColCount = sheet.headers ? sheet.headers.length : 0;
  if (currentColCount < templateColCount) {
    // 舊資料欄位數不足，補上新的系統欄位
    if (!sheet.headers) sheet.headers = [];
    while (sheet.headers.length < templateColCount) {
      var templateIndex = sheet.headers.length;
      sheet.headers.push(config.headers[templateIndex] || "");
    }
    // 更新所有資料列的欄位數
    if (!sheet.data) sheet.data = [];
    for (var i = 0; i < sheet.data.length; i++) {
      if (!sheet.data[i]) sheet.data[i] = [];
      while (sheet.data[i].length < templateColCount) {
        sheet.data[i].push("");
      }
    }
    // 如果沒有資料列，建立一個空列
    if (sheet.data.length === 0) {
      sheet.data.push(Array(templateColCount).fill(""));
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

// --- Required Fields Admin Modal ---

function showRequiredFieldsModal() {
  var modal = document.getElementById("requiredFieldsModal");
  var sheetSelect = document.getElementById("requiredFieldsSheetSelect");
  var columnList = document.getElementById("requiredFieldsColumnList");
  
  // Populate sheet dropdown with PeriodData sheets only
  sheetSelect.innerHTML = '<option value="">-- Select Sheet --</option>';
  var periodSheets = getSheetsForWorkbook("PeriodData");
  periodSheets.forEach(function(sheetName) {
    var config = getSheetConfig(sheetName);
    if (config && !config.hidden) {
      var option = document.createElement("option");
      option.value = sheetName;
      option.textContent = getExcelSheetName(sheetName);
      sheetSelect.appendChild(option);
    }
  });
  
  columnList.innerHTML = '<p style="color: var(--color-text-muted); font-size: 13px;">Please select a sheet to configure required fields.</p>';
  modal.classList.remove("hidden");
}

function hideRequiredFieldsModal() {
  document.getElementById("requiredFieldsModal").classList.add("hidden");
}

function renderRequiredFieldsColumnList(sheetName) {
  var columnList = document.getElementById("requiredFieldsColumnList");
  var config = getSheetConfig(sheetName);
  if (!config || !config.headers) {
    columnList.innerHTML = '<p style="color: var(--color-text-muted); font-size: 13px;">No columns found for this sheet.</p>';
    return;
  }
  
  var override = getRequiredOverride();
  var sheetOverride = override[sheetName] || {};
  var defaultRequired = config.required || [];
  
  var html = '<div style="display: flex; flex-direction: column; gap: 8px;">';
  config.headers.forEach(function(headerName) {
    var normalizedCol = norm(headerName);
    var isOverrideSet = sheetOverride.hasOwnProperty(normalizedCol);
    var isRequiredValue = isOverrideSet ? sheetOverride[normalizedCol] : defaultRequired.some(function(r) { return norm(r) === normalizedCol; });
    
    var checkboxId = "req_" + sheetName + "_" + normalizedCol.replace(/[^a-zA-Z0-9]/g, "_");
    html += '<label style="display: flex; align-items: center; gap: 8px; padding: 8px; border: 1px solid var(--color-border); border-radius: 4px; cursor: pointer;">';
    html += '<input type="checkbox" id="' + checkboxId + '" data-sheet="' + sheetName + '" data-column="' + normalizedCol + '" ' + (isRequiredValue ? 'checked' : '') + '>';
    html += '<span>' + (headerName || "(empty)") + '</span>';
    if (isOverrideSet) {
      html += '<span style="color: var(--color-text-muted); font-size: 11px; margin-left: auto;">(override)</span>';
    }
    html += '</label>';
  });
  html += '</div>';
  columnList.innerHTML = html;
}

function saveRequiredFieldsOverride() {
  var sheetSelect = document.getElementById("requiredFieldsSheetSelect");
  var sheetName = sheetSelect.value;
  if (!sheetName) {
    showStatus("Please select a sheet first.", "error");
    return;
  }
  
  var override = getRequiredOverride();
  if (!override[sheetName]) {
    override[sheetName] = {};
  }
  
  var checkboxes = document.querySelectorAll("#requiredFieldsColumnList input[type='checkbox']");
  checkboxes.forEach(function(cb) {
    if (cb.dataset.sheet === sheetName) {
      var normalizedCol = cb.dataset.column;
      var isChecked = cb.checked;
      override[sheetName][normalizedCol] = isChecked;
    }
  });
  
  setRequiredOverride(override);
  showStatus("Required fields configuration saved.", "success");
  renderTable(); // Refresh table to show updated stars
  hideRequiredFieldsModal();
}

function resetRequiredFieldsOverride() {
  var sheetSelect = document.getElementById("requiredFieldsSheetSelect");
  var sheetName = sheetSelect.value;
  if (!sheetName) {
    showStatus("Please select a sheet first.", "error");
    return;
  }
  
  var override = getRequiredOverride();
  if (override[sheetName]) {
    delete override[sheetName];
    setRequiredOverride(override);
    showStatus("Required fields reset to default for " + getExcelSheetName(sheetName) + ".", "success");
    renderRequiredFieldsColumnList(sheetName);
    renderTable(); // Refresh table to show updated stars
  } else {
    showStatus("No override found for this sheet.", "error");
  }
}

// --- Optional Tabs Admin Modal ---

function getPeriodTabDimPrefs(studentId) {
  if (!studentId) {
    // No studentId, return default set
    return new Set(PERIOD_DIM_SHEETS);
  }
  
  var key = "excelForm_v1_" + studentId;
  try {
    var saved = localStorage.getItem(key);
    if (!saved) {
      return new Set(PERIOD_DIM_SHEETS);
    }
    
    var parsed = JSON.parse(saved);
    if (parsed.uiPrefs && parsed.uiPrefs.periodTabDim) {
      var dimArray = parsed.uiPrefs.periodTabDim;
      if (Array.isArray(dimArray)) {
        return new Set(dimArray);
      }
    }
    
    // No saved prefs, return default
    return new Set(PERIOD_DIM_SHEETS);
  } catch (e) {
    console.error("Error reading period tab dim prefs:", e);
    return new Set(PERIOD_DIM_SHEETS);
  }
}

function getOptionalTabsSetForUI() {
  if (state.activeGroup !== "PeriodData") return new Set();
  return getPeriodTabDimPrefs(state.studentId);
}

function setPeriodTabDimPrefs(studentId, dimArray) {
  if (!studentId) {
    showStatus("Please enter company name first", "error");
    return;
  }
  
  var key = "excelForm_v1_" + studentId;
  
  // Read existing data
  var existingData = {};
  try {
    var existing = localStorage.getItem(key);
    if (existing) {
      existingData = JSON.parse(existing);
    }
  } catch (e) {
    console.error("Error reading existing data:", e);
    existingData = {};
  }
  
  // Initialize uiPrefs if needed
  if (!existingData.uiPrefs) {
    existingData.uiPrefs = {};
  }
  
  // Update periodTabDim
  existingData.uiPrefs.periodTabDim = dimArray;
  existingData.lastModified = new Date().toISOString();
  
  // Save back
  try {
    localStorage.setItem(key, JSON.stringify(existingData));
  } catch (e) {
    console.error("Error saving period tab dim prefs:", e);
    showStatus("Storage error. Please try again.", "error");
  }
}

function showOptionalTabsModal() {
  var modal = document.getElementById("optionalTabsModal");
  if (!modal) return;
  renderOptionalTabsList();
  modal.classList.remove("hidden");
}

function hideOptionalTabsModal() {
  var modal = document.getElementById("optionalTabsModal");
  if (modal) {
    modal.classList.add("hidden");
  }
}

function renderOptionalTabsList() {
  var listContainer = document.getElementById("optionalTabsList");
  if (!listContainer) return;
  
  var periodSheets = getSheetsForWorkbook("PeriodData");
  var prefSet = getPeriodTabDimPrefs(state.studentId);
  
  var html = '<div style="display: flex; flex-direction: column; gap: 8px;">';
  periodSheets.forEach(function(sheetName) {
    var config = getSheetConfig(sheetName);
    if (!config || config.hidden) return;
    
    var isDim = prefSet.has(sheetName);
    var checkboxId = "optional_tab_" + sheetName.replace(/[^a-zA-Z0-9]/g, "_");
    var itemClass = isDim ? "dropdown-item-selected" : "dropdown-item-unselected";
    html += '<label class="' + itemClass + '" style="display: flex; align-items: center; gap: 8px; padding: 8px; border: 1px solid var(--color-border); border-radius: 4px; cursor: pointer;">';
    html += '<input type="checkbox" id="' + checkboxId + '" data-sheet="' + sheetName + '" ' + (isDim ? 'checked' : '') + '>';
    html += '<span>' + getExcelSheetName(sheetName) + '</span>';
    html += '</label>';
  });
  html += '</div>';
  listContainer.innerHTML = html;
  
  // Add event listeners to update styling when checkbox state changes
  var checkboxes = listContainer.querySelectorAll('input[type="checkbox"]');
  checkboxes.forEach(function(checkbox) {
    checkbox.addEventListener('change', function() {
      var label = this.closest('label');
      if (label) {
        if (this.checked) {
          label.classList.remove('dropdown-item-unselected');
          label.classList.add('dropdown-item-selected');
        } else {
          label.classList.remove('dropdown-item-selected');
          label.classList.add('dropdown-item-unselected');
        }
      }
    });
  });
}

function saveOptionalTabsFromUI() {
  if (!state.studentId) {
    showStatus("Please enter company name first", "error");
    return;
  }
  
  var checkboxes = document.querySelectorAll("#optionalTabsList input[type='checkbox']");
  var dimArray = [];
  
  checkboxes.forEach(function(cb) {
    if (cb.checked && cb.dataset.sheet) {
      dimArray.push(cb.dataset.sheet);
    }
  });
  
  setPeriodTabDimPrefs(state.studentId, dimArray);
  showStatus("Optional Tabs saved.", "success");
  // Re-render to reflect saved state (in case of any sync issues)
  renderOptionalTabsList();
  // Update nav pills styling immediately
  renderGroupedNav();
}

// --- Storage ---

function saveToStorage() {
  if (!state.studentId) return;
  const key = "excelForm_v1_" + state.studentId;
  
  // 讀取現有資料（保留 periods 和其他欄位）
  var existingData = {};
  try {
    var existing = localStorage.getItem(key);
    if (existing) {
      existingData = JSON.parse(existing);
    }
  } catch (e) {
    console.error("Error reading existing data:", e);
  }
  
  // 根據 activeGroup 決定儲存方式
  if (state.activeGroup === "PeriodData" && state.activePeriod) {
    // PeriodData 模式：儲存到 periods[activePeriod]
    if (!existingData.periods) existingData.periods = {};
    existingData.periods[state.activePeriod] = {
      data: state.data,
      changeLog: state.changeLog,
      activeSheet: state.activeSheet
    };
    existingData.activePeriod = state.activePeriod;
    existingData.activeGroup = state.activeGroup;
    existingData.studentId = state.studentId;
    existingData.lastModified = new Date().toISOString();
    existingData.version = existingData.version || 1;
  } else {
    // ModelData 模式：儲存到根層級（向下相容）
    existingData.version = existingData.version || 1;
    existingData.studentId = state.studentId;
    existingData.lastModified = new Date().toISOString();
    existingData.activeGroup = state.activeGroup;
    existingData.activeSheet = state.activeSheet;
    existingData.data = state.data;
    existingData.changeLog = state.changeLog;
    // 保留 periods（如果存在）
    if (!existingData.periods) existingData.periods = {};
  }
  
  try {
    localStorage.setItem(key, JSON.stringify(existingData));
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
    if (ch.type === "insert_col") {
      removeColumn(ch.sheet, ch.colIndex);
      last = { row: 0, col: Math.min(ch.colIndex, (state.data[ch.sheet] && state.data[ch.sheet].headers ? state.data[ch.sheet].headers.length - 1 : 0)) };
    } else if (ch.type === "delete_col") {
      var s = state.data[ch.sheet];
      insertColumnAt(ch.sheet, ch.colIndex, ch.deletedHeaderName);
      if (s) {
        if (s.headers2) s.headers2[ch.colIndex] = ch.deletedHeader2 || "";
        if (s.headers3) s.headers3[ch.colIndex] = ch.deletedHeader3 || "";
        for (var vi = 0; vi < (ch.deletedValues || []).length && vi < (s.data || []).length; vi++) s.data[vi][ch.colIndex] = ch.deletedValues[vi];
        ensureUserAddedColIds(s);
        if (s.userAddedColIds && ch.userAddedId) s.userAddedColIds[ch.colIndex] = ch.userAddedId;
      }
      last = { row: 0, col: ch.colIndex };
    } else if (ch.type === "rename_col") {
      var s = state.data[ch.sheet];
      if (s && s.headers) s.headers[ch.colIndex] = ch.oldName;
      last = { row: 0, col: ch.colIndex };
    } else {
      updateCell(ch.sheet, ch.row, ch.col, ch.oldValue, { skipLog: true, skipUndo: true });
      last = ch;
    }
  }
  autoSave();
  renderTable();
  if (last && last.row != null && last.col != null) state.activeCell = { row: last.row, col: last.col };
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
    if (ch.type === "insert_col") {
      insertColumnAt(ch.sheet, ch.colIndex, ch.insertedHeaderName);
      var s = state.data[ch.sheet];
      if (s && ch.userAddedId) { ensureUserAddedColIds(s); s.userAddedColIds[ch.colIndex] = ch.userAddedId; }
      last = { row: 0, col: ch.colIndex };
    } else if (ch.type === "delete_col") {
      removeColumn(ch.sheet, ch.colIndex);
      last = { row: 0, col: Math.min(ch.colIndex, (state.data[ch.sheet] && state.data[ch.sheet].headers ? state.data[ch.sheet].headers.length - 1 : 0)) };
    } else if (ch.type === "rename_col") {
      var s = state.data[ch.sheet];
      if (s && s.headers) s.headers[ch.colIndex] = ch.newName;
      last = { row: 0, col: ch.colIndex };
    } else {
      updateCell(ch.sheet, ch.row, ch.col, ch.newValue, { skipLog: true, skipUndo: true });
      last = ch;
    }
  }
  autoSave();
  renderTable();
  if (last && last.row != null && last.col != null) state.activeCell = { row: last.row, col: last.col };
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
  // TableMapping 是系統對照表，不允許修改
  if (sheetName === "TableMapping") {
    return;
  }
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
    showStatus("Enter your company name first", "error");
    return;
  }
  if (state.activeGroup === "PeriodData" && !state.activePeriod) {
    showStatus("Please select a period first", "error");
    return;
  }
  const timestamp = new Date().toISOString().slice(0, 16).replace(/[-:T]/g, "");
  const wbKey = state.activeGroup;
  var periodSuffix = "";
  if (wbKey === "PeriodData" && state.activePeriod) {
    periodSuffix = "_" + state.activePeriod.replace("-", "");
  }
  downloadWorkbook(wbKey, timestamp, periodSuffix);
  if (wbKey === "ModelData") {
    showStatus("Downloaded ModelData file");
  } else if (wbKey === "PeriodData") {
    showStatus("Downloaded PeriodData file for " + state.activePeriod);
  } else {
    showStatus("Downloaded " + wbKey + " file");
  }
}

function downloadWorkbook(workbookKey, timestamp, periodSuffix) {
  periodSuffix = periodSuffix || "";
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
    var wsData;
    if (internalName === "Resource Driver(M. A. C.)" && config.headerRows && sheet) {
      ensureMACHeaderRows();
      var row1 = sheet.headers;
      var row2 = []; for (var i = 0; i < sheet.headers.length; i++) row2.push(i < 3 ? "" : (sheet.headers2[i] || ""));
      var row3 = []; for (var i = 0; i < sheet.headers.length; i++) row3.push(i < 3 ? "" : (sheet.headers3[i] || ""));
      wsData = [row1, row2, row3].concat(sheet.data);
    } else if (internalName === "Resource Driver(S. A. C.)" && config.headerRows && sheet) {
      ensureSACHeaderRows();
      var row1 = sheet.headers;
      var row2 = []; for (var i = 0; i < sheet.headers.length; i++) row2.push(i < 3 ? "" : (sheet.headers2[i] || ""));
      var row3 = []; for (var i = 0; i < sheet.headers.length; i++) row3.push(i < 3 ? "" : (sheet.headers3[i] || ""));
      wsData = [row1, row2, row3].concat(sheet.data);
    } else {
      wsData = [headers].concat(data);
    }
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
      ["3", "資源動因(作業中心)", "Resource Driver(Actvity Center)", "Resource Driver(Actvity Center)", "Required", ""],
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

    // PeriodData：覆蓋 "TableMapping" 工作表為固定系統對照表（5 欄表頭 + 74 列資料），確保下載時永遠是正確內容
    // 使用系統定義的 TABLE_MAPPING_HEADERS 和 TABLE_MAPPING_DATA，保留所有尾部空格（不 trim）
    var tmAoA = [TABLE_MAPPING_HEADERS.map(function (h) { return h; })].concat(
      TABLE_MAPPING_DATA.map(function (row) {
        return row.map(function (cell) { return cell; }); // 深拷貝，保留尾部空格
      })
    );
    var tmWs = XLSX.utils.aoa_to_sheet(tmAoA);
    if (wb.SheetNames.indexOf("TableMapping") !== -1) {
      wb.Sheets["TableMapping"] = tmWs;
    } else {
      XLSX.utils.book_append_sheet(wb, tmWs, "TableMapping");
    }

    // PeriodData：指定分頁在輸出的 Excel 中預設為 Hidden（只影響 Period Download；不存在的 sheet 略過）
    var PERIOD_HIDDEN_SHEETS = new Set(["工作表2", "Sheet2", "Item", "TableMapping"]);
    wb.Workbook = wb.Workbook || {};
    wb.Workbook.Sheets = wb.SheetNames.map(function (name) {
      var o = {};
      if (PERIOD_HIDDEN_SHEETS.has(name)) o.Hidden = 1;
      return o;
    });
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

  const filename = workbookConfig.filename + periodSuffix + "_" + state.studentId + "_" + timestamp + ".xlsx";
  XLSX.writeFile(wb, filename);
}

// --- Upload (auto-detect workbook) ---

function uploadBackup(file) {
  if (!state.studentId) {
    showStatus("Enter your company name first", "error");
    return;
  }
  if (state.activeGroup === "PeriodData" && !state.activePeriod) {
    showStatus("Please select a period first", "error");
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
          if (internalName === "Resource Driver(M. A. C.)" && config.headerRows) {
            var colCount, finalH, finalH2, finalH3, nd;
            if (jsonData.length >= 3 && Array.isArray(jsonData[0]) && Array.isArray(jsonData[1]) && Array.isArray(jsonData[2])) {
              var r0 = jsonData[0], r1 = jsonData[1], r2 = jsonData[2];
              colCount = Math.max(r0.length, r1.length, r2.length, (templateHeaders || []).length);
              finalH = []; for (var i = 0; i < colCount; i++) finalH.push((r0[i] != null) ? String(r0[i]).trim() : "");
              finalH2 = []; for (var i = 0; i < colCount; i++) finalH2.push((r1[i] != null) ? String(r1[i]).trim() : "");
              finalH3 = []; for (var i = 0; i < colCount; i++) finalH3.push((r2[i] != null) ? String(r2[i]).trim() : "");
              nd = jsonData.slice(3).map(function (row) { var n = [].concat(row); while (n.length < colCount) n.push(""); return n.slice(0, colCount); });
            } else {
              var row0 = Array.isArray(jsonData[0]) ? jsonData[0] : [];
              colCount = Math.max(row0.length, (templateHeaders || []).length);
              finalH = []; for (var i = 0; i < colCount; i++) finalH.push((row0[i] != null) ? String(row0[i]).trim() : "");
              finalH2 = Array(colCount).fill(""); finalH3 = Array(colCount).fill("");
              nd = jsonData.slice(1).map(function (row) { var n = [].concat(row); while (n.length < colCount) n.push(""); return n.slice(0, colCount); });
            }
            state.data[internalName] = { headers: finalH, headers2: finalH2, headers3: finalH3, data: nd.length > 0 ? nd : [Array(colCount).fill("")] };
            return;
          }
          if (internalName === "Resource Driver(S. A. C.)" && config.headerRows) {
            var colCount, finalH, finalH2, finalH3, nd;
            if (jsonData.length >= 3 && Array.isArray(jsonData[0]) && Array.isArray(jsonData[1]) && Array.isArray(jsonData[2])) {
              var r0 = jsonData[0], r1 = jsonData[1], r2 = jsonData[2];
              colCount = Math.max(r0.length, r1.length, r2.length, (templateHeaders || []).length);
              finalH = []; for (var i = 0; i < colCount; i++) finalH.push((r0[i] != null) ? String(r0[i]).trim() : "");
              finalH2 = []; for (var i = 0; i < colCount; i++) finalH2.push((r1[i] != null) ? String(r1[i]).trim() : "");
              finalH3 = []; for (var i = 0; i < colCount; i++) finalH3.push((r2[i] != null) ? String(r2[i]).trim() : "");
              nd = jsonData.slice(3).map(function (row) { var n = [].concat(row); while (n.length < colCount) n.push(""); return n.slice(0, colCount); });
            } else {
              var row0 = Array.isArray(jsonData[0]) ? jsonData[0] : [];
              colCount = Math.max(row0.length, (templateHeaders || []).length);
              finalH = []; for (var i = 0; i < colCount; i++) finalH.push((row0[i] != null) ? String(row0[i]).trim() : "");
              finalH2 = Array(colCount).fill(""); finalH3 = Array(colCount).fill("");
              nd = jsonData.slice(1).map(function (row) { var n = [].concat(row); while (n.length < colCount) n.push(""); return n.slice(0, colCount); });
            }
            state.data[internalName] = { headers: finalH, headers2: finalH2, headers3: finalH3, data: nd.length > 0 ? nd : [Array(colCount).fill("")] };
            return;
          }

          var finalHeaders = templateHeaders;
          if (internalName === "Resource Driver(Actvity Center)") {
            // Resource Driver(Actvity Center): 保留所有上傳的欄位，以實際欄位結構為準
            var uploadHeaders = Array.isArray(jsonData[0]) ? jsonData[0] : [];
            var uploadColCount = uploadHeaders.length;
            var templateColCount = templateHeaders.length;
            // 確保至少保留模板的必要欄位（前2欄固定）
            var minColCount = Math.max(uploadColCount, templateColCount);
            finalHeaders = [];
            for (var i = 0; i < minColCount; i++) {
              if (i < 2) {
                // 前2欄使用模板（Activity Center Code, (Activity Center)）
                finalHeaders.push(templateHeaders[i] || "");
              } else if (i < uploadColCount) {
                // 使用上傳的欄位名（保留所有新增欄位）
                var uploadHeader = (typeof uploadHeaders[i] === "string" ? String(uploadHeaders[i]).trim() : "");
                finalHeaders.push(uploadHeader || "");
              } else {
                // 如果上傳欄位比模板少，補空字串（不應該發生，但以防萬一）
                finalHeaders.push("");
              }
            }
          }
          var colCount = finalHeaders.length;
          const uploadedData = jsonData.slice(1);
          const normalizedData = uploadedData.map(function (row) {
            const newRow = [].concat(row);
            while (newRow.length < colCount) newRow.push("");
            return newRow.slice(0, colCount);
          });

          state.data[internalName] = {
            headers: finalHeaders,
            data: normalizedData.length > 0 ? normalizedData : [Array(colCount).fill("")]
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
      // TableMapping 不論上傳內容為何，一律恢復為系統定義
      if (workbookKey === "PeriodData") {
        var tmHeaders = TABLE_MAPPING_HEADERS.map(function (h) { return h; });
        var tmData = TABLE_MAPPING_DATA.map(function (row) { return row.map(function (cell) { return cell; }); });
        state.data["TableMapping"] = { headers: tmHeaders, data: tmData };
      }
      migrateMACHeaderNames();
      migrateSACHeaderNames();
      migrateACAPDescriptionColumn();
      migrateADriverDescriptionColumn();
      migrateADriverActivityDescriptionColumn();
      // PeriodData 上傳後只套用到當前月份
      if (workbookKey === "PeriodData" && state.activePeriod) {
        // 資料已經寫入 state.data，現在儲存到當前月份
        saveToStorage();
        showStatus("Restored " + workbookKey + " backup for " + state.activePeriod + ". TableMapping restored to system default.");
      } else {
        saveToStorage();
        if (workbookKey === "PeriodData") {
          showStatus("Restored " + workbookKey + " backup. TableMapping restored to system default.");
        } else {
          showStatus("Restored " + workbookKey + " backup");
        }
      }
      renderAll();
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
  if (state.activeGroup === "PeriodData" && state.activePeriod) {
    // PeriodData 模式：只重置當前月份
    initFromTemplate();
    // 只保留 PeriodData 的 sheets
    var periodSheets = {};
    for (var sheetName in state.data) {
      var config = getSheetConfig(sheetName);
      if (config && config.workbook === "PeriodData") {
        periodSheets[sheetName] = state.data[sheetName];
      }
    }
    state.data = periodSheets;
    state.changeLog = [];
    state.undoStack = [];
    state.redoStack = [];
    state._tx = null;
    ensureAllSheets();
    saveToStorage();
    renderAll();
    showStatus("Period " + state.activePeriod + " reset to template");
  } else {
    // ModelData 模式：重置所有
    initFromTemplate();
    saveToStorage();
    renderAll();
    showStatus("Data reset to template");
  }
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
  // 顯示/隱藏 period selector
  var periodContainer = document.getElementById("periodSelectorContainer");
  if (periodContainer) {
    periodContainer.style.display = state.activeGroup === "PeriodData" ? "flex" : "none";
  }
  // 顯示/隱藏 Required Fields 按鈕（僅在 PeriodData 模式下顯示）
  var btnRequiredFields = document.getElementById("btnRequiredFields");
  if (btnRequiredFields) {
    btnRequiredFields.style.display = state.activeGroup === "PeriodData" ? "" : "none";
  }
  // 顯示/隱藏 Optional Tabs 按鈕（僅在 PeriodData 模式下顯示）
  var btnOptionalTabs = document.getElementById("btnOptionalTabs");
  if (btnOptionalTabs) {
    btnOptionalTabs.style.display = state.activeGroup === "PeriodData" ? "" : "none";
  }
  // 更新 period selector 選項
  if (state.activeGroup === "PeriodData") {
    updatePeriodSelector();
  }
}

function updatePeriodSelector() {
  var select = document.getElementById("periodSelect");
  if (!select) return;
  var periods = getAllPeriodsFromStorage();
  select.innerHTML = '<option value="">-- Select Period --</option>';
  periods.forEach(function (period) {
    var option = document.createElement("option");
    option.value = period;
    option.textContent = period;
    if (period === state.activePeriod) {
      option.selected = true;
    }
    select.appendChild(option);
  });
}

function switchPeriod(period) {
  if (!isValidPeriodFormat(period)) {
    showStatus("Invalid period format. Must be YYYY-MM (2017-01 to 2037-12)", "error");
    return;
  }
  if (!state.studentId) {
    showStatus("Please enter company name first", "error");
    return;
  }
  
  // 儲存當前月份的資料（如果有）
  if (state.activePeriod) {
    saveToStorage();
  }
  
  // 切換到新月份
  state.activePeriod = period;
  var key = "excelForm_v1_" + state.studentId;
  var saved = localStorage.getItem(key);
  if (saved) {
    try {
      var parsed = JSON.parse(saved);
      if (!parsed.periods) parsed.periods = {};
      
      if (parsed.periods[period]) {
        // 月份已存在，載入資料
        var periodData = parsed.periods[period];
        state.data = periodData.data || {};
        state.changeLog = periodData.changeLog || [];
        state.activeSheet = periodData.activeSheet || "Exchange Rate";
      } else {
        // 月份不存在，初始化（使用 template）
        initFromTemplate();
        // 只保留 PeriodData 的 sheets
        var periodSheets = {};
        for (var sheetName in state.data) {
          var config = getSheetConfig(sheetName);
          if (config && config.workbook === "PeriodData") {
            periodSheets[sheetName] = state.data[sheetName];
          }
        }
        state.data = periodSheets;
        state.activeSheet = "Exchange Rate";
        ensureAllSheets();
      }
      
      // 更新 storage
      parsed.activePeriod = period;
      parsed.activeGroup = "PeriodData";
      localStorage.setItem(key, JSON.stringify(parsed));
      
      state.undoStack = [];
      state.redoStack = [];
      state._tx = null;
      clearCutState();
      renderAll();
      showStatus("Switched to period " + period);
    } catch (e) {
      console.error("Error switching period:", e);
      showStatus("Error switching period", "error");
    }
  }
}

function createNewPeriod() {
  if (!state.studentId) {
    showStatus("Please enter company name first", "error");
    return;
  }
  
  var period = prompt("Enter period (YYYY-MM, e.g. 2024-01):");
  if (!period) return;
  
  period = String(period).trim();
  if (!isValidPeriodFormat(period)) {
    showStatus("Invalid period format. Must be YYYY-MM (2017-01 to 2037-12)", "error");
    return;
  }
  
  var periods = getAllPeriodsFromStorage();
  if (periods.indexOf(period) >= 0) {
    showStatus("Period " + period + " already exists", "error");
    return;
  }
  
  // 儲存當前月份的資料（如果有）
  if (state.activePeriod) {
    saveToStorage();
  }
  
  // 建立新月份：使用 template（空白）或複製當前月份
  var key = "excelForm_v1_" + state.studentId;
  var saved = localStorage.getItem(key);
  var parsed = saved ? JSON.parse(saved) : {};
  if (!parsed.periods) parsed.periods = {};
  
  // 策略：如果當前有月份，複製當前月份；否則使用 template
  var newPeriodData;
  if (state.activePeriod && parsed.periods[state.activePeriod]) {
    // 複製當前月份
    var currentData = parsed.periods[state.activePeriod];
    newPeriodData = {
      data: JSON.parse(JSON.stringify(currentData.data || {})),
      changeLog: [],
      activeSheet: currentData.activeSheet || "Exchange Rate"
    };
  } else {
    // 使用 template
    initFromTemplate();
    var periodSheets = {};
    for (var sheetName in state.data) {
      var config = getSheetConfig(sheetName);
      if (config && config.workbook === "PeriodData") {
        periodSheets[sheetName] = JSON.parse(JSON.stringify(state.data[sheetName]));
      }
    }
    newPeriodData = {
      data: periodSheets,
      changeLog: [],
      activeSheet: "Exchange Rate"
    };
  }
  
  parsed.periods[period] = newPeriodData;
  parsed.activePeriod = period;
  parsed.activeGroup = "PeriodData";
  localStorage.setItem(key, JSON.stringify(parsed));
  
  // 切換到新月份
  state.activePeriod = period;
  state.data = newPeriodData.data;
  state.changeLog = [];
  state.activeSheet = newPeriodData.activeSheet;
  state.undoStack = [];
  state.redoStack = [];
  state._tx = null;
  clearCutState();
  ensureAllSheets();
  renderAll();
  showStatus("Created and switched to period " + period);
}

function renderGroupedNav() {
  const container = document.getElementById("groupedNav");
  if (!container) return;
  const wbKey = state.activeGroup;
  const wb = TEMPLATE_DATA.workbooks[wbKey];
  const groups = wb && wb.uiGroups;
  container.innerHTML = "";
  if (!groups || !groups.length) return;

  // Get optional tabs set for PeriodData
  var optionalSet = (wbKey === "PeriodData") ? getOptionalTabsSetForUI() : null;

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
      // Apply tab-dim class based on saved preferences for PeriodData
      if (wbKey === "PeriodData" && optionalSet) {
        pill.classList.toggle("tab-dim", optionalSet.has(internalName));
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

  // TableMapping 是系統對照表，唯讀
  const isTableMapping = (state.activeSheet === "TableMapping");
  if (isTableMapping) {
    document.getElementById("btnAddRow").style.display = "none";
  } else {
    document.getElementById("btnAddRow").style.display = "";
  }

  if (sheet.headers && sheet.headers.length > 0 && (!sheet.data || sheet.data.length === 0)) {
    sheet.data = [Array(sheet.headers.length).fill("")];
  }
  // 確保資料列的欄位數與 headers 匹配（安全檢查）
  if (sheet.headers && sheet.data) {
    var headerCount = sheet.headers.length;
    for (var i = 0; i < sheet.data.length; i++) {
      if (!sheet.data[i]) sheet.data[i] = [];
      while (sheet.data[i].length < headerCount) {
        sheet.data[i].push("");
      }
      if (sheet.data[i].length > headerCount) {
        sheet.data[i] = sheet.data[i].slice(0, headerCount);
      }
    }
  }

  const headers = sheet.headers;
  console.log("[headers]", state.activeSheet, sheet.headers);
  const isRDAC = (state.activeSheet === "Resource Driver(Actvity Center)");
  const isMAC = (state.activeSheet === "Resource Driver(M. A. C.)" && config.headerRows);
  const isSAC = (state.activeSheet === "Resource Driver(S. A. C.)" && config.headerRows);
  if (isRDAC) {
    thead.innerHTML = "";
    const tr = document.createElement("tr");
    tr.appendChild(createCell("th", "row-num", "#"));
    headers.forEach(function (h, colIndex) {
      const req = isRequired(state.activeSheet, h);
      const th = document.createElement("th");
      if (req) th.classList.add("required");
      if (colIndex >= 2) {
        const inp = document.createElement("input");
        inp.className = "th-input";
        inp.type = "text";
        inp.value = (h != null) ? String(h) : "";
        inp.addEventListener("input", function () {
          var s = state.data[state.activeSheet];
          if (s && s.headers) { s.headers[colIndex] = inp.value; autoSave(); }
        });
        if (state.activeGroup === "PeriodData") {
          if (isUserAddedColumn(state.activeSheet, colIndex)) {
            var headerFocusValue;
            inp.addEventListener("focus", function () { headerFocusValue = inp.value; });
            inp.addEventListener("blur", function () { if (inp.value !== headerFocusValue) renameColumn(state.activeSheet, colIndex, inp.value); });
          }
        }
        if (colIndex === 3) {
          const wrap = document.createElement("span");
          wrap.className = "th-dc2-wrap";
          wrap.appendChild(inp);
          if (req) {
            const star = document.createElement("span");
            star.className = "req-star";
            star.textContent = "*";
            wrap.appendChild(star);
          }
          const btn = document.createElement("button");
          btn.type = "button";
          btn.className = "btn-add-column";
          btn.textContent = "+";
          btn.title = "Add Driver Code";
          btn.addEventListener("click", function (e) {
            e.stopPropagation();
            addDriverCodeColumn();
          });
          wrap.appendChild(btn);
          th.appendChild(wrap);
        } else if (isUserAddedColumn(state.activeSheet, colIndex)) {
          const wrap = document.createElement("span");
          wrap.className = "th-dc2-wrap";
          wrap.appendChild(inp);
          if (req) {
            const star = document.createElement("span");
            star.className = "req-star";
            star.textContent = "*";
            wrap.appendChild(star);
          }
          const btn = document.createElement("button");
          btn.type = "button";
          btn.className = "btn-delete-column";
          btn.textContent = "\u00D7";
          btn.title = "Delete column";
          btn.addEventListener("click", function (e) {
            e.stopPropagation();
            deleteUserColumn(state.activeSheet, colIndex);
          });
          wrap.appendChild(btn);
          th.appendChild(wrap);
        } else {
          if (req) {
            const wrap = document.createElement("span");
            wrap.className = "th-header-wrap";
            wrap.appendChild(inp);
            const star = document.createElement("span");
            star.className = "req-star";
            star.textContent = "*";
            wrap.appendChild(star);
            th.appendChild(wrap);
          } else {
            th.appendChild(inp);
          }
        }
      } else {
        if (req) {
          const wrap = document.createElement("span");
          wrap.className = "th-header-wrap";
          wrap.textContent = (h != null) ? String(h) : "";
          const star = document.createElement("span");
          star.className = "req-star";
          star.textContent = "*";
          wrap.appendChild(star);
          th.appendChild(wrap);
        } else {
          th.textContent = (h != null) ? String(h) : "";
        }
      }
      if (state.activeGroup === "PeriodData") {
        th.addEventListener("contextmenu", function (e) { e.preventDefault(); insertColumnRight(state.activeSheet, colIndex); });
      }
      tr.appendChild(th);
    });
    tr.appendChild(createCell("th", "row-actions", ""));
    thead.appendChild(tr);
  } else if (isMAC) {
    ensureMACHeaderRows();
    thead.innerHTML = "";
    var h2 = sheet.headers2 || [];
    var h3 = sheet.headers3 || [];
    // tr1: # rowspan=3, A rowspan=3, B rowspan=3, C rowspan=3, D+ (headers, th-input, + on first D+, X on rest), row-actions rowspan=3
    var tr1 = document.createElement("tr");
    var thNum = document.createElement("th");
    thNum.className = "row-num";
    thNum.setAttribute("rowspan", "3");
    thNum.textContent = "#";
    tr1.appendChild(thNum);
    for (var c = 0; c < 3; c++) {
      var th = document.createElement("th");
      th.setAttribute("rowspan", "3");
      var req = isRequired(state.activeSheet, headers[c]);
      if (req) th.classList.add("required");
      if (req) {
        var wrap = document.createElement("span");
        wrap.className = "th-header-wrap";
        wrap.textContent = (headers[c] != null) ? String(headers[c]) : "";
        var star = document.createElement("span");
        star.className = "req-star";
        star.textContent = "*";
        wrap.appendChild(star);
        th.appendChild(wrap);
      } else {
        th.textContent = (headers[c] != null) ? String(headers[c]) : "";
      }
      if (state.activeGroup === "PeriodData") (function (col) { th.addEventListener("contextmenu", function (e) { e.preventDefault(); insertColumnRight(state.activeSheet, col); }); })(c);
      tr1.appendChild(th);
    }
    for (var c = 3; c < headers.length; c++) {
      var th = document.createElement("th");
      var req = isRequired(state.activeSheet, headers[c]);
      if (req) th.classList.add("required");
      var inp = document.createElement("input");
      inp.className = "th-input";
      inp.type = "text";
      inp.value = (headers[c] != null) ? String(headers[c]) : "";
      inp.addEventListener("input", function (idx) {
        return function () {
          var s = state.data[state.activeSheet];
          if (s && s.headers) { s.headers[idx] = inp.value; autoSave(); }
        };
      }(c));
      if (state.activeGroup === "PeriodData" && isUserAddedColumn(state.activeSheet, c)) {
        var headerFocusVal;
        inp.addEventListener("focus", function () { headerFocusVal = inp.value; });
        inp.addEventListener("blur", function () { if (inp.value !== headerFocusVal) renameColumn(state.activeSheet, c, inp.value); });
      }
      if (c === 3) {
        var wrap = document.createElement("span");
        wrap.className = "th-dc2-wrap";
        wrap.appendChild(inp);
        if (req) {
          var star = document.createElement("span");
          star.className = "req-star";
          star.textContent = "*";
          wrap.appendChild(star);
        }
        var btnAdd = document.createElement("button");
        btnAdd.type = "button";
        btnAdd.className = "btn-add-column";
        btnAdd.textContent = "+";
        btnAdd.title = "Add Driver column";
        btnAdd.addEventListener("click", function (e) { e.stopPropagation(); e.preventDefault(); addDriverCodeColumnMAC(); });
        wrap.appendChild(btnAdd);
        th.appendChild(wrap);
      } else if (isUserAddedColumn(state.activeSheet, c)) {
        var wrap = document.createElement("span");
        wrap.className = "th-dc2-wrap";
        wrap.appendChild(inp);
        if (req) {
          var star = document.createElement("span");
          star.className = "req-star";
          star.textContent = "*";
          wrap.appendChild(star);
        }
        var btnDel = document.createElement("button");
        btnDel.type = "button";
        btnDel.className = "btn-delete-column";
        btnDel.textContent = "\u00D7";
        btnDel.title = "Delete column";
        btnDel.addEventListener("click", (function (col) { return function (e) { e.stopPropagation(); e.preventDefault(); deleteUserColumn(state.activeSheet, col); }; })(c));
        wrap.appendChild(btnDel);
        th.appendChild(wrap);
      } else {
        if (req) {
          var wrap = document.createElement("span");
          wrap.className = "th-header-wrap";
          wrap.appendChild(inp);
          var star = document.createElement("span");
          star.className = "req-star";
          star.textContent = "*";
          wrap.appendChild(star);
          th.appendChild(wrap);
        } else {
          th.appendChild(inp);
        }
      }
      if (state.activeGroup === "PeriodData") (function (col) { th.addEventListener("contextmenu", function (e) { e.preventDefault(); insertColumnRight(state.activeSheet, col); }); })(c);
      tr1.appendChild(th);
    }
    var thAct = document.createElement("th");
    thAct.className = "row-actions";
    thAct.setAttribute("rowspan", "3");
    tr1.appendChild(thAct);
    thead.appendChild(tr1);
    // tr2: D+ only, headers2
    var tr2 = document.createElement("tr");
    for (var c = 3; c < headers.length; c++) {
      var th = document.createElement("th");
      var inp = document.createElement("input");
      inp.className = "th-input";
      inp.type = "text";
      inp.value = (h2[c] != null) ? String(h2[c]) : "";
      inp.addEventListener("input", function (idx) {
        return function () {
          var s = state.data[state.activeSheet];
          if (!s.headers2) s.headers2 = [];
          s.headers2[idx] = inp.value;
          autoSave();
        };
      }(c));
      th.appendChild(inp);
      tr2.appendChild(th);
    }
    thead.appendChild(tr2);
    // tr3: D+ only, headers3
    var tr3 = document.createElement("tr");
    for (var c = 3; c < headers.length; c++) {
      var th = document.createElement("th");
      var inp = document.createElement("input");
      inp.className = "th-input";
      inp.type = "text";
      inp.value = (h3[c] != null) ? String(h3[c]) : "";
      inp.addEventListener("input", function (idx) {
        return function () {
          var s = state.data[state.activeSheet];
          if (!s.headers3) s.headers3 = [];
          s.headers3[idx] = inp.value;
          autoSave();
        };
      }(c));
      th.appendChild(inp);
      tr3.appendChild(th);
    }
    thead.appendChild(tr3);
  } else if (isSAC) {
    ensureSACHeaderRows();
    thead.innerHTML = "";
    var h2 = sheet.headers2 || [];
    var h3 = sheet.headers3 || [];
    var tr1 = document.createElement("tr");
    var thNum = document.createElement("th");
    thNum.className = "row-num";
    thNum.setAttribute("rowspan", "3");
    thNum.textContent = "#";
    tr1.appendChild(thNum);
    for (var c = 0; c < 3; c++) {
      var th = document.createElement("th");
      th.setAttribute("rowspan", "3");
      var req = isRequired(state.activeSheet, headers[c]);
      if (req) th.classList.add("required");
      if (req) {
        var wrap = document.createElement("span");
        wrap.className = "th-header-wrap";
        wrap.textContent = (headers[c] != null) ? String(headers[c]) : "";
        var star = document.createElement("span");
        star.className = "req-star";
        star.textContent = "*";
        wrap.appendChild(star);
        th.appendChild(wrap);
      } else {
        th.textContent = (headers[c] != null) ? String(headers[c]) : "";
      }
      if (state.activeGroup === "PeriodData") (function (col) { th.addEventListener("contextmenu", function (e) { e.preventDefault(); insertColumnRight(state.activeSheet, col); }); })(c);
      tr1.appendChild(th);
    }
    for (var c = 3; c < headers.length; c++) {
      var th = document.createElement("th");
      var req = isRequired(state.activeSheet, headers[c]);
      if (req) th.classList.add("required");
      var inp = document.createElement("input");
      inp.className = "th-input";
      inp.type = "text";
      inp.value = (headers[c] != null) ? String(headers[c]) : "";
      inp.addEventListener("input", function (idx) {
        return function () {
          var s = state.data[state.activeSheet];
          if (s && s.headers) { s.headers[idx] = inp.value; autoSave(); }
        };
      }(c));
      if (state.activeGroup === "PeriodData" && isUserAddedColumn(state.activeSheet, c)) {
        var headerFocusValSAC;
        inp.addEventListener("focus", function () { headerFocusValSAC = inp.value; });
        inp.addEventListener("blur", function () { if (inp.value !== headerFocusValSAC) renameColumn(state.activeSheet, c, inp.value); });
      }
      if (c === 3) {
        var wrap = document.createElement("span");
        wrap.className = "th-dc2-wrap";
        wrap.appendChild(inp);
        if (req) {
          var star = document.createElement("span");
          star.className = "req-star";
          star.textContent = "*";
          wrap.appendChild(star);
        }
        var btnAdd = document.createElement("button");
        btnAdd.type = "button";
        btnAdd.className = "btn-add-column";
        btnAdd.textContent = "+";
        btnAdd.title = "Add Driver column";
        btnAdd.addEventListener("click", function (e) { e.stopPropagation(); e.preventDefault(); addDriverCodeColumnSAC(); });
        wrap.appendChild(btnAdd);
        th.appendChild(wrap);
      } else if (isUserAddedColumn(state.activeSheet, c)) {
        var wrap = document.createElement("span");
        wrap.className = "th-dc2-wrap";
        wrap.appendChild(inp);
        if (req) {
          var star = document.createElement("span");
          star.className = "req-star";
          star.textContent = "*";
          wrap.appendChild(star);
        }
        var btnDel = document.createElement("button");
        btnDel.type = "button";
        btnDel.className = "btn-delete-column";
        btnDel.textContent = "\u00D7";
        btnDel.title = "Delete column";
        btnDel.addEventListener("click", (function (col) { return function (e) { e.stopPropagation(); e.preventDefault(); deleteUserColumn(state.activeSheet, col); }; })(c));
        wrap.appendChild(btnDel);
        th.appendChild(wrap);
      } else {
        if (req) {
          var wrap = document.createElement("span");
          wrap.className = "th-header-wrap";
          wrap.appendChild(inp);
          var star = document.createElement("span");
          star.className = "req-star";
          star.textContent = "*";
          wrap.appendChild(star);
          th.appendChild(wrap);
        } else {
          th.appendChild(inp);
        }
      }
      if (state.activeGroup === "PeriodData") (function (col) { th.addEventListener("contextmenu", function (e) { e.preventDefault(); insertColumnRight(state.activeSheet, col); }); })(c);
      tr1.appendChild(th);
    }
    var thAct = document.createElement("th");
    thAct.className = "row-actions";
    thAct.setAttribute("rowspan", "3");
    tr1.appendChild(thAct);
    thead.appendChild(tr1);
    var tr2 = document.createElement("tr");
    for (var c = 3; c < headers.length; c++) {
      var th = document.createElement("th");
      var inp = document.createElement("input");
      inp.className = "th-input";
      inp.type = "text";
      inp.value = (h2[c] != null) ? String(h2[c]) : "";
      inp.addEventListener("input", function (idx) {
        return function () {
          var s = state.data[state.activeSheet];
          if (!s.headers2) s.headers2 = [];
          s.headers2[idx] = inp.value;
          autoSave();
        };
      }(c));
      th.appendChild(inp);
      tr2.appendChild(th);
    }
    thead.appendChild(tr2);
    var tr3 = document.createElement("tr");
    for (var c = 3; c < headers.length; c++) {
      var th = document.createElement("th");
      var inp = document.createElement("input");
      inp.className = "th-input";
      inp.type = "text";
      inp.value = (h3[c] != null) ? String(h3[c]) : "";
      inp.addEventListener("input", function (idx) {
        return function () {
          var s = state.data[state.activeSheet];
          if (!s.headers3) s.headers3 = [];
          s.headers3[idx] = inp.value;
          autoSave();
        };
      }(c));
      th.appendChild(inp);
      tr3.appendChild(th);
    }
    thead.appendChild(tr3);
  } else {
    thead.innerHTML = "";
    var tr = document.createElement("tr");
    tr.appendChild(createCell("th", "row-num", "#"));
    var isPeriodDataHeader = (state.activeGroup === "PeriodData" && !isTableMapping);
    // Service Driver special case: allow renaming only the rightmost two columns (Column4 and Column5)
    var isServiceDriver = (state.activeSheet === "Service Driver_Period" && state.activeGroup === "PeriodData");
    var lastTwoColIndices = isServiceDriver ? [headers.length - 2, headers.length - 1] : [];
    headers.forEach(function (h, colIndex) {
      var req = isRequired(state.activeSheet, h);
      var th = document.createElement("th");
      if (req) th.classList.add("required");
      th.setAttribute("data-col", String(colIndex));
      
      // Service Driver: render last two columns as editable inputs
      if (isServiceDriver && lastTwoColIndices.indexOf(colIndex) !== -1) {
        var inp = document.createElement("input");
        inp.className = "th-input";
        inp.type = "text";
        inp.value = (h != null) ? String(h) : "";
        var headerFocusValue;
        inp.addEventListener("focus", function () { headerFocusValue = inp.value; });
        inp.addEventListener("blur", function () {
          var newValue = inp.value.trim();
          if (newValue !== headerFocusValue) {
            // Ensure uniqueness
            var uniqueName = makeUniqueHeaderName(sheet.headers, newValue, colIndex);
            sheet.headers[colIndex] = uniqueName;
            autoSave();
            renderTable(); // Re-render to show updated header name
          }
        });
        inp.addEventListener("keydown", function (e) {
          if (e.key === "Enter") {
            inp.blur();
          }
        });
        th.appendChild(inp);
        if (req) {
          var star = document.createElement("span");
          star.className = "req-star";
          star.textContent = "*";
          th.appendChild(star);
        }
      } else {
        // Normal header rendering for other columns
        var wrap = document.createElement("span");
        wrap.className = "th-header-wrap";
        wrap.textContent = (h != null) ? String(h) : "";
        if (req) {
          var star = document.createElement("span");
          star.className = "req-star";
          star.textContent = "*";
          wrap.appendChild(star);
        }
        th.appendChild(wrap);
      }
      
      if (isPeriodDataHeader) {
        var btnIns = document.createElement("button");
        btnIns.type = "button";
        btnIns.className = "btn-insert-col";
        btnIns.textContent = "+";
        btnIns.title = "Insert column right";
        btnIns.addEventListener("click", function (e) { e.stopPropagation(); insertColumnRight(state.activeSheet, colIndex); });
        th.classList.add("th-has-insert");
        th.appendChild(btnIns);
        if (isUserAddedColumn(state.activeSheet, colIndex)) {
          var btnDel = document.createElement("button");
          btnDel.type = "button";
          btnDel.className = "btn-delete-column";
          btnDel.textContent = "\u00D7";
          btnDel.title = "Delete column";
          btnDel.addEventListener("click", function (e) { e.stopPropagation(); deleteUserColumn(state.activeSheet, colIndex); });
          th.appendChild(btnDel);
          th.addEventListener("dblclick", function () {
            var newName = prompt("Rename column:", (sheet.headers[colIndex] != null) ? String(sheet.headers[colIndex]) : "");
            if (newName !== null) renameColumn(state.activeSheet, colIndex, newName);
          });
        }
        th.addEventListener("contextmenu", function (e) {
          e.preventDefault();
          insertColumnRight(state.activeSheet, colIndex);
        });
      }
      tr.appendChild(th);
    });
    tr.appendChild(createCell("th", "row-actions", ""));
    thead.appendChild(tr);
  }

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
      // TableMapping 永遠唯讀
      if (isTableMapping) {
        input.setAttribute("readonly", "readonly");
        input.classList.add("cell-readonly");
      } else {
        input.setAttribute("readonly", "readonly");
      }
      input.dataset.sheet = state.activeSheet;
      input.dataset.row = String(rowIndex);
      input.dataset.col = String(colIndex);
      if (!validation.valid) {
        input.classList.add("cell-error");
        input.placeholder = validation.error;
      }
      if (!isTableMapping) {
        input.addEventListener("input", function () {
          updateCell(state.activeSheet, rowIndex, colIndex, input.value);
        });
      }
      const td = document.createElement("td");
      td.appendChild(input);
      tr.appendChild(td);
    });
    const actTd = document.createElement("td");
    actTd.className = "row-actions";
    // TableMapping 隱藏刪除按鈕
    if (!isTableMapping) {
      const delBtn = document.createElement("button");
      delBtn.type = "button";
      delBtn.className = "btn-delete-row";
      delBtn.textContent = "\u00D7";
      delBtn.title = "Delete row";
      delBtn.addEventListener("click", function () {
        deleteRow(state.activeSheet, rowIndex);
      });
      actTd.appendChild(delBtn);
    }
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

function addDriverCodeColumn() {
  var sheetName = "Resource Driver(Actvity Center)";
  if (state.activeSheet !== sheetName || state.activeGroup !== "PeriodData") return;
  var sheet = state.data[sheetName];
  if (!sheet || !sheet.headers || !sheet.data) return;
  var maxN = 0;
  for (var i = 2; i < sheet.headers.length; i++) {
    var m = String(sheet.headers[i] || "").match(/^Driver Code (\d+)$/);
    if (m) { var x = parseInt(m[1], 10); if (x > maxN) maxN = x; }
  }
  var N = maxN + 1;
  if (N < 3) N = 3;
  sheet.headers.push("Driver Code " + N);
  ensureUserAddedColIds(sheet);
  sheet.userAddedColIds.push("u_" + Date.now());
  for (var i = 0; i < sheet.data.length; i++) {
    sheet.data[i].push("");
  }
  renderTable();
  autoSave();
}

function deleteDriverCodeColumn(colIndex) {
  var sheetName = "Resource Driver(Actvity Center)";
  if (state.activeSheet !== sheetName || state.activeGroup !== "PeriodData") return;
  if (!isUserAddedColumn(sheetName, colIndex)) return;
  deleteUserColumn(sheetName, colIndex);
}

function addDriverCodeColumnMAC() {
  var sheetName = "Resource Driver(M. A. C.)";
  if (state.activeSheet !== sheetName || state.activeGroup !== "PeriodData") return;
  var sheet = state.data[sheetName];
  if (!sheet || !sheet.headers || !sheet.data) return;
  var maxN = 0;
  for (var i = 3; i < sheet.headers.length; i++) {
    var m = String(sheet.headers[i] || "").match(/^Driver (\d+)$/);
    if (m) { var x = parseInt(m[1], 10); if (x > maxN) maxN = x; }
  }
  var N = maxN + 1;
  if (N < 4) N = 4;
  sheet.headers.push("Driver " + N);
  if (!sheet.headers2) sheet.headers2 = [];
  sheet.headers2.push("");
  if (!sheet.headers3) sheet.headers3 = [];
  sheet.headers3.push("");
  ensureUserAddedColIds(sheet);
  sheet.userAddedColIds.push("u_" + Date.now());
  for (var i = 0; i < sheet.data.length; i++) {
    sheet.data[i].push("");
  }
  ensureMACHeaderRows();
  renderTable();
  autoSave();
}

// MAC column delete X start
function canDeleteMacColumn(colIndex) { return colIndex >= 4; }
// MAC column delete X end

function deleteDriverCodeColumnMAC(colIndex) {
  var sheetName = "Resource Driver(M. A. C.)";
  if (state.activeSheet !== sheetName || state.activeGroup !== "PeriodData") return;
  if (!isUserAddedColumn(sheetName, colIndex)) return;
  deleteUserColumn(sheetName, colIndex);
}

function addDriverCodeColumnSAC() {
  var sheetName = "Resource Driver(S. A. C.)";
  if (state.activeSheet !== sheetName || state.activeGroup !== "PeriodData") return;
  var sheet = state.data[sheetName];
  if (!sheet || !sheet.headers || !sheet.data) return;
  var maxN = 0;
  for (var i = 3; i < sheet.headers.length; i++) {
    var m = String(sheet.headers[i] || "").match(/^Driver (\d+)$/);
    if (m) { var x = parseInt(m[1], 10); if (x > maxN) maxN = x; }
  }
  var N = maxN + 1;
  if (N < 7) N = 7;
  sheet.headers.push("Driver " + N);
  if (!sheet.headers2) sheet.headers2 = [];
  sheet.headers2.push("");
  if (!sheet.headers3) sheet.headers3 = [];
  sheet.headers3.push("");
  ensureUserAddedColIds(sheet);
  sheet.userAddedColIds.push("u_" + Date.now());
  for (var i = 0; i < sheet.data.length; i++) {
    sheet.data[i].push("");
  }
  ensureSACHeaderRows();
  renderTable();
  autoSave();
}

// SAC column delete X start
function canDeleteSacColumn(colIndex) { return colIndex >= 4; }
// SAC column delete X end

function deleteDriverCodeColumnSAC(colIndex) {
  var sheetName = "Resource Driver(S. A. C.)";
  if (state.activeSheet !== sheetName || state.activeGroup !== "PeriodData") return;
  if (!isUserAddedColumn(sheetName, colIndex)) return;
  deleteUserColumn(sheetName, colIndex);
}

// --- Insert column right / Rename column (PeriodData only) ---
function makeUniqueHeaderName(headers, desired, excludeColIndex) {
  var base = (desired != null && String(desired).trim() !== "") ? String(desired).trim() : "New Column";
  var set = {};
  headers.forEach(function (h, i) {
    if (excludeColIndex === i) return;
    set[String(h || "").trim()] = true;
  });
  if (!set[base]) return base;
  if (base.length === 1 && base.match(/^[A-Z]$/i)) {
    for (var n = 2; n <= 1000; n++) {
      var cand = base + n;
      if (!set[cand]) return cand;
    }
  }
  for (var n = 2; n <= 1000; n++) {
    var cand = base + " " + n;
    if (!set[cand]) return cand;
  }
  return base + " " + Date.now();
}

function ensureUserAddedColIds(sheet) {
  if (!sheet || !sheet.headers) return;
  if (!sheet.userAddedColIds) sheet.userAddedColIds = [];
  while (sheet.userAddedColIds.length < sheet.headers.length) sheet.userAddedColIds.push("");
  if (sheet.userAddedColIds.length > sheet.headers.length) sheet.userAddedColIds.splice(sheet.headers.length);
}

function isUserAddedColumn(sheetName, colIndex) {
  var sheet = state.data[sheetName];
  if (!sheet || !sheet.headers || colIndex < 0 || colIndex >= sheet.headers.length) return false;
  var ids = sheet.userAddedColIds;
  if (ids && ids[colIndex]) return true;
  if (!ids && (sheetName === "Resource Driver(Actvity Center)" || sheetName === "Resource Driver(M. A. C.)" || sheetName === "Resource Driver(S. A. C.)")) return colIndex >= 4;
  return false;
}

function insertColumnAt(sheetName, colIndex, headerName) {
  var sheet = state.data[sheetName];
  if (!sheet || !sheet.headers) return;
  sheet.headers.splice(colIndex, 0, headerName);
  if (sheet.headers2) sheet.headers2.splice(colIndex, 0, "");
  if (sheet.headers3) sheet.headers3.splice(colIndex, 0, "");
  if (sheet.userAddedColIds) sheet.userAddedColIds.splice(colIndex, 0, "");
  for (var i = 0; i < (sheet.data || []).length; i++) {
    sheet.data[i].splice(colIndex, 0, "");
  }
}

function removeColumn(sheetName, colIndex) {
  var sheet = state.data[sheetName];
  if (!sheet || !sheet.headers) return;
  sheet.headers.splice(colIndex, 1);
  if (sheet.headers2) sheet.headers2.splice(colIndex, 1);
  if (sheet.headers3) sheet.headers3.splice(colIndex, 1);
  if (sheet.userAddedColIds) sheet.userAddedColIds.splice(colIndex, 1);
  for (var i = 0; i < (sheet.data || []).length; i++) {
    sheet.data[i].splice(colIndex, 1);
  }
}

function insertColumnRight(sheetName, colIndex) {
  if (state.activeGroup !== "PeriodData" || sheetName === "TableMapping") return;
  var sheet = state.data[sheetName];
  if (!sheet || !sheet.headers) return;
  var base = (sheet.headers[colIndex] != null) ? String(sheet.headers[colIndex]).trim() : "";
  var newName = makeUniqueHeaderName(sheet.headers, base || null);
  var insertAt = colIndex + 1;
  var userAddedId = "u_" + Date.now();
  beginAction("Insert column");
  recordChangeSpecial({ type: "insert_col", sheet: sheetName, colIndex: insertAt, insertedHeaderName: newName, userAddedId: userAddedId });
  insertColumnAt(sheetName, insertAt, newName);
  ensureUserAddedColIds(sheet);
  sheet.userAddedColIds[insertAt] = userAddedId;
  commitAction();
  state.changeLog.push({
    timestamp: new Date().toISOString(),
    sheet: getExcelSheetName(sheetName),
    row: "",
    column: newName,
    oldValue: "",
    newValue: "[Column inserted]"
  });
  autoSave();
  var saved = { active: state.activeCell, sel: state.selection };
  renderTable();
  if (saved.active) {
    if (saved.active.col > colIndex) state.activeCell = { row: saved.active.row, col: saved.active.col + 1 };
    else if (saved.active.col === colIndex) state.activeCell = { row: saved.active.row, col: insertAt };
    else state.activeCell = { row: saved.active.row, col: saved.active.col };
  } else if (sheet.data.length) {
    state.activeCell = { row: 0, col: insertAt };
  }
  state.selection = null;
  state.multiSelection = [];
  updateSelectionUI();
  focusActiveCell();
}

function renameColumn(sheetName, colIndex, newName, opts) {
  opts = opts || {};
  if (!isUserAddedColumn(sheetName, colIndex)) return;
  var sheet = state.data[sheetName];
  if (!sheet || !sheet.headers) return;
  var oldName = (sheet.headers[colIndex] != null) ? String(sheet.headers[colIndex]) : "";
  var trimmed = String(newName).trim();
  if (trimmed === "") return;
  if (trimmed === oldName) return;
  trimmed = makeUniqueHeaderName(sheet.headers, trimmed, colIndex);
  if (!opts.skipUndo) {
    beginAction("Rename column");
    recordChangeSpecial({ type: "rename_col", sheet: sheetName, colIndex: colIndex, oldName: oldName, newName: trimmed });
    commitAction();
  }
  sheet.headers[colIndex] = trimmed;
  if (!opts.skipLog) {
    state.changeLog.push({
      timestamp: new Date().toISOString(),
      sheet: getExcelSheetName(sheetName),
      row: "",
      column: trimmed,
      oldValue: oldName,
      newValue: trimmed
    });
  }
  autoSave();
  renderTable();
  updateSelectionUI();
}

function recordChangeSpecial(changeObj) {
  var label = (changeObj.type === "insert_col" ? "Insert column" : changeObj.type === "delete_col" ? "Delete column" : "Rename column");
  if (state._tx) {
    state._tx.changes.push(changeObj);
  } else {
    state.undoStack.push({ id: Date.now(), label: label, changes: [changeObj] });
    clearRedo();
  }
}

function deleteUserColumn(sheetName, colIndex) {
  if (state.activeGroup !== "PeriodData" || sheetName === "TableMapping") return;
  if (!isUserAddedColumn(sheetName, colIndex)) return;
  var sheet = state.data[sheetName];
  if (!sheet || !sheet.headers || !sheet.data) return;
  if (!confirm("Are you sure you want to delete this column? All data in this column will be deleted.")) return;
  var deletedHeaderName = (sheet.headers[colIndex] != null) ? String(sheet.headers[colIndex]) : "";
  var deletedHeader2 = (sheet.headers2 && sheet.headers2[colIndex] != null) ? String(sheet.headers2[colIndex]) : "";
  var deletedHeader3 = (sheet.headers3 && sheet.headers3[colIndex] != null) ? String(sheet.headers3[colIndex]) : "";
  var deletedValues = sheet.data.map(function (row) { return (row[colIndex] != null) ? String(row[colIndex]) : ""; });
  var userAddedId = (sheet.userAddedColIds && sheet.userAddedColIds[colIndex]) ? sheet.userAddedColIds[colIndex] : "";
  beginAction("Delete column");
  recordChangeSpecial({ type: "delete_col", sheet: sheetName, colIndex: colIndex, deletedHeaderName: deletedHeaderName, deletedHeader2: deletedHeader2, deletedHeader3: deletedHeader3, deletedValues: deletedValues, userAddedId: userAddedId });
  removeColumn(sheetName, colIndex);
  commitAction();
  state.changeLog.push({
    timestamp: new Date().toISOString(),
    sheet: getExcelSheetName(sheetName),
    row: "",
    column: deletedHeaderName,
    oldValue: "[Column deleted]",
    newValue: ""
  });
  autoSave();
  renderTable();
  updateSelectionUI();
}

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
      // 儲存當前資料
      if (state.activeGroup === "PeriodData" && state.activePeriod) {
        saveToStorage();
      } else if (state.activeGroup === "ModelData") {
        saveToStorage();
      }
      
      state.activeGroup = wb;
      if (wb === "PeriodData") {
        // 切換到 PeriodData：載入當前月份或第一個可用月份
        var periods = getAllPeriodsFromStorage();
        if (periods.length > 0) {
          state.activePeriod = periods[0];
          switchPeriod(state.activePeriod);
        } else {
          state.activePeriod = null;
          const order = getSheetsForWorkbook(wb);
          state.activeSheet = order[0] || "Exchange Rate";
          clearCutState();
          renderAll();
        }
      } else {
        // 切換到 ModelData：清除 activePeriod
        state.activePeriod = null;
        const order = getSheetsForWorkbook(wb);
        state.activeSheet = order[0] || "Company";
        clearCutState();
        renderAll();
      }
      autoSave();
    });
  });

  // Period selector 事件
  var periodSelect = document.getElementById("periodSelect");
  if (periodSelect) {
    periodSelect.addEventListener("change", function () {
      var period = this.value;
      if (period && period !== state.activePeriod) {
        switchPeriod(period);
      }
    });
  }

  var btnCreatePeriod = document.getElementById("btnCreatePeriod");
  if (btnCreatePeriod) {
    btnCreatePeriod.addEventListener("click", function () {
      createNewPeriod();
    });
  }

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

  // Required Fields Admin Modal events
  var btnRequiredFields = document.getElementById("btnRequiredFields");
  if (btnRequiredFields) {
    btnRequiredFields.addEventListener("click", function () {
      showRequiredFieldsModal();
    });
  }

  var requiredFieldsSheetSelect = document.getElementById("requiredFieldsSheetSelect");
  if (requiredFieldsSheetSelect) {
    requiredFieldsSheetSelect.addEventListener("change", function () {
      var sheetName = this.value;
      if (sheetName) {
        renderRequiredFieldsColumnList(sheetName);
      } else {
        document.getElementById("requiredFieldsColumnList").innerHTML = '<p style="color: var(--color-text-muted); font-size: 13px;">Please select a sheet to configure required fields.</p>';
      }
    });
  }

  var btnRequiredFieldsSave = document.getElementById("btnRequiredFieldsSave");
  if (btnRequiredFieldsSave) {
    btnRequiredFieldsSave.addEventListener("click", function () {
      saveRequiredFieldsOverride();
    });
  }

  var btnRequiredFieldsReset = document.getElementById("btnRequiredFieldsReset");
  if (btnRequiredFieldsReset) {
    btnRequiredFieldsReset.addEventListener("click", function () {
      resetRequiredFieldsOverride();
    });
  }

  var btnRequiredFieldsCancel = document.getElementById("btnRequiredFieldsCancel");
  if (btnRequiredFieldsCancel) {
    btnRequiredFieldsCancel.addEventListener("click", function () {
      hideRequiredFieldsModal();
    });
  }

  // Close modal when clicking outside
  var requiredFieldsModal = document.getElementById("requiredFieldsModal");
  if (requiredFieldsModal) {
    requiredFieldsModal.addEventListener("click", function (e) {
      if (e.target === requiredFieldsModal) {
        hideRequiredFieldsModal();
      }
    });
  }

  // Optional Tabs Modal events
  var btnOptionalTabs = document.getElementById("btnOptionalTabs");
  if (btnOptionalTabs) {
    btnOptionalTabs.addEventListener("click", function () {
      showOptionalTabsModal();
    });
  }

  var btnOptionalTabsClose = document.getElementById("btnOptionalTabsClose");
  if (btnOptionalTabsClose) {
    btnOptionalTabsClose.addEventListener("click", function () {
      hideOptionalTabsModal();
    });
  }

  var optionalTabsModal = document.getElementById("optionalTabsModal");
  if (optionalTabsModal) {
    optionalTabsModal.addEventListener("click", function (e) {
      if (e.target === optionalTabsModal) {
        hideOptionalTabsModal();
      }
    });
  }

  var btnOptionalTabsSave = document.getElementById("btnOptionalTabsSave");
  if (btnOptionalTabsSave) {
    btnOptionalTabsSave.addEventListener("click", function () {
      saveOptionalTabsFromUI();
    });
  }

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
      // TableMapping 不允許編輯
      if (state.activeSheet === "TableMapping") {
        showStatus("TableMapping is system-defined and read-only.", "error");
        return;
      }
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
        // TableMapping 不允許剪下
        if (state.activeSheet === "TableMapping") {
          e.preventDefault();
          showStatus("TableMapping is system-defined and read-only.", "error");
          return;
        }
        console.log("[keydown] Ctrl+X", { k: k, hasTextSelection: hasTextSelection, inTable: !!inTable });
        e.preventDefault();
        doCut();
        return;
      }
      if (k === "v") {
        // TableMapping 不允許貼上
        if (state.activeSheet === "TableMapping") {
          e.preventDefault();
          showStatus("TableMapping is system-defined and read-only.", "error");
          return;
        }
        e.preventDefault();
        doPaste();
        return;
      }
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
      // TableMapping 不允許編輯
      if (state.activeSheet === "TableMapping") {
        showStatus("TableMapping is system-defined and read-only.", "error");
        return;
      }
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
      // TableMapping 不允許編輯
      if (state.activeSheet === "TableMapping") {
        e.preventDefault();
        showStatus("TableMapping is system-defined and read-only.", "error");
        return;
      }
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
