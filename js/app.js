/**
 * Excel Teaching Form - Application Logic
 * UI table data = Source of Truth. Download rebuilds from headers + data arrays.
 */

const state = {
  studentId: null,
  activeGroup: "ModelData",
  activeSheet: "Company",
  data: {},
  changeLog: []
};

let saveTimeout = null;

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

function initFromTemplate() {
  state.data = {};
  state.changeLog = [];
  for (const [sheetName, config] of Object.entries(TEMPLATE_DATA.sheets)) {
    if (!config.headers || config.headers.length === 0) {
      state.data[sheetName] = { headers: [], data: [] };
      continue;
    }
    state.data[sheetName] = {
      headers: [...config.headers],
      data: config.data.map(row => [...row])
    };
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
      state.data[sheetName] = {
        headers: [...config.headers],
        data: config.data.map(function (row) { return [...row]; })
      };
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

// --- Cell update & ChangeLog ---

function updateCell(sheetName, rowIndex, colIndex, newValue) {
  const sheet = state.data[sheetName];
  if (!sheet || !sheet.data[rowIndex]) return;
  const oldValue = sheet.data[rowIndex][colIndex];
  if (oldValue === newValue) return;

  sheet.data[rowIndex][colIndex] = newValue;
  const excelSheetName = getExcelSheetName(sheetName);
  state.changeLog.push({
    timestamp: new Date().toISOString(),
    sheet: excelSheetName,
    row: rowIndex + 2,
    column: sheet.headers[colIndex],
    oldValue: oldValue,
    newValue: newValue
  });
  autoSave();
  updateValidation(sheetName, rowIndex, colIndex);
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

function downloadExcel() {
  if (!state.studentId) {
    showStatus("Enter Student ID first", "error");
    return;
  }
  const timestamp = new Date().toISOString().slice(0, 16).replace(/[-:T]/g, "");
  downloadWorkbook("ModelData", timestamp);
  setTimeout(function () {
    downloadWorkbook("PeriodData", timestamp);
  }, 500);
  showStatus("Downloaded both files");
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

    if (config.hidden) {
      const ws = XLSX.utils.aoa_to_sheet([[]]);
      XLSX.utils.book_append_sheet(wb, ws, excelSheetName);
      return;
    }

    if (!sheet) return;
    const wsData = [sheet.headers].concat(sheet.data);
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
  renderSheetTabs();
  renderTable();
}

function renderWorkbookToggle() {
  document.querySelectorAll(".workbook-btn").forEach(function (btn) {
    btn.classList.toggle("active", btn.getAttribute("data-workbook") === state.activeGroup);
  });
}

function renderSheetTabs() {
  const container = document.getElementById("sheetTabs");
  container.innerHTML = "";
  const order = getSheetsForWorkbook(state.activeGroup);
  order.forEach(function (internalName) {
    const config = getSheetConfig(internalName);
    if (config && config.hidden) return;
    const tab = document.createElement("button");
    tab.type = "button";
    tab.className = "sheet-tab" + (internalName === state.activeSheet ? " active" : "");
    tab.textContent = getExcelSheetName(internalName);
    tab.addEventListener("click", function () {
      state.activeSheet = internalName;
      renderAll();
      autoSave();
    });
    container.appendChild(tab);
  });
}

function renderTable() {
  const sheet = state.data[state.activeSheet];
  const thead = document.getElementById("tableHead");
  const tbody = document.getElementById("tableBody");
  const config = getSheetConfig(state.activeSheet);

  document.getElementById("sheetTitle").textContent = getExcelSheetName(state.activeSheet);

  if (!sheet || !config || config.hidden) {
    thead.innerHTML = "";
    tbody.innerHTML = "<tr><td colspan='3'>No data</td></tr>";
    document.getElementById("btnAddRow").style.display = "none";
    document.getElementById("validationSummary").classList.add("hidden");
    return;
  }

  document.getElementById("btnAddRow").style.display = "";

  const headers = sheet.headers;
  let thHtml = "<tr><th class=\"row-num\">#</th>";
  headers.forEach(function (h) {
    const req = isRequired(state.activeSheet, h);
    thHtml += "<th" + (req ? " class=\"required\"" : "") + ">" + escapeHtml(h) + "</th>";
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
}

// --- Run ---

if (typeof XLSX !== "undefined") {
  init();
} else {
  document.addEventListener("DOMContentLoaded", function () {
    if (typeof XLSX !== "undefined") init();
  });
}
