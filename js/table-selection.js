/**
 * Table selection: drag-select, Delete, copy/cut/paste (Excel-compatible TSV).
 * Depends on: state, updateCell (from app.js). Init runs after DOM ready.
 * Exposes: clearTableSelection (for renderTable to call), initTableSelection.
 */
(function () {
  "use strict";

  var isDragging = false;
  var rect = null; // { r0, c0, r1, c1 } r0,c0=start, r1,c1=end (can be < r0/c0)

  var TABLE_ID = "dataTable";

  function getTable() {
    return document.getElementById(TABLE_ID);
  }

  function getCellFromElement(el) {
    if (!el) return null;
    var td = el.closest ? el.closest("td") : null;
    if (!td) return null;
    var inp = td.querySelector("input[data-sheet][data-row][data-col]");
    if (!inp) return null;
    var r = parseInt(inp.dataset.row, 10);
    var c = parseInt(inp.dataset.col, 10);
    if (isNaN(r) || isNaN(c)) return null;
    return { td: td, inp: inp, sheet: inp.dataset.sheet, row: r, col: c };
  }

  function getCellFromEvent(e) {
    var td = e.target && e.target.closest ? e.target.closest("td") : null;
    if (!td) return null;
    return getCellFromElement(td);
  }

  function clearSelectionState() {
    rect = null;
    isDragging = false;
  }

  function applyVisual() {
    var table = getTable();
    if (!table) return;
    table.querySelectorAll(".cell-selected, .cell-active").forEach(function (el) {
      el.classList.remove("cell-selected", "cell-active");
    });
    if (!rect) return;
    var r0 = rect.r0, c0 = rect.c0, r1 = rect.r1, c1 = rect.c1;
    var rMin = Math.min(r0, r1), rMax = Math.max(r0, r1);
    var cMin = Math.min(c0, c1), cMax = Math.max(c0, c1);
    var sheet = typeof state !== "undefined" && state && state.activeSheet ? state.activeSheet : "";
    for (var r = rMin; r <= rMax; r++) {
      for (var c = cMin; c <= cMax; c++) {
        var inp = table.querySelector('input[data-sheet="' + sheet + '"][data-row="' + r + '"][data-col="' + c + '"]');
        if (inp) {
          var cell = inp.closest("td");
          if (cell) {
            cell.classList.add("cell-selected");
            if (r === r0 && c === c0) cell.classList.add("cell-active");
          }
        }
      }
    }
  }

  function clearTableSelection() {
    clearSelectionState();
    applyVisual();
  }

  function clearSelectedCells() {
    if (!rect || typeof state === "undefined" || !state || typeof updateCell !== "function") return;
    var sheet = state.activeSheet;
    if (!sheet) return;
    var r0 = rect.r0, c0 = rect.c0, r1 = rect.r1, c1 = rect.c1;
    var rMin = Math.min(r0, r1), rMax = Math.max(r0, r1);
    var cMin = Math.min(c0, c1), cMax = Math.max(c0, c1);
    var table = getTable();
    for (var r = rMin; r <= rMax; r++) {
      for (var c = cMin; c <= cMax; c++) {
        updateCell(sheet, r, c, "");
        var inp = table && table.querySelector('input[data-sheet="' + sheet + '"][data-row="' + r + '"][data-col="' + c + '"]');
        if (inp) inp.value = "";
      }
    }
  }

  function getSelectedTsv() {
    if (!rect || typeof state === "undefined" || !state) return "";
    var sheet = state.data[state.activeSheet];
    if (!sheet || !sheet.data) return "";
    var r0 = rect.r0, c0 = rect.c0, r1 = rect.r1, c1 = rect.c1;
    var rMin = Math.min(r0, r1), rMax = Math.max(r0, r1);
    var cMin = Math.min(c0, c1), cMax = Math.max(c0, c1);
    var rows = [];
    for (var r = rMin; r <= rMax; r++) {
      var cells = [];
      for (var c = cMin; c <= cMax; c++) {
        var v = (sheet.data[r] || [])[c];
        cells.push(v == null ? "" : String(v));
      }
      rows.push(cells.join("\t"));
    }
    return rows.join("\n");
  }

  function pasteIntoSelection(text) {
    if (!rect || typeof state === "undefined" || !state || typeof updateCell !== "function") return;
    var sheet = state.activeSheet;
    if (!sheet) return;
    var rows = (text || "").split(/\r?\n/);
    var r0 = rect.r0, c0 = rect.c0;
    var table = getTable();
    for (var i = 0; i < rows.length; i++) {
      var cols = rows[i].split("\t");
      for (var j = 0; j < cols.length; j++) {
        var r = r0 + i, c = c0 + j;
        if (!sheet.data[r]) continue;
        if (c >= (sheet.headers || []).length) continue;
        var val = cols[j] || "";
        updateCell(sheet, r, c, val);
        var inp = table && table.querySelector('input[data-sheet="' + sheet + '"][data-row="' + r + '"][data-col="' + c + '"]');
        if (inp) inp.value = val;
      }
    }
  }

  function hasSelection() {
    return !!rect;
  }

  function isFocusedTableInput() {
    var a = document.activeElement;
    if (!a || (a.tagName !== "INPUT" && a.tagName !== "TEXTAREA")) return false;
    var t = getTable();
    return t && t.contains(a);
  }

  function onMousedown(e) {
    if (e.target.tagName === "INPUT" && document.activeElement === e.target) return;
    var cell = getCellFromEvent(e);
    if (!cell) return;
    e.preventDefault();
    isDragging = true;
    rect = { r0: cell.row, c0: cell.col, r1: cell.row, c1: cell.col };
    applyVisual();
  }

  function onMousemove(e) {
    if (!isDragging) return;
    var el = document.elementFromPoint(e.clientX, e.clientY);
    var cell = getCellFromElement(el);
    if (!cell) return;
    rect.r1 = cell.row;
    rect.c1 = cell.col;
    applyVisual();
  }

  function onMouseup() {
    isDragging = false;
  }

  function onClickOutside(e) {
    if (!hasSelection()) return;
    var table = getTable();
    if (table && table.contains(e.target)) return;
    clearTableSelection();
  }

  function onKeydown(e) {
    if (e.key === "Escape") {
      clearTableSelection();
      return;
    }
    if (!hasSelection()) return;
    if (e.key === "Delete" || e.key === "Backspace") {
      if (isFocusedTableInput()) return;
      e.preventDefault();
      clearSelectedCells();
    }
  }

  function onCopy(e) {
    if (!hasSelection()) return;
    var tsv = getSelectedTsv();
    e.preventDefault();
    if (e.clipboardData) e.clipboardData.setData("text/plain", tsv);
  }

  function onCut(e) {
    if (!hasSelection()) return;
    var tsv = getSelectedTsv();
    e.preventDefault();
    if (e.clipboardData) e.clipboardData.setData("text/plain", tsv);
    clearSelectedCells();
    clearTableSelection();
  }

  function onPaste(e) {
    if (!hasSelection()) return;
    var text = (e.clipboardData && e.clipboardData.getData("text/plain")) || "";
    if (!text) return;
    e.preventDefault();
    pasteIntoSelection(text);
  }

  function initTableSelection() {
    var table = getTable();
    if (!table) return;
    table.addEventListener("mousedown", onMousedown);
    document.addEventListener("mousemove", onMousemove);
    document.addEventListener("mouseup", onMouseup);
    document.addEventListener("click", onClickOutside, true);
    document.addEventListener("keydown", onKeydown);
    document.addEventListener("copy", onCopy);
    document.addEventListener("cut", onCut);
    document.addEventListener("paste", onPaste);
  }

  window.clearTableSelection = clearTableSelection;
  window.initTableSelection = initTableSelection;

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initTableSelection);
  } else {
    initTableSelection();
  }
})();
