console.log("✅ [table_core.js] loaded");

window.DEFS = window.DEFS || {};
window.DEFS.TABLE_CORE = window.DEFS.TABLE_CORE || {};

(function installTableCore(){

  window.DEFS.TABLE_CORE.init = function initTableCore(ctx){
    ctx = ctx || {};
    console.log("✅ [table_core.js] init");

    // ---------- state (global compatible) ----------
    window.__CELL_EDIT_MODE = window.__CELL_EDIT_MODE || false;
    window.__EDIT_CELL_KEY  = window.__EDIT_CELL_KEY  || "";

    // ---------- helpers ----------
    function ensureSize(sheet) {
      while (sheet.data.length < sheet.rows) sheet.data.push(Array(sheet.cols).fill(""));
      for (let r=0; r<sheet.data.length; r++) while (sheet.data[r].length < sheet.cols) sheet.data[r].push("");
    }

    function parseClipboardGrid(text) {
      const cleaned = String(text).replace(/\r/g, "");
      return cleaned.split("\n").filter(line => line !== "").map(line => line.split("\t"));
    }

    function placeCaretFromMouseEvent(td, evt){
      try{
        td.focus();
        const x = evt.clientX, y = evt.clientY;
        let range = null;

        if (document.caretRangeFromPoint) {
          range = document.caretRangeFromPoint(x, y);
        } else if (document.caretPositionFromPoint) {
          const pos = document.caretPositionFromPoint(x, y);
          if (pos) {
            range = document.createRange();
            range.setStart(pos.offsetNode, pos.offset);
            range.collapse(true);
          }
        }

        const sel = window.getSelection();
        if (!sel) return;
        sel.removeAllRanges();

        if (range) sel.addRange(range);
        else sel.collapse(td, 1);
      } catch {}
    }

    function enterEditMode(td, evt){
      window.__CELL_EDIT_MODE = true;
      window.__EDIT_CELL_KEY = `${td.dataset.r},${td.dataset.c}`;
      if (document.body) document.body.classList.add("cell-edit-mode");

      const sel = window.getSelection();
      try { sel?.removeAllRanges(); } catch {}

      requestAnimationFrame(() => {
        requestAnimationFrame(() => placeCaretFromMouseEvent(td, evt));
      });
    }

    function exitEditMode(){
      window.__CELL_EDIT_MODE = false;
      window.__EDIT_CELL_KEY = "";
      if (document.body) document.body.classList.remove("cell-edit-mode");
    }

    function focusCell(r, c) {
      const gridBody = (ctx && ctx.gridBody) ? ctx.gridBody : document.getElementById("gridBody");

      const td = gridBody?.querySelector?.(`td[data-r="${r}"][data-c="${c}"]`);
      if (!td) return false;
      td.focus();
      return true;
    }

    // ---------- bind events (moved from app.js MODULE 12) ----------
    function bindTableEvents(){
  const gridBody = (ctx && ctx.gridBody) ? ctx.gridBody : document.getElementById("gridBody");
  if (!gridBody) return;

  // ✅ global guard（不綁第二次）
  if (window.__TABLE_CORE_EVENTS_BOUND__) return;
  window.__TABLE_CORE_EVENTS_BOUND__ = true;


      // click: edit mode toggle
      gridBody.addEventListener("click", (e) => {
        const td = e.target;
        if (!(td instanceof HTMLElement) || td.tagName !== "TD") return;

        const key = `${td.dataset.r},${td.dataset.c}`;

        if (window.__CELL_EDIT_MODE && window.__EDIT_CELL_KEY === key) return;

        if (e.detail >= 2) {
          enterEditMode(td, e);
          return;
        }

        if (window.__CELL_EDIT_MODE && window.__EDIT_CELL_KEY !== key) {
          exitEditMode();
        }
      });

      // click outside -> exit edit mode
      document.addEventListener("mousedown", (e) => {
        const inside = e.target instanceof Node && gridBody.contains(e.target);
        if (!inside && window.__CELL_EDIT_MODE) exitEditMode();
      });

      // esc -> exit edit mode
      document.addEventListener("keydown", (e) => {
        if (e.key === "Escape") exitEditMode();
      });

      // input -> write back
      gridBody.addEventListener("input", (e) => {
        const tEl = e.target;
        if (!(tEl instanceof HTMLElement) || tEl.tagName !== "TD") return;

        const r = Number(tEl.dataset.r), c = Number(tEl.dataset.c);
        if (!Number.isFinite(r) || !Number.isFinite(c)) return;

        const s = (typeof ctx.activeSheet === "function") ? ctx.activeSheet() : null;
        if (!s) return;

        ensureSize(s);
        s.data[r][c] = tEl.textContent ?? "";
        if (typeof ctx.saveToLocalByMode === "function") ctx.saveToLocalByMode(ctx.activeMode);
      });

      // paste: edit mode -> native, else block paste
      gridBody.addEventListener("paste", (e) => {
        const tEl = e.target;
        if (!(tEl instanceof HTMLElement) || tEl.tagName !== "TD") return;

        if (window.__CELL_EDIT_MODE) return;

        const text = e.clipboardData?.getData("text") ?? "";
        if (!(text.includes("\t") || text.includes("\n"))) return;

        e.preventDefault();

        const startR = Number(tEl.dataset.r);
        const startC = Number(tEl.dataset.c);
        if (!Number.isFinite(startR) || !Number.isFinite(startC)) return;

        const block = parseClipboardGrid(text);
        if (!block.length) return;

        const s = (typeof ctx.activeSheet === "function") ? ctx.activeSheet() : null;
        if (!s) return;

        const needRows = startR + block.length;
        const needCols = startC + Math.max(...block.map(row => row.length));

        if (needRows > s.rows) s.rows = needRows;
        if (needCols > s.cols) s.cols = needCols;

        if (typeof ctx.ensureHeadersForActiveSheet === "function") ctx.ensureHeadersForActiveSheet();
        ensureSize(s);

        for (let r=0; r<block.length; r++) {
          for (let c=0; c<block[r].length; c++) s.data[startR + r][startC + c] = block[r][c] ?? "";
        }

        const doRender = (typeof ctx.render === "function") ? ctx.render : (typeof window.render === "function" ? window.render : null);
        if (doRender) doRender();

        focusCell(startR, startC);
        setTimeout(() => {
          if (typeof ctx.saveToLocalByMode === "function") ctx.saveToLocalByMode(ctx.activeMode);
        }, 0);
      });

      // arrow keys: edit mode -> native, else move cell
      gridBody.addEventListener("keydown", (e) => {
        const td = e.target;
        if (!(td instanceof HTMLElement) || td.tagName !== "TD") return;

        if (window.__CELL_EDIT_MODE) return;

        const key = e.key;
        if (!["ArrowUp","ArrowDown","ArrowLeft","ArrowRight"].includes(key)) return;

        e.preventDefault();

        const r = Number(td.dataset.r);
        const c = Number(td.dataset.c);
        if (!Number.isFinite(r) || !Number.isFinite(c)) return;

        const s = (typeof ctx.activeSheet === "function") ? ctx.activeSheet() : null;
        if (!s) return;

        let nr = r, nc = c;

        if (key === "ArrowUp") nr = Math.max(0, r - 1);
        if (key === "ArrowDown") nr = Math.min(s.rows - 1, r + 1);
        if (key === "ArrowLeft") nc = Math.max(0, c - 1);
        if (key === "ArrowRight") nc = Math.min(s.cols - 1, c + 1);

        focusCell(nr, nc);

        if (typeof window.setSelection === "function") {
          try { window.setSelection({r:nr, c:nc}, {r:nr, c:nc}); } catch {}
        }
      });
    }

    // ---------- export to ctx + window (compat) ----------
    ctx.ensureSize = ensureSize;
    ctx.parseClipboardGrid = parseClipboardGrid;
    ctx.placeCaretFromMouseEvent = placeCaretFromMouseEvent;
    ctx.enterEditMode = enterEditMode;
    ctx.exitEditMode = exitEditMode;
    ctx.focusCell = focusCell;
    ctx.bindTableEvents = bindTableEvents;

    window.ensureSize = ensureSize;
    window.parseClipboardGrid = parseClipboardGrid;
    window.placeCaretFromMouseEvent = placeCaretFromMouseEvent;
    window.enterEditMode = enterEditMode;
    window.exitEditMode = exitEditMode;
    window.focusCell = focusCell;
    window.bindTableEvents = bindTableEvents;

    // auto-bind now
    bindTableEvents();

    return ctx;
  };

})();
