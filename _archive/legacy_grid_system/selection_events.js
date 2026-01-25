console.log("✅ [selection_events.js] loaded");

window.DEFS = window.DEFS || {};
window.DEFS.SELECTION_EVENTS = window.DEFS.SELECTION_EVENTS || {};

(function installSelectionEvents(){

  function tryBindOnce(){
    const API = window.DEFS?.SELECTION_CORE?.api;
    if (!API) return false;

    if (window.__SELECTION_EVENTS_BINDED__) return true;
    window.__SELECTION_EVENTS_BINDED__ = true;

    // =========================================================
    //  Excel-like Visual Override Layer (black active cell vs blue range)
    // =========================================================
    const VIS = {
      activeTD: null,
      rangeTDs: [],
      activeOutline: "2px solid #000",
      rangeBg: "rgba(120, 170, 255, 0.25)"
    };

    function clearCaret(){
      try{
        const sel = window.getSelection?.();
        if (sel) sel.removeAllRanges();
      } catch {}
    }

    function clearVisual(){
      try{
        if (VIS.activeTD){
          VIS.activeTD.style.outline = "";
          VIS.activeTD.style.outlineOffset = "";
        }
        VIS.rangeTDs.forEach(td => td.style.background = "");
      } catch {}
      VIS.activeTD = null;
      VIS.rangeTDs = [];
    }

    function inRect(r, c, r1, c1, r2, c2){
      const rr1 = Math.min(r1, r2), rr2 = Math.max(r1, r2);
      const cc1 = Math.min(c1, c2), cc2 = Math.max(c1, c2);
      return r >= rr1 && r <= rr2 && c >= cc1 && c <= cc2;
    }

    function getGridBody(){
      return API.getGridBody?.() || null;
    }

    function getTDByRC(r, c){
      const gridBody = getGridBody();
      if (!gridBody) return null;
      return gridBody.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
    }

    function selectionIsSingleCell(){
      const s = API.Sel?.start, e = API.Sel?.end;
      if (!s || !e) return false;
      return (Number(s.r) === Number(e.r) && Number(s.c) === Number(e.c));
    }

    function getSelectedTDsByDOM(){
      const s = API.Sel?.start, e = API.Sel?.end;
      const gridBody = getGridBody();
      if (!gridBody || !s || !e) return [];

      const r1 = Number(s.r), c1 = Number(s.c);
      const r2 = Number(e.r), c2 = Number(e.c);
      if (![r1,c1,r2,c2].every(Number.isFinite)) return [];

      const tds = Array.from(gridBody.querySelectorAll("td[data-r][data-c]"));
      return tds.filter(td => {
        const r = Number(td.dataset.r), c = Number(td.dataset.c);
        return Number.isFinite(r) && Number.isFinite(c) && inRect(r,c,r1,c1,r2,c2);
      });
    }

    function applyExcelVisual(){
      clearVisual();

      if (!API.Sel?.start || !API.Sel?.end) return;

      if (selectionIsSingleCell()){
        const r = Number(API.Sel.start.r);
        const c = Number(API.Sel.start.c);
        const td = getTDByRC(r, c);
        if (!td) return;

        td.style.outline = VIS.activeOutline;
        td.style.outlineOffset = "-2px";
        td.style.background = "";

        VIS.activeTD = td;
      } else {
        const tds = getSelectedTDsByDOM();
        tds.forEach(td => td.style.background = VIS.rangeBg);
        VIS.rangeTDs = tds;
      }
    }

    function safeApplySelectionDOM(){
      API.applySelectionDOM?.();
      applyExcelVisual();
      setTimeout(applyExcelVisual, 0);
      setTimeout(applyExcelVisual, 30);
    }

    // =========================================================
    //  Edit helpers
    // =========================================================
    function placeCaretAtEnd(td){
      try{
        const sel = window.getSelection?.();
        if (!sel) return;
        const range = document.createRange();
        range.selectNodeContents(td);
        range.collapse(false);
        sel.removeAllRanges();
        sel.addRange(range);
      } catch {}
    }

    function enterEditOnTD(td, evt){
      if (!td) return false;
      if (typeof window.enterEditMode !== "function") return false;
      window.enterEditMode(td, evt || {});
      setTimeout(() => placeCaretAtEnd(td), 0);
      setTimeout(() => placeCaretAtEnd(td), 50);
      return true;
    }

    function exitEdit(){
      if (typeof window.exitEditMode === "function") window.exitEditMode();
    }

    function moveSelectionDown(){
      const s = API.Sel?.start;
      if (!s) return;

      const r = Number(s.r), c = Number(s.c);
      if (!Number.isFinite(r) || !Number.isFinite(c)) return;

      const tdNext = getTDByRC(r + 1, c);
      if (!tdNext) return;

      API.setSelection?.({ r: r + 1, c }, { r: r + 1, c });
      safeApplySelectionDOM();
      clearCaret();
    }

    function replaceCellWithChar(td, ch, evt){
      enterEditOnTD(td, evt);
      setTimeout(() => {
        try{
          td.textContent = ch;
          placeCaretAtEnd(td);
        } catch {}
      }, 0);
    }

    // =========================================================
    // ✅ TSV helpers for COPY/CUT (FIX: multi-cell copy guaranteed)
    //    不用 API.selectionToTSV，直接依照 Sel.start/end 讀 DOM
    // =========================================================
    function selectionRect(){
      const s = API.Sel?.start, e = API.Sel?.end;
      if (!s || !e) return null;

      const r1 = Number(s.r), c1 = Number(s.c);
      const r2 = Number(e.r), c2 = Number(e.c);
      if (![r1,c1,r2,c2].every(Number.isFinite)) return null;

      const top = Math.min(r1, r2);
      const left = Math.min(c1, c2);
      const bottom = Math.max(r1, r2);
      const right = Math.max(c1, c2);

      return { top, left, bottom, right, rows: bottom-top+1, cols: right-left+1 };
    }

    function selectionToTSV_ByDOM(){
      const rect = selectionRect();
      if (!rect) return "";

      const lines = [];
      for (let r = rect.top; r <= rect.bottom; r++){
        const row = [];
        for (let c = rect.left; c <= rect.right; c++){
          const td = getTDByRC(r, c);
          row.push(String(td?.textContent ?? "").replace(/\r?\n/g, " "));
        }
        lines.push(row.join("\t"));
      }
      return lines.join("\n");
    }

    // =========================================================
    //  TSV helpers (IMPORTANT: always use pasteTSVAt)
    // =========================================================
    function normalizeClipboardText(s){
      return String(s ?? "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
    }

    function parseTSVToMatrix(text){
      const t = normalizeClipboardText(text);

      if (!(t.includes("\t") || t.includes("\n"))){
        return [[t]];
      }

      let lines = t.split("\n");
      if (lines.length > 1 && lines[lines.length - 1] === "") lines.pop();

      const rows = lines.map(line => line.split("\t"));
      if (!rows.length) return [[""]];
      return rows;
    }

    function matrixToTSV(mat){
      return mat.map(row => row.join("\t")).join("\n");
    }

    function tileMatrixToSize(src, outRows, outCols){
      const srcRows = src.length || 1;
      const srcCols = Math.max(1, ...src.map(r => r.length || 1));

      const out = [];
      for (let r = 0; r < outRows; r++){
        const row = [];
        for (let c = 0; c < outCols; c++){
          const v = (src[r % srcRows] && src[r % srcRows][c % srcCols] !== undefined)
            ? src[r % srcRows][c % srcCols]
            : "";
          row.push(String(v ?? ""));
        }
        out.push(row);
      }
      return out;
    }

    function doPasteTextExcelStyle(text){
      if (!API.hasSelection?.() || !API.Sel?.start) return;

      const rect = selectionRect();
      if (!rect) return;

      const clipMat = parseTSVToMatrix(text);

      const isSingleSel = (rect.rows === 1 && rect.cols === 1);

      if (isSingleSel){
        const t = normalizeClipboardText(text);
        API.pasteTSVAt?.(API.Sel.start.r, API.Sel.start.c, t);
        return;
      }

      const tiled = tileMatrixToSize(clipMat, rect.rows, rect.cols);
      const outTSV = matrixToTSV(tiled);
      API.pasteTSVAt?.(rect.top, rect.left, outTSV);
    }

    // =========================================================
    //  Drag state + DEBUG flag (set window.__SELECTION_DEBUG_DRAG = true to log)
    // =========================================================
    const DRAG = { down: false, moved: false, downAt: { x: 0, y: 0 }, pointerId: null };
    function markDown(e) {
      DRAG.down = true;
      DRAG.moved = false;
      DRAG.downAt = { x: e.clientX || 0, y: e.clientY || 0 };
      DRAG.pointerId = e.pointerId != null ? e.pointerId : null;
    }
    function markUp() {
      DRAG.down = false;
      DRAG.moved = false;
      DRAG.pointerId = null;
    }

    // =========================================================
    //  Events: pointerdown / pointermove / pointerup (robust range drag)
    // =========================================================

    document.addEventListener("pointerdown", (e) => {
      if (!(e.target instanceof Node)) return;

      const gridHead = document.getElementById("gridHead");
      const gridBody = getGridBody();

      // B3) Click column header (gridHead TH with data-col) -> select entire column
      if (gridHead && gridHead.contains(e.target)) {
        const th = e.target.closest("th");
        if (th && th.dataset.col !== undefined) {
          exitEdit();
          e.preventDefault();
          clearCaret();
          const s = (typeof window.activeSheet === "function") ? window.activeSheet() : null;
          if (s && typeof window.ensureSize === "function") {
            window.ensureSize(s);
            const col = Number(th.dataset.col);
            if (Number.isFinite(col) && col >= 0 && col < s.cols) {
              API.setSelection?.({ r: 0, c: col }, { r: s.rows - 1, c: col });
              safeApplySelectionDOM();
            }
          }
          clearCaret();
          return;
        }
      }

      // B3) Click row-number header (gridBody TH with data-row) -> select entire row
      if (gridBody && gridBody.contains(e.target)) {
        const th = e.target.closest("th");
        if (th && th.dataset.row !== undefined) {
          exitEdit();
          e.preventDefault();
          clearCaret();
          const s = (typeof window.activeSheet === "function") ? window.activeSheet() : null;
          if (s && typeof window.ensureSize === "function") {
            window.ensureSize(s);
            const row = Number(th.dataset.row);
            if (Number.isFinite(row) && row >= 0 && row < s.rows) {
              API.setSelection?.({ r: row, c: 0 }, { r: row, c: s.cols - 1 });
              safeApplySelectionDOM();
            }
          }
          clearCaret();
          return;
        }
      }

      // ---- TD-only logic (with pointer-based drag) ----
      if (!gridBody || !gridBody.contains(e.target)) return;

      const td = API.getTDFromEventTarget?.(e.target);
      if (!td) return;

      // Right-click: do not collapse selection if right-click inside selection
      if (e.button === 2) {
        const rect = selectionRect();
        if (rect && API.hasSelection?.()) {
          const r = Number(td.dataset.r), c = Number(td.dataset.c);
          if (Number.isFinite(r) && Number.isFinite(c) && r >= rect.top && r <= rect.bottom && c >= rect.left && c <= rect.right) {
            e.preventDefault();
            if (window.__SELECTION_DEBUG_DRAG) console.log("[selection] right-click inside selection, keep selection");
            return;
          }
        }
        return;
      }
      // Only primary button (0) or touch (-1) start selection
      if (e.button >= 1) return;

      const key = `${td.dataset.r},${td.dataset.c}`;
      if (window.__CELL_EDIT_MODE && window.__EDIT_CELL_KEY === key) return;

      // C) Force exit edit mode when selection begins
      exitEdit();

      if (window.__SELECTION_DEBUG_DRAG) console.log("[selection] pointerdown __CELL_EDIT_MODE=", !!window.__CELL_EDIT_MODE);

      e.preventDefault();
      clearCaret();

      if (e.shiftKey) {
        API.shiftClickSelect?.(td);
        API.Sel.isDown = false;
        safeApplySelectionDOM();
        clearCaret();
        return;
      }

      const r = Number(td.dataset.r);
      const c = Number(td.dataset.c);
      if (!Number.isFinite(r) || !Number.isFinite(c)) return;

      API.Sel.isDown = true;
      API.setSelection?.({ r, c }, { r, c });
      safeApplySelectionDOM();
      clearCaret();
      markDown(e);
    }, true);

    // pointermove: use elementFromPoint to resolve TD under pointer (works when moving fast or outside grid)
    document.addEventListener("pointermove", (e) => {
      if (!API.Sel?.isDown) return;
      if (DRAG.pointerId != null && e.pointerId !== DRAG.pointerId) return;

      const gridBody = getGridBody();
      const el = document.elementFromPoint(e.clientX, e.clientY);
      const td = el ? API.getTDFromEventTarget?.(el) : null;
      if (!td || !gridBody || !gridBody.contains(td)) return;

      const r = Number(td.dataset.r);
      const c = Number(td.dataset.c);
      if (!Number.isFinite(r) || !Number.isFinite(c)) return;

      const start = API.Sel?.start;
      if (!start) return;

      if (typeof API.setSelection === "function") {
        API.setSelection({ r: Number(start.r), c: Number(start.c) }, { r, c });
      } else {
        API.Sel.end = { r, c };
      }
      safeApplySelectionDOM();
      clearCaret();

      if (window.__SELECTION_DEBUG_DRAG) console.log("[selection] pointermove end cell", r, c);
    }, true);

    document.addEventListener("pointerup", (e) => {
      if (DRAG.pointerId != null && e.pointerId !== DRAG.pointerId) return;
      API.Sel.isDown = false;
      markUp();
      clearCaret();
    }, true);

    document.addEventListener("dblclick", (e) => {
      const gridBody = getGridBody();
      if (!gridBody || !(e.target instanceof Node) || !gridBody.contains(e.target)) return;

      const td = API.getTDFromEventTarget?.(e.target);
      if (!td) return;

      e.preventDefault();

      const r = Number(td.dataset.r);
      const c = Number(td.dataset.c);
      if (Number.isFinite(r) && Number.isFinite(c)){
        API.setSelection?.({r,c}, {r,c});
        safeApplySelectionDOM();
      }

      enterEditOnTD(td, e);
    }, true);

    document.addEventListener("keydown", (e) => {
      const ae = document.activeElement;
      if (ae && (ae.tagName === "INPUT" || ae.tagName === "TEXTAREA")) return;

      if (window.__CELL_EDIT_MODE){
        if (e.key === "Enter"){
          e.preventDefault();
          exitEdit();
          moveSelectionDown();
        }
        return;
      }

      if (e.key === "Enter"){
        if (!API.hasSelection?.()) return;
        e.preventDefault();
        moveSelectionDown();
        return;
      }

      // ---- BLOCK 12_EVENTS_KEYDOWN: Excel-like keyboard selection ----
      // A1) Ctrl+A / Cmd+A: select all
      if ((e.ctrlKey || e.metaKey) && (e.key === "a" || e.key === "A")) {
        e.preventDefault();
        const s = (typeof window.activeSheet === "function") ? window.activeSheet() : null;
        if (s) {
          if (typeof window.ensureSize === "function") window.ensureSize(s);
          API.setSelection?.({ r: 0, c: 0 }, { r: s.rows - 1, c: s.cols - 1 });
          safeApplySelectionDOM();
        }
        clearCaret();
        return;
      }

      // A2) Shift+Arrow: extend selection (anchor stays, end moves)
      if (e.shiftKey && ["ArrowUp","ArrowDown","ArrowLeft","ArrowRight"].includes(e.key)) {
        e.preventDefault();
        const s = (typeof window.activeSheet === "function") ? window.activeSheet() : null;
        if (!s) return;
        if (typeof window.ensureSize === "function") window.ensureSize(s);
        let anchor = API.Sel?.start;
        if (!anchor) {
          const af = document.activeElement;
          if (af && af.tagName === "TD" && af.dataset.r != null && af.dataset.c != null) {
            anchor = { r: Number(af.dataset.r), c: Number(af.dataset.c) };
          } else {
            anchor = { r: 0, c: 0 };
          }
        } else {
          anchor = { r: Number(anchor.r), c: Number(anchor.c) };
        }
        let end = API.Sel?.end;
        if (!end) end = { r: anchor.r, c: anchor.c };
        else end = { r: Number(end.r), c: Number(end.c) };
        if (e.key === "ArrowUp") end.r = Math.max(0, end.r - 1);
        if (e.key === "ArrowDown") end.r = Math.min(s.rows - 1, end.r + 1);
        if (e.key === "ArrowLeft") end.c = Math.max(0, end.c - 1);
        if (e.key === "ArrowRight") end.c = Math.min(s.cols - 1, end.c + 1);
        API.setSelection?.(anchor, end);
        safeApplySelectionDOM();
        clearCaret();
        return;
      }

      // A4) Ctrl+Space: select entire column
      if (e.key === " " && (e.ctrlKey || e.metaKey) && !e.shiftKey) {
        e.preventDefault();
        const s = (typeof window.activeSheet === "function") ? window.activeSheet() : null;
        if (!s) return;
        if (typeof window.ensureSize === "function") window.ensureSize(s);
        let col = API.Sel?.start?.c;
        if (col == null) {
          const af = document.activeElement;
          if (af && af.tagName === "TD" && af.dataset.c != null) col = Number(af.dataset.c);
          else col = 0;
        } else {
          col = Number(col);
        }
        if (!Number.isFinite(col) || col < 0) col = 0;
        if (col >= s.cols) col = s.cols - 1;
        API.setSelection?.({ r: 0, c: col }, { r: s.rows - 1, c: col });
        safeApplySelectionDOM();
        clearCaret();
        return;
      }

      // A5) Shift+Space: select entire row
      if (e.key === " " && e.shiftKey && !(e.ctrlKey || e.metaKey)) {
        e.preventDefault();
        const s = (typeof window.activeSheet === "function") ? window.activeSheet() : null;
        if (!s) return;
        if (typeof window.ensureSize === "function") window.ensureSize(s);
        let row = API.Sel?.start?.r;
        if (row == null) {
          const af = document.activeElement;
          if (af && af.tagName === "TD" && af.dataset.r != null) row = Number(af.dataset.r);
          else row = 0;
        } else {
          row = Number(row);
        }
        if (!Number.isFinite(row) || row < 0) row = 0;
        if (row >= s.rows) row = s.rows - 1;
        API.setSelection?.({ r: row, c: 0 }, { r: row, c: s.cols - 1 });
        safeApplySelectionDOM();
        clearCaret();
        return;
      }

      const kl = (e.key || "").toLowerCase();
      if (kl === "delete" || kl === "backspace"){
        if (!API.hasSelection?.()) return;
        e.preventDefault();
        API.clearSelectedCells?.();
        safeApplySelectionDOM();
        clearCaret();
        return;
      }

      const k = e.key || "";
      const isPrintable = (k.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey);
      if (isPrintable){
        if (!API.hasSelection?.() || !API.Sel?.start) return;
        if (!selectionIsSingleCell()) return;

        const r = Number(API.Sel.start.r);
        const c = Number(API.Sel.start.c);
        if (!Number.isFinite(r) || !Number.isFinite(c)) return;

        const td = getTDByRC(r, c);
        if (!td) return;

        e.preventDefault();
        replaceCellWithChar(td, k, e);
        return;
      }
    }, true);

    // ✅ Copy/Cut/Paste（大量 cell）
    document.addEventListener("copy", (e) => {
      if (window.__CELL_EDIT_MODE) return;
      if (!API.hasSelection?.()) return;

      // ✅ FIX #2：用 DOM 直接組 TSV，保證多格一定會複製到
      const tsv = selectionToTSV_ByDOM();
      if (!tsv) return;

      e.preventDefault();
      e.clipboardData?.setData("text/plain", tsv);
    }, true);

    document.addEventListener("cut", (e) => {
      if (window.__CELL_EDIT_MODE) return;
      if (!API.hasSelection?.()) return;

      const tsv = selectionToTSV_ByDOM();
      if (!tsv) return;

      e.preventDefault();
      e.clipboardData?.setData("text/plain", tsv);

      setTimeout(() => {
        API.clearSelectedCells?.();
        if (typeof window.render === "function") window.render();
        safeApplySelectionDOM();
        clearCaret();
      }, 0);
    }, true);

    document.addEventListener("paste", (e) => {
      if (window.__CELL_EDIT_MODE) return;
      if (!API.hasSelection?.() || !API.Sel?.start) return;

      e.preventDefault();

      const text = e.clipboardData?.getData("text") ?? "";
      doPasteTextExcelStyle(text);

      setTimeout(() => {
        if (typeof window.render === "function") window.render();
        safeApplySelectionDOM();
        clearCaret();
      }, 0);
    }, true);

    setTimeout(() => {
      safeApplySelectionDOM();
      clearCaret();
    }, 0);

    return true;
  }

  function bind(){
    if (window.__SELECTION_EVENTS_BINDED__) return true;
    if (tryBindOnce()) return true;

    let tries = 0;
    const timer = setInterval(() => {
      tries++;
      if (tryBindOnce()) clearInterval(timer);
      else if (tries >= 50){
        clearInterval(timer);
        console.warn("⚠️ SELECTION_EVENTS bind timeout: SELECTION_CORE api still not ready");
      }
    }, 100);

    return false;
  }

  window.DEFS.SELECTION_EVENTS.bind = bind;

})();
