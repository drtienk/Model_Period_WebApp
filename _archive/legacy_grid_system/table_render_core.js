console.log("✅ [table_render_core.js] loaded");

window.DEFS = window.DEFS || {};
window.DEFS.TABLE_RENDER = window.DEFS.TABLE_RENDER || {};

(function installTableRenderCore(){

  function safeSave(){
    if (typeof window.saveToLocalByMode === "function") {
      window.saveToLocalByMode(window.activeMode);
    }
  }

  function makeHeaderInput(value, onChange) {
    const inp = document.createElement("input");
    inp.type = "text";
    inp.className = "th-input";
    inp.value = value ?? "";
    inp.addEventListener("input", () => onChange(inp.value));
    return inp;
  }

  function renderHeaderDefault(s, gridHead) {
    gridHead.innerHTML = "";
    const trH = document.createElement("tr");
    const thCorner = document.createElement("th");
    thCorner.textContent = "";
    thCorner.dataset.corner = "1";
    trH.appendChild(thCorner);

    for (let c=0; c<s.cols; c++) {
      const th = document.createElement("th");
      th.dataset.col = String(c);
      let display = String((s.headers[c] ?? "")).trim();
      if (display === "") display = "Col " + (c+1);

      th.appendChild(makeHeaderInput(display, (v)=>{
        s.headers[c] = String(v ?? "");
        safeSave();
      }));
      trH.appendChild(th);
    }
    gridHead.appendChild(trH);
  }

  function renderHeaderSAC(s, gridHead) {
    if (typeof window.ensureDafMeta === "function") window.ensureDafMeta(s);
    gridHead.innerHTML = "";

    const tr1 = document.createElement("tr");
    const tr2 = document.createElement("tr");
    const tr3 = document.createElement("tr");

    const thCorner = document.createElement("th");
    thCorner.rowSpan = 3;
    thCorner.textContent = "";
    thCorner.dataset.corner = "1";
    tr1.appendChild(thCorner);

    for (let c=0; c<3; c++) {
      const th = document.createElement("th");
      th.rowSpan = 3;
      th.dataset.col = String(c);
      th.textContent = String(s.headers[c] ?? ("Col " + (c+1)));
      tr1.appendChild(th);
    }

    for (let c=3; c<s.cols; c++) {
      const thTop = document.createElement("th");
      thTop.dataset.col = String(c);
      thTop.appendChild(makeHeaderInput(
        String(s.headers[c] ?? "Fill your Support Activity Center"),
        (v)=>{ s.headers[c]=v; safeSave(); }
      ));
      tr1.appendChild(thTop);

      const thDesc = document.createElement("th");
      thDesc.dataset.col = String(c);
      thDesc.appendChild(makeHeaderInput(
        String(s?.meta?.dafDesc?.[c] ?? "Description"),
        (v)=>{ s.meta.dafDesc[c]=v; safeSave(); }
      ));
      tr2.appendChild(thDesc);

      const thEnt = document.createElement("th");
      thEnt.dataset.col = String(c);
      thEnt.appendChild(makeHeaderInput(
        String(s?.meta?.dafEnt?.[c] ?? "EDU"),
        (v)=>{ s.meta.dafEnt[c]=v; safeSave(); }
      ));
      tr3.appendChild(thEnt);
    }

    gridHead.appendChild(tr1);
    gridHead.appendChild(tr2);
    gridHead.appendChild(tr3);
  }

  function render() {
    const gridHead = document.getElementById("gridHead");
    const gridBody = document.getElementById("gridBody");
    if (!gridHead || !gridBody) return;

    // ✅ if helper not ready, don't crash
    const ensureSize = window.ensureSize;
    if (typeof ensureSize !== "function") return;

    const s = (typeof window.activeSheet === "function") ? window.activeSheet() : null;
    if (!s) return;

    if (typeof window.ensureHeadersForActiveSheet === "function") window.ensureHeadersForActiveSheet();
    ensureSize(s);

    if (window.activeMode === "period" && window.activeKey === "daf") renderHeaderSAC(s, gridHead);
    else renderHeaderDefault(s, gridHead);

    gridBody.innerHTML = "";
    for (let r=0; r<s.rows; r++) {
      const tr = document.createElement("tr");
      const rowTh = document.createElement("th");
      rowTh.textContent = String(r+1);
      rowTh.dataset.row = String(r);
      tr.appendChild(rowTh);

      for (let c=0; c<s.cols; c++) {
        const td = document.createElement("td");
        td.contentEditable = "true";
        td.spellcheck = false;
        td.textContent = s.data?.[r]?.[c] ?? "";
        td.dataset.r = String(r);
        td.dataset.c = String(c);
        tr.appendChild(td);
      }
      gridBody.appendChild(tr);
    }

    if (typeof window.applySelectionDOM === "function") window.applySelectionDOM();
  }

  // expose
  window.DEFS.TABLE_RENDER.safeSave = safeSave;
  window.DEFS.TABLE_RENDER.makeHeaderInput = makeHeaderInput;
  window.DEFS.TABLE_RENDER.renderHeaderDefault = renderHeaderDefault;
  window.DEFS.TABLE_RENDER.renderHeaderSAC = renderHeaderSAC;
  window.DEFS.TABLE_RENDER.render = render;

})();
