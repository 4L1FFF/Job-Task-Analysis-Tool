/* =========================================================
 JOB TASK ANALYSIS Tool — app.js (CAT theme + Noto Sans)
 v6 — Clean, validated, syntax-safe — Updated with new UI elements
 ========================================================= */
document.addEventListener("DOMContentLoaded", () => {
  /* ========= Helpers ========= */
  const on = (id, ev, fn) => { const el = document.getElementById(id); if (el) el.addEventListener(ev, fn); return el; };
  const $  = (s) => document.querySelector(s);
  const $$ = (s) => Array.from(document.querySelectorAll(s));
  const esc = (s) => String(s).replace(/[&<>"]/g,c=>({'&':'&','<':'<','>':'>'}[c]||c));
  const norm = (v) => String(v ?? "").trim();
  const lower= (v) => norm(v).toLowerCase();
  const unique= (a) => [...new Set(a)];
  const cssEscape = (s) => (window.CSS && CSS.escape) ? CSS.escape(s) : String(s).replace(/[^a-zA-Z0-9_\-]/g, m => '\\' + m);

  /* ========= CONFIG ========= */
  const CONFIG = {
    STORAGE_KEY: "jtat_ui_state_v6",
    TIME_KEY: "jtat_time_settings_v4",
    PARTICIPANTS_KEY: "jtat_participants",
    SKILL_SNAPSHOT_KEY: "jtat_skill_snapshot",

    CASCADES: {
      JobRole: {
        filters: ["Job Title","Level","Industry/Segment","Job Role","Duty"],
        output : ["Skills","Learning Plan Code","Learning Plan Name","Course Number","Course Name","SMCS Code"]
      },
      System: {
        filters: ["Job Title","Level","System","Duty","Task"],
        output : ["Skills","Learning Plan Code","Learning Plan Name","SMCS Code"]
      },
      Product: {
        filters: ["Job Title","Level","Product","Duty","Task"],
        output : ["Skills","Learning Plan Code","Learning Plan Name","SMCS Code"]
      }
    },

    MAIN_COLUMNS: ["Learning Plan Code","Learning Plan Name","built-in enrollment attribute"],
    PREFERRED_SKILL_COLUMNS: ["Job Role","Industry/Segment","Level","Skills","SMCS Code","SMCS Job","Task"],

    DISPLAY_NAME: { "built-in enrollment attribute": "enrollment attribute" },

    LEVEL_TIPS: {
      needF : "Make sure the Foundational level is completed.",
      needFA: "Make sure Foundational and Advanced levels are completed."
    },

    CSV: {
      header: ["username","user ID","learning plan ID","learning plan uuid","timezone","active from","active until","built-in enrollment attribute"],
      timezoneDefault: "Europe/Rome",
      filenamePrefix: "enrollments_multi_learning_plan_"
    }
  };

  /* ========= Diagnostics ========= */
  const DIAG = { rowsIn: 0, rowsFiltered: 0, processingMs: 0, lastUpdate: null };
  function renderDiagnostics(){
    const box = document.getElementById("diagPanel"); if (!box) return;
    const ts = DIAG.lastUpdate ? new Date(DIAG.lastUpdate).toLocaleString() : "—";
    box.innerHTML = `
      <div class="diagnostics">
        <h4>Diagnostics</h4>
        <div class="diagnostics__grid">
          <div class="diagnostics__key">Rows in cascade</div><div>${DIAG.rowsIn}</div>
          <div class="diagnostics__key">Rows after filter/search</div><div>${DIAG.rowsFiltered}</div>
          <div class="diagnostics__key">Processing time</div><div>${DIAG.processingMs} ms</div>
          <div class="diagnostics__key">Last update</div><div>${ts}</div>
        </div>
      </div>`;
  }

  /* ========= Participants counters ========= */
  function loadParticipantsLS(){ try { return JSON.parse(localStorage.getItem(CONFIG.PARTICIPANTS_KEY) || "[]"); } catch { return []; } }
  function renderParticipantsCounters(){
    const list = loadParticipantsLS();
    const total = list.length;
    const selected = list.filter(p => p && p.active).length;
    [
      ["top_total", `Total: ${total}`],
      ["top_selected", `Selected: ${selected}`],
      ["adm_total", String(total)],
      ["adm_selected", String(selected)]
    ].forEach(([id, txt]) => { const el = document.getElementById(id); if (el) el.textContent = txt; });
  }
  function openParticipantsTab(){ window.open("data/participants.html","_blank","noopener,noreferrer"); }
  on("btnOpenParticipants","click", openParticipantsTab);
  on("btnOpenParticipantsTop","click", openParticipantsTab);
  on("btnOpenParticipantsAdmin","click", openParticipantsTab);
  renderParticipantsCounters();
  window.addEventListener("storage",(e)=>{ if (e.key===CONFIG.PARTICIPANTS_KEY) renderParticipantsCounters(); });

  /* ========= Admin / DB ========= */
  let DB=null, activeCascade="JobRole", currentRows=[], displayedRows=[];
  function normalizeCascadeName(name){
    const s=String(name).replace(/\s+/g,"").toLowerCase();
    if(s==="jobrole")return"JobRole"; if(s==="system")return"System"; if(s==="product")return"Product"; return null;
  }
  function renderDBSummary(){
    const box=document.getElementById("dbSummary"); if(!box) return;
    if(!DB){ box.textContent=""; return; }
    let out="";
    Object.keys(DB).forEach(k=>{
      const it=DB[k];
      out += `${k}: ${it.rows.length} rows\nFilters: ${it.filters.join(", ")}\nOutput: ${it.output.join(", ")}\n\n`;
    });
    box.textContent = out;
  }
  on("btnProcess","click", async ()=>{
    const f=document.getElementById("excelInput")?.files?.[0]; if(!f) return alert("Choose an Excel file.");
    const adminLog=document.getElementById("adminLog"); if (adminLog) adminLog.textContent="Reading Excel…";
    try{
      const t0 = performance.now();
      const buf = await f.arrayBuffer();
      const wb = XLSX.read(buf);
      const db = {};
      wb.SheetNames.forEach(sn=>{
        const nn = normalizeCascadeName(sn);
        if(!nn || !CONFIG.CASCADES[nn]) return;
        db[nn] = {
          filters: CONFIG.CASCADES[nn].filters.slice(),
          output : CONFIG.CASCADES[nn].output.slice(),
          rows   : XLSX.utils.sheet_to_json(wb.Sheets[sn], { defval:"" })
        };
      });
      DB=db; if (adminLog) adminLog.textContent="Database built successfully."; renderDBSummary();
      DIAG.rowsIn = DB?.[activeCascade]?.rows?.length || 0;
      DIAG.processingMs = Math.round(performance.now()-t0);
      DIAG.lastUpdate = Date.now(); renderDiagnostics();
    }catch(err){ console.error(err); if (adminLog) adminLog.textContent="Error while reading Excel. See console."; }
  });
  on("btnSaveDB","click", ()=>{
    if(!DB) return alert("No DB loaded.");
    const blob=new Blob([JSON.stringify(DB,null,2)],{type:"application/json"});
    const a=document.createElement("a");
    a.href=URL.createObjectURL(blob); a.download="jtat_database.json"; a.click();
  });
  on("btnLoadDB","click", async ()=>{
    const f=document.getElementById("dbInput")?.files?.[0]; if(!f) return alert("Pick a JSON file.");
    try{ const text=await f.text(); DB=JSON.parse(text); renderDBSummary(); initUserPanel();
      DIAG.rowsIn = DB?.[activeCascade]?.rows?.length || 0;
      DIAG.lastUpdate = Date.now(); renderDiagnostics();
    } catch(e){ alert("Invalid JSON: "+e.message); }
  });

  /* ========= Filters / Table ========= */
  function highlight(text,q){ if(!q) return esc(text); const re = new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g,"\\$&"),"ig"); return esc(text).replace(re, m=>`<mark>${esc(m)}</mark>`); }
  function dedupeByLearningPlan(rows){
    const seen=new Set(), out=[]; for(const r of rows){
      const code=norm(r["Learning Plan Code"]); if(!code || seen.has(code)) continue;
      seen.add(code); out.push(r);
    } return out;
  }
  function renderFiltersForCascade(cascade){
    const box=document.getElementById("filtersBox"); if (!box||!DB) return;
    const order=CONFIG.CASCADES[cascade].filters; const rows=DB[cascade].rows; let html="";
    order.forEach((f)=>{
      const id=`f_${cascade}_${f.replace(/[^a-z0-9_\-]/gi,"_")}`;
      html +=
      `<div class="facet-box" data-filter-wrap="${esc(f)}">
         <div class="facet-title">${esc(f)}</div>
         <div class="facet-options" data-checkboxes-for="${esc(id)}"></div>
         <select id="${id}" data-filter="${esc(f)}" multiple class="hidden" aria-hidden="true"></select>
       </div>`;
    });
    box.innerHTML = html;
    populateCascadeOptions(cascade, rows);
    updateGatedVisibility(cascade);
    checkLevelPrereqs();
  }
function populateCascadeOptions(cascade, scope){
  for(const f of CONFIG.CASCADES[cascade].filters){
    const id=`f_${cascade}_${f.replace(/[^a-z0-9_\-]/gi,"_")}`;
    const sel=document.getElementById(id);
    const wrap=document.querySelector(`[data-checkboxes-for="${cssEscape(id)}"]`);
    if(!sel || !wrap) continue;
    sel.innerHTML=""; wrap.innerHTML="";
    
    // Get unique values and sort them
    const vals=unique(scope.map(r=>r[f]).filter(v=>norm(v)!=="")).sort((a,b)=>String(a).localeCompare(String(b)));
    
    // Create a dropdown/select menu instead of checkboxes
    const dropdown = document.createElement('select');
    dropdown.className = 'filter-dropdown';
    dropdown.setAttribute('data-filter', f);
    dropdown.setAttribute('multiple', true);
    dropdown.style.minHeight = '120px';
    dropdown.style.width = '100%';
    dropdown.style.padding = '8px';
    
    // Add all values as options
    vals.forEach(v=>{
      const opt = document.createElement('option');
      opt.value = v;
      opt.textContent = v;
      dropdown.appendChild(opt);
    });
    
    // Replace the hidden select with this visible one
    sel.parentNode.replaceChild(dropdown, sel);
    dropdown.id = id;
    
    // Handle selection changes
    dropdown.addEventListener('change', () => {
      updateDependentFilters(cascade, f);
      updateGatedVisibility(cascade);
      saveUIState();
      if (f === "Level") checkLevelPrereqs();
    });
  }
}

  function updateDependentFilters(cascade, changedFilterName){
    const order = CONFIG.CASCADES[cascade].filters;
    const changedIdx = order.indexOf(changedFilterName);
    if (changedIdx < 0) return;
    const rows = DB[cascade].rows;
    const constrained = rows.filter(row => {
      for (let i = 0; i <= changedIdx; i++) {
        const fName = order[i];
        const sel = document.querySelector(`#filtersBox select[data-filter="${cssEscape(fName)}"]`);
        const chosen = sel ? Array.from(sel.selectedOptions).map(o => o.value) : [];
        if (chosen.length && !chosen.includes(row[fName])) return false;
      } return true;
    });
    for (let j = changedIdx + 1; j < order.length; j++) {
      const fName = order[j];
      const sel = document.querySelector(`#filtersBox select[data-filter="${cssEscape(fName)}"]`);
      if (!sel) continue;
      const wrap = document.querySelector(`[data-checkboxes-for="${cssEscape(sel.id)}"]`);
      if (!wrap) continue;
      const prev = Array.from(sel.selectedOptions).map(o => o.value);
      sel.innerHTML = ""; wrap.innerHTML = "";
      const vals = unique(constrained.map(r => r[fName]).filter(v => norm(v) !== "")).sort((a,b)=>String(a).localeCompare(String(b)));
      vals.forEach(v=>{
        const opt = document.createElement("option"); opt.value = v; opt.textContent = v; sel.appendChild(opt);
        const lab = document.createElement("label");
        lab.className = "facet-option";
        lab.innerHTML = `<input type="checkbox" class="facet-cb" value="${esc(v)}"> <span>${esc(v)}</span>`;
        const cb = lab.querySelector("input"); cb.checked = prev.includes(v);
        cb.addEventListener("change", ()=>{
          Array.from(sel.options).forEach(o => { if (o.value === v) o.selected = cb.checked; });
          updateDependentFilters(cascade, fName);
          updateGatedVisibility(cascade);
          saveUIState();
          if (fName === "Level") checkLevelPrereqs();
        });
        wrap.appendChild(lab);
      });
      Array.from(sel.options).forEach(o => o.selected = prev.includes(o.value));
    }
    checkLevelPrereqs();
  }
  function updateGatedVisibility(cascade){
    const order = CONFIG.CASCADES[cascade].filters;
    let canShow = getSelectedCount(order[0]) > 0;
    for (let i = 1; i < order.length; i++) {
      const fname = order[i];
      toggleWrap(fname, canShow);
      if (canShow) canShow = getSelectedCount(fname) > 0;
    }
    function toggleWrap(fname, show){
      const el = document.querySelector(`#filtersBox [data-filter-wrap="${cssEscape(fname)}"]`);
      if (!el) return; el.classList.toggle("hidden", !show);
    }
    function getSelectedCount(fname){
      const sel = document.querySelector(`#filtersBox select[data-filter="${cssEscape(fname)}"]`);
      return sel ? Array.from(sel.selectedOptions).length : 0;
    }
    checkLevelPrereqs();
  }
  function gatherFilteredRows(cascade){
    const order=CONFIG.CASCADES[cascade].filters; const rows=DB[cascade].rows;
    return rows.filter(row=>{
      for(const fName of order){
        const sel=document.querySelector(`#filtersBox select[data-filter="${cssEscape(fName)}"]`);
        const chosen = sel ? Array.from(sel.selectedOptions).map(o=>o.value) : [];
        if(chosen.length && !chosen.includes(row[fName])) return false;
      } return true;
    });
  }
  function renderTable(rows, outCols, q){
    const tbl=document.getElementById("resultTable"); if (!tbl) return;
    const cols=CONFIG.MAIN_COLUMNS;
    const thead="<thead><tr>"+cols.map(c=>`<th>${esc(CONFIG.DISPLAY_NAME[c]||c)}</th>`).join("")+"</tr></thead>";
    const tbody="<tbody>"+rows.map(r=>"<tr>"+cols.map(c=>`<td>${highlight(String(r[c]??""),q)}</td>`).join("")+"</tr>").join("")+"</tbody>";
    tbl.innerHTML = thead+tbody;
  }
  function updateResultCounters(){
    const resCount = document.getElementById("resultCount");
    if (resCount) resCount.textContent = `${displayedRows.length} learning plans`;
    const attrSummary = document.getElementById("attrSummary");
    if (!attrSummary) return;
    const attrs = unique(displayedRows.map(r=>norm(r["built-in enrollment attribute"]).toLowerCase()).filter(Boolean));
    const cfg = loadTimeCfg();
    const configured = attrs.filter(a => {
      const d = (cfg.perAttr||{})[a];
      return d && (d.from || d.until);
    });
    attrSummary.textContent = attrs.length ? `Attributes: ${attrs.length} · with dates: ${configured.length}` : `Attributes: 0`;
  }
  const debounced=(fn,d=250)=>{let t;return(...a)=>{clearTimeout(t);t=setTimeout(()=>fn(...a),d)}};
  const applySearch = debounced(()=>{
    if(!DB) return;
    const t0 = performance.now();
    const q = lower(document.getElementById("searchInput")?.value || "");
    const cols = CONFIG.CASCADES[activeCascade].output;
    const filtered = q ? currentRows.filter(r=>cols.some(c=>String(r[c]||"").toLowerCase().includes(q))) : currentRows.slice();
    displayedRows = dedupeByLearningPlan(filtered);
    renderTable(displayedRows, cols, q);
    updateResultCounters();
    DIAG.rowsIn = DB?.[activeCascade]?.rows?.length || 0;
    DIAG.rowsFiltered = displayedRows.length;
    DIAG.processingMs = Math.round(performance.now()-t0);
    DIAG.lastUpdate = Date.now(); renderDiagnostics();
  },250);

  /* ========= UI controls ========= */
  document.querySelectorAll('input[name="cascade"]').forEach((el) => {
    el.addEventListener("change", (e) => {
      if (!DB) return;
      const next = String(e?.target?.value || "").trim();
      if (!next || !CONFIG.CASCADES[next]) return;
      activeCascade = next;
      renderFiltersForCascade(activeCascade);
      currentRows = gatherFilteredRows(activeCascade);
      applySearch();
      saveUIState();
    });
  });
  on("btnApply","click", ()=>{ if(!DB) return; currentRows = gatherFilteredRows(activeCascade); applySearch(); saveUIState(); });
  on("btnReset","click", ()=>{ if(!DB) return; renderFiltersForCascade(activeCascade); const s=document.getElementById("searchInput"); if (s) s.value=""; localStorage.removeItem(CONFIG.STORAGE_KEY); currentRows = DB[activeCascade].rows.slice(); applySearch(); });
  on("searchInput","input", applySearch);
  on("btnClearSearch","click", ()=>{ const s=document.getElementById("searchInput"); if (s) s.value=""; applySearch(); saveUIState(); });

  /* ========= New UI Elements from screenshot ========= */
  // Connect the new buttons to appropriate functionality
  
  // File Search button - could open a search panel or focus on search
  on("btnFileSearch", "click", () => {
    // Switch to User tab if needed
    switchTab("User");
    // Focus on search input
    setTimeout(() => document.getElementById("searchInput")?.focus(), 100);
  });
  
  // File Upload button - could trigger Excel upload
  on("btnFileUpload", "click", () => {
    // Switch to Admin tab
    switchTab("Admin");
    // Trigger file upload click
    document.getElementById("excelInput")?.click();
  });
  
  // Activity Log button - could show diagnostics or logs
  on("btnActivity", "click", () => {
    // Scroll to diagnostics panel
    document.getElementById("diagPanel")?.scrollIntoView({ behavior: "smooth" });
    // Flash the diagnostics panel to draw attention
    const diag = document.getElementById("diagPanel");
    if (diag) {
      diag.style.transition = "background-color 0.5s";
      diag.style.backgroundColor = "#fff7d1";
      setTimeout(() => diag.style.backgroundColor = "", 500);
    }
  });
  
  // Advanced Search link
  document.querySelector(".advanced-search")?.addEventListener("click", (e) => {
    e.preventDefault();
    // Expand/collapse advanced filters - in your case, switch to User panel with all filters visible
    switchTab("User");
    // Make sure all filters are visible
    const filtersBox = document.getElementById("filtersBox");
    if (filtersBox) {
      filtersBox.querySelectorAll('.facet-box').forEach(el => el.classList.remove('hidden'));
    }
  });
  
  // Reset button
  document.querySelector(".reset-btn")?.addEventListener("click", () => {
    // Clear filters and search
    if (DB) {
      renderFiltersForCascade(activeCascade);
      const s = document.getElementById("searchInput");
      if (s) s.value = "";
      currentRows = DB[activeCascade].rows.slice();
      applySearch();
    }
  });

  function saveUIState(){
    if(!DB) return;
    const st={activeCascade, selections:{}, search: document.getElementById("searchInput")?.value || ""};
    Object.keys(CONFIG.CASCADES).forEach(c=>{
      st.selections[c]={};
      CONFIG.CASCADES[c].filters.forEach(f=>{
        const el=document.querySelector(`#filtersBox select[data-filter="${cssEscape(f)}"]`);
        st.selections[c][f] = el ? Array.from(el.selectedOptions).map(o=>o.value) : [];
      });
    });
    localStorage.setItem(CONFIG.STORAGE_KEY, JSON.stringify(st));
  }
  function loadUIState(){ try{ const raw=localStorage.getItem(CONFIG.STORAGE_KEY); return raw?JSON.parse(raw):null; }catch{ return null; } }
  function restoreSelections(cascade, sel){
    CONFIG.CASCADES[cascade].filters.forEach(f=>{
      const id=`f_${cascade}_${f.replace(/[^a-z0-9_\-]/gi,"_")}`;
      const el=document.getElementById(id);
      const wrap=document.querySelector(`[data-checkboxes-for="${cssEscape(id)}"]`);
      if(!el || !wrap) return;
      const chosen=Array.isArray(sel[f])?sel[f]:[];
      Array.from(el.options).forEach(o=>o.selected = chosen.includes(o.value));
      wrap.querySelectorAll('input[type="checkbox"]').forEach(cb=>{ cb.checked = chosen.includes(cb.value); });
    });
  }

  /* ========= Level tips ========= */
  function checkLevelPrereqs(){
    if (!DB?.[activeCascade]) return;
    const levelSel = document.querySelector(`#filtersBox select[data-filter="${cssEscape("Level")}"]`);
    const levelWrap = document.querySelector(`#filtersBox [data-filter-wrap="${cssEscape("Level")}"]`);
    if (!levelSel || !levelWrap) return;
    const chosen = Array.from(levelSel.selectedOptions).map(o => (o.value || "").toLowerCase());
    let tip = levelWrap.querySelector(".level-tip");
    if (tip) tip.remove();
    if (!chosen.length) return;
    const hasF = chosen.includes("foundational");
    const hasA = chosen.includes("advanced");
    const hasE = chosen.includes("expert");
    let msg = "";
    if (hasE && (!hasF || !hasA)) msg = CONFIG.LEVEL_TIPS.needFA;
    else if (hasA && !hasF) msg = CONFIG.LEVEL_TIPS.needF;
    if (msg) {
      tip = document.createElement("div");
      tip.className = "level-tip";
      tip.textContent = msg;
      levelWrap.appendChild(tip);
    }
  }

  function initUserPanel(){
    const fb=document.getElementById("filtersBox");
    const rt=document.getElementById("resultTable");
    const rc=document.getElementById("resultCount");
    const as=document.getElementById("attrSummary");
    if(!DB){
      if (fb) fb.innerHTML = "<div class='muted'>No DB loaded.</div>";
      if (rt) rt.innerHTML = "";
      if (rc) rc.textContent = "";
      if (as) as.textContent = "—";
      return;
    }
    const saved = loadUIState();
    if(saved?.activeCascade && CONFIG.CASCADES[saved.activeCascade]) activeCascade = saved.activeCascade;
    renderFiltersForCascade(activeCascade);
    if(saved?.selections?.[activeCascade]) restoreSelections(activeCascade, saved.selections[activeCascade]);
    if(saved?.search){ const s=document.getElementById("searchInput"); if (s) s.value = saved.search; }
    currentRows = gatherFilteredRows(activeCascade);
    applySearch();
  }

  /* ========= Time zone / deadlines ========= */
  const timeBackdrop=document.getElementById("timeBackdrop"),
        timeModal   =document.getElementById("timeModal"),
        timeAttrRows=document.getElementById("timeAttrRows"),
        btnTimeSave =document.getElementById("btnTimeSave");

  const TIMEZONES = ["Europe/Rome","Asia/Dubai","Europe/Paris","Europe/Berlin","UTC"];
  // Настройки проверки длительности (в днях, включительно)
  const TIME_VALIDATION = { minDays: 1, maxDays: 3650 }; // минимум 1 день, максимум 10 лет

  function defaultTimeCfg(){ return { perAttr: {} }; }
  function loadTimeCfg(){ try{ const raw=localStorage.getItem(CONFIG.TIME_KEY); return raw?JSON.parse(raw):defaultTimeCfg(); }catch{ return defaultTimeCfg(); } }
  function saveTimeCfg(cfg){ localStorage.setItem(CONFIG.TIME_KEY, JSON.stringify(cfg)); }

  // миграция старых структур: переносим attrDates -> perAttr, игнорируем любые defaultTimezone/timezone
  function migrateOldCfgIfNeeded(cfg, attrKeys){
    if (!cfg || (cfg.perAttr && typeof cfg.perAttr === "object")) return cfg || defaultTimeCfg();
    const old = cfg || {};
    const out = { perAttr: {} };
    const dates = old.attrDates || {};
    for (const k of attrKeys){
      const rec = dates[k] || {};
      out.perAttr[k] = { tz: CONFIG.CSV.timezoneDefault, from: rec.from || "", until: rec.until || "" };
    }
    return out;
  }

  // Берём TZ только из perAttr[key].tz, иначе — системный дефолт CONFIG.CSV.timezoneDefault
  function getAttrTimezone(attrRaw){
    const key = String(attrRaw||"").trim().toLowerCase();
    const cfg = loadTimeCfg();
    if (cfg?.perAttr?.[key]?.tz) return cfg.perAttr[key].tz;
    return CONFIG.CSV.timezoneDefault;
  }

  // Форматируем 'YYYY-MM-DD' -> 'DD/MM/YYYY 00:00'
  function toCSVDate(s){
    if (!s) return "";
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s).trim());
    if (!m) return String(s); // если вдруг пришёл иной формат — не падаем
    const [, yyyy, mm, dd] = m;
    return `${dd}/${mm}/${yyyy} 00:00`;
  }

  // Вспомогательная: разница в днях (включительно)
  function diffDaysInclusive(from, until){
    const a = Date.parse(from), b = Date.parse(until);
    if (isNaN(a) || isNaN(b)) return null;
    const ms = b - a;
    return Math.floor(ms/86400000) + 1;
  }

  function openTimeModal(){
    if(!timeAttrRows||!timeBackdrop||!timeModal) return;
    const attrs = Array.from(new Set(displayedRows.map(r => (String(r["built-in enrollment attribute"]||"").trim())).filter(Boolean))).sort((a,b)=>a.localeCompare(b));
    let cfg = loadTimeCfg();
    const keys = attrs.map(a => a.toLowerCase());
    cfg = migrateOldCfgIfNeeded(cfg, keys);
    timeAttrRows.innerHTML = "";

    if (!attrs.length){
      const div = document.createElement("div"); div.className = "muted pad-v-8"; div.textContent = "No attributes found in the current result."; timeAttrRows.appendChild(div);
    } else {
      attrs.forEach(attr => {
        const key = attr.toLowerCase();
        const rec = (cfg.perAttr && cfg.perAttr[key]) || { tz: CONFIG.CSV.timezoneDefault, from:"", until:"" };
        const row = document.createElement("div"); row.className = "attr-row"; row.dataset.attr = attr;

        const tzWrap = document.createElement("div");
        const tzSel = document.createElement("select"); tzSel.className = "t-tz"; tzSel.setAttribute("data-attr", attr);
        TIMEZONES.forEach(z=>{ const opt=document.createElement("option"); opt.value=z; opt.textContent=z; if(z===rec.tz) opt.selected=true; tzSel.appendChild(opt); });
        tzWrap.appendChild(tzSel);

        const attrSpan = document.createElement("div"); attrSpan.textContent = attr;

        const fromWrap = document.createElement("div");
        const fromInp = document.createElement("input"); fromInp.type="date"; fromInp.className="t-from"; fromInp.setAttribute("data-attr",attr); fromInp.value=rec.from || "";
        fromWrap.appendChild(fromInp);

        const untilWrap= document.createElement("div");
        const untilInp = document.createElement("input"); untilInp.type="date"; untilInp.className="t-until"; untilInp.setAttribute("data-attr",attr); untilInp.value=rec.until || "";
        untilWrap.appendChild(untilInp);

        row.appendChild(tzWrap); row.appendChild(attrSpan); row.appendChild(fromWrap); row.appendChild(untilWrap);
        timeAttrRows.appendChild(row);
      });
    }

    // Добавляем кнопку "Clear all dates" в панель действий модалки (рядом с Save/Cancel)
    if (btnTimeSave) {
      const actions = btnTimeSave.parentElement;
      if (actions && !document.getElementById("btnTimeClear")) {
        const clearBtn = document.createElement("button");
        clearBtn.id = "btnTimeClear";
        clearBtn.type = "button";
        clearBtn.textContent = "Clear all dates";
        actions.insertBefore(clearBtn, btnTimeSave); // перед Save
        clearBtn.addEventListener("click", clearAllDates);
      }
    }

    timeBackdrop.style.display="block"; timeModal.style.display="block";
    validateAttrDates();
  }
  function closeTimeModal(){ if(!timeModal||!timeBackdrop) return; timeModal.style.display="none"; timeBackdrop.style.display="none"; updateResultCounters(); }

  // Очистка всех дат в модалке
  function clearAllDates(){
    if (!timeAttrRows) return;
    timeAttrRows.querySelectorAll('.attr-row[data-attr]').forEach(row=>{
      const from = row.querySelector(".t-from");
      const until= row.querySelector(".t-until");
      if (from) from.value = "";
      if (until) until.value = "";
    });
    validateAttrDates();
  }

  // Валидация дат: until >= from, а также длительность в пределах TIME_VALIDATION
  function validateAttrDates(){
    if (!timeAttrRows || !btnTimeSave) return true;
    let ok = true;
    timeAttrRows.querySelectorAll(".err-note").forEach(n=>n.remove());
    timeAttrRows.querySelectorAll(".invalid").forEach(n=>n.classList.remove("invalid"));

    const rows = Array.from(timeAttrRows.querySelectorAll(".attr-row"));
    rows.forEach(row=>{
      const from = row.querySelector(".t-from")?.value || "";
      const until= row.querySelector(".t-until")?.value || "";
      if (!from && !until) return; // обе пустые — ок
      if (from && until) {
        if (until < from){
          ok=false; row.classList.add("invalid");
          const m=document.createElement("div"); m.className="err-note";
          m.textContent="Active until must be the same as or later than Active from.";
          row.appendChild(m);
          return;
        }
        const diff = diffDaysInclusive(from, until);
        if (diff !== null) {
          if (diff < TIME_VALIDATION.minDays){
            ok=false; row.classList.add("invalid");
            const m=document.createElement("div"); m.className="err-note";
            m.textContent = `Active period must be at least ${TIME_VALIDATION.minDays} day(s).`;
            row.appendChild(m);
          } else if (diff > TIME_VALIDATION.maxDays){
            ok=false; row.classList.add("invalid");
            const m=document.createElement("div"); m.className="err-note";
            m.textContent = `Active period must not exceed ${TIME_VALIDATION.maxDays} day(s).`;
            row.appendChild(m);
          }
        }
      }
      // если указан только один конец — валидно. Скажешь — сделаю требование обоих дат.
    });

    btnTimeSave.disabled = !ok; return ok;
  }

  on("btnTime","click", openTimeModal);
  on("btnTimeCancel","click", closeTimeModal);
  timeBackdrop?.addEventListener("click", closeTimeModal);
  timeAttrRows?.addEventListener("input", validateAttrDates);
  on("btnTimeSave","click", ()=>{
    if(!validateAttrDates()) return alert("Fix date errors first.");
    const cfg = loadTimeCfg();
    const out = { perAttr: (cfg && cfg.perAttr) ? cfg.perAttr : {} };
    // Только строки с атрибутами
    const rows = Array.from(timeAttrRows?.querySelectorAll('.attr-row[data-attr]') || []);
    rows.forEach(row=>{
      const attr = row.dataset.attr || "";
      const key = attr.toLowerCase();
      const tz = row.querySelector(".t-tz")?.value || CONFIG.CSV.timezoneDefault;
      const from = row.querySelector(".t-from")?.value || "";
      const until = row.querySelector(".t-until")?.value || "";
      out.perAttr[key] = { tz, from, until };
    });
    saveTimeCfg(out); closeTimeModal();
  });

  /* ========= Learning plan CSV ========= */
  on("btnCSV","click", ()=>{
    if(!DB) return;
    const activeUsers = loadParticipantsLS().filter(p => p && p.active);
    if(!activeUsers.length) return alert("No participants selected.");
    if(!displayedRows.length) return alert("No data rows to export.");
    const cfg = loadTimeCfg();
    const header = CONFIG.CSV.header.join(",");
    const lines = [header];
    for (const p of activeUsers) {
      for (const r of displayedRows) {
        const lpUUID = r["Learning Plan UUID"];
        const attrRaw = norm(r["built-in enrollment attribute"]);
        const tz = getAttrTimezone(attrRaw);
        const rec = (cfg.perAttr||{})[attrRaw.toLowerCase()] || {};
        const from = toCSVDate(rec.from || "");
        const until = toCSVDate(rec.until || "");
        lines.push([
          `"${String(p.username ?? p.user ?? "").replace(/"/g,'""')}"`,
          `"${String(p.user_id ?? "").replace(/"/g,'""')}"`,
          `""`,
          `"${String(lpUUID ?? "").replace(/"/g,'""')}"`,
          `"${tz}"`,
          `"${from}"`,
          `"${until}"`,
          `"${attrRaw.replace(/"/g,'""')}"`
        ].join(","));
      }
    }
    const blob=new Blob([lines.join("\n")],{type:"text/csv"});
    const a=document.createElement("a");
    a.href=URL.createObjectURL(blob);
    a.download = `${CONFIG.CSV.filenamePrefix}${new Date().toISOString().slice(0,10)}.csv`;
    a.click();
    setTimeout(()=>URL.revokeObjectURL(a.href), 1000);
  });

  /* ========= Skill list (снимок + экспорт с главной) ========= */
  on("btnSkillList","click", ()=>{
    if(!DB) return;
    const scope = gatherFilteredRows(activeCascade);
    const q = (document.getElementById("searchInput")?.value || "").trim().toLowerCase();
    const colsForSearch = (CONFIG.CASCADES[activeCascade]?.output) || [];
    const rows = q ? scope.filter(r => colsForSearch.some(c => String(r[c]||"").toLowerCase().includes(q))) : scope;
    const cols = CONFIG.PREFERRED_SKILL_COLUMNS.slice();
    const reduced = rows.map(r=>{ const o={}; cols.forEach(c => { o[c] = r[c] ?? ""; }); return o; });
    const snapshot = { ts: Date.now(), cascade: activeCascade, columns: cols, rows: reduced };
    try { localStorage.setItem(CONFIG.SKILL_SNAPSHOT_KEY, JSON.stringify(snapshot)); } catch(e){ console.warn("Skill snapshot save failed", e); }
    window.open("data/skilllist.html","_blank","noopener,noreferrer");
  });
  on("btnSkillCSV","click", ()=>{
    if(!DB) return;
    const scope = gatherFilteredRows(activeCascade);
    const q = (document.getElementById("searchInput")?.value || "").trim().toLowerCase();
    const colsForSearch = (CONFIG.CASCADES[activeCascade]?.output) || [];
    const rows = q ? scope.filter(r => colsForSearch.some(c => String(r[c]||"").toLowerCase().includes(q))) : scope;
    if(!rows.length){ alert("No rows to export."); return; }
    const cols = CONFIG.PREFERRED_SKILL_COLUMNS.slice();
    const table = rows.map(r=>{ const o = {}; cols.forEach(c => { o[c] = r[c] ?? ""; }); return o; });
    if (typeof XLSX === "undefined") { alert("Excel export library (XLSX) is not loaded on the main page. Use Skill list page or include XLSX."); return; }
    const ws = XLSX.utils.json_to_sheet(table, { header: cols });
    XLSX.utils.sheet_add_aoa(ws, [cols], { origin: "A1" });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Skill list");
    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `skill_list_${new Date().toISOString().slice(0,10)}.xlsx`;
    a.click(); setTimeout(()=>URL.revokeObjectURL(a.href), 1000);
  });

  /* ========= Tabs ========= */
  function switchTab(name){
    const admin = name === "Admin";
    const adminEl = document.getElementById("adminPanel");
    const userEl  = document.getElementById("userPanel");
    if (adminEl) adminEl.classList.toggle("hidden", !admin);
    if (userEl)  userEl.classList.toggle("hidden", admin);
    if (!admin) initUserPanel();
  }
  on("goAdmin","click", () => switchTab("Admin"));
  on("goUser","click", () => switchTab("User"));
  
  // Connect the renamed nav items to original functionality
  // The IDs remain the same but text labels changed in HTML
  // goAdmin and goUser are already connected above
  
  switchTab("Admin");
});