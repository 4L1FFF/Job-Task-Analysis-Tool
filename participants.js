(function(){
  const PKEY = "jtat_participants";
  const $ = s => document.querySelector(s);
  const $$ = s => Array.from(document.querySelectorAll(s));
  const norm = v => String(v??"").trim();
  const lower= v => norm(v).toLowerCase();

  let LIST = [];

  /* ---------- storage ---------- */
  function load(){ try{ LIST = JSON.parse(localStorage.getItem(PKEY) || "[]"); }catch{ LIST = []; } ensureIds(); save(); }
  function save(){ localStorage.setItem(PKEY, JSON.stringify(LIST)); }
  function ensureIds(){
    let changed=false;
    LIST.forEach(p=>{ if(!p._id){ p._id = crypto?.randomUUID?.() || (Date.now()+Math.random().toString(16).slice(2)); changed=true; } });
    if(changed) save();
  }

  /* ---------- counters & header checkbox ---------- */
  function counters(){
    const tot=LIST.length, sel=LIST.filter(p=>p.active).length;
    $("#pl_total") && ($("#pl_total").textContent = "Total: "+tot);
    $("#pl_selected") && ($("#pl_selected").textContent = "Selected: "+sel);
    const dups = LIST.reduce((n,p)=> n + ((p.__dup_username||p.__dup_user_id)?1:0), 0);
    $("#pl_dups") && ($("#pl_dups").textContent = "Duplicates: "+dups);

    const useAll = $("#useAll");
    if (useAll) {
      if (!LIST.length) {
        useAll.checked = false;
        useAll.indeterminate = false;
        useAll.disabled = true;
      } else {
        useAll.disabled = false;
        if (sel === 0) { useAll.checked = false; useAll.indeterminate = false; }
        else if (sel === tot) { useAll.checked = true; useAll.indeterminate = false; }
        else { useAll.checked = false; useAll.indeterminate = true; }
      }
    }
  }

  /* ---------- duplicates mark ---------- */
  function markDups(){
    const cU=new Map(), cI=new Map();
    LIST.forEach(p=>{ const u=lower(p.username); if(u) cU.set(u,(cU.get(u)||0)+1); const id=norm(p.user_id); if(id) cI.set(id,(cI.get(id)||0)+1); });
    LIST.forEach(p=>{ const u=lower(p.username), id=norm(p.user_id); p.__dup_username=!!u&&cU.get(u)>1; p.__dup_user_id=!!id&&cI.get(id)>1; });
  }

  /* ---------- render ---------- */
  function render(){
    markDups();

    const q=lower($("#q")?.value||"");
    const only=$("#onlySel")?.checked;

    let rows=LIST.filter(p=>{
      if(only && !p.active) return false;
      if(!q) return true;
      return [p.user_id,p.username,p.first_name,p.last_name].some(v=>lower(v).includes(q));
    }).sort((a,b)=> lower(a.username).localeCompare(lower(b.username)));

    const tb=$("#tbl tbody"); if(!tb) return; tb.innerHTML="";
    if(!rows.length){
      tb.innerHTML = `<tr><td class="muted" colspan="6">No data. Import CSV or add a participant.</td></tr>`;
    } else {
      rows.forEach(p=>{
        const tr=document.createElement("tr");
        if (p.__dup_username) tr.classList.add("dup-username");
        if (p.__dup_user_id) tr.classList.add("dup-user_id");
        tr.innerHTML = `
          <td style="text-align:center;width:56px">
            <input type="checkbox" class="use" data-id="${p._id}" ${p.active?'checked':''}>
          </td>
          <td><input class="ed uid" data-id="${p._id}" value="${norm(p.user_id)}"></td>
          <td><input class="ed uname" data-id="${p._id}" value="${norm(p.username)}"></td>
          <td><input class="ed fn" data-id="${p._id}" value="${norm(p.first_name)}"></td>
          <td><input class="ed ln" data-id="${p._id}" value="${norm(p.last_name)}"></td>
          <td><button class="rm" data-id="${p._id}">Delete</button></td>`;
        tb.appendChild(tr);
      });
    }
    counters();
  }

  /* ---------- events: table ---------- */
  $("#tbl")?.addEventListener("change", (e)=>{
    const t=e.target;

    // чекбокс строки
    if (t.classList.contains("use")) {
      const id=t.getAttribute("data-id");
      const i=LIST.findIndex(p=>p._id===id); if(i<0) return;
      LIST[i].active = t.checked;
      save(); counters();
      return;
    }

    // инпуты строки
    const id=t.getAttribute("data-id"); if(!id) return;
    const i=LIST.findIndex(p=>p._id===id); if(i<0) return;
    if (t.classList.contains("ed")){
      const v=t.value;
      if (t.classList.contains("uid")) LIST[i].user_id   = v;
      if (t.classList.contains("uname")) LIST[i].username  = v;
      if (t.classList.contains("fn")) LIST[i].first_name = v;
      if (t.classList.contains("ln")) LIST[i].last_name  = v;
      save(); render();
    }
  });

  $("#tbl")?.addEventListener("click",(e)=>{
    const t=e.target; if (!t.classList.contains("rm")) return;
    const id=t.getAttribute("data-id"); const i=LIST.findIndex(p=>p._id===id);
    if(i>=0 && confirm('Remove user "'+(LIST[i].username||'')+'"?')){ LIST.splice(i,1); save(); render(); }
  });

  /* ---------- header select-all ---------- */
  $("#useAll")?.addEventListener("change", (e)=>{
    const selectAll = e.target.indeterminate ? true : e.target.checked;
    if (e.target.indeterminate) { e.target.indeterminate = false; e.target.checked = true; }
    LIST.forEach(p => { p.active = !!selectAll; });
    save(); render();
  });

  /* ---------- toolbar: search / filter ---------- */
  $("#q")?.addEventListener("input", ()=>render());
  $("#onlySel")?.addEventListener("change", ()=>render());

  /* ---------- toolbar: add / remove ---------- */
  document.getElementById("btnRemoveAll")?.addEventListener("click", ()=>{
    if(!LIST.length) return;
    if(confirm("Remove ALL users?")){ LIST=[]; save(); render(); }
  });

  /* =========================
       Add participant modal
     ========================= */
  const ap = {
    backdrop: $("#addBackdrop"),
    modal:    $("#addModal"),
    active:   $("#ap_active"),
    user_id:  $("#ap_user_id"),
    username: $("#ap_username"),
    first:    $("#ap_first_name"),
    last:     $("#ap_last_name"),
    errUser:  $("#ap_err_username"),
    btnSave:  $("#ap_save"),
    btnCancel:$("#ap_cancel"),
  };

  function openAddModal(){
    // reset
    ap.active.checked = true;
    ap.user_id.value = "";
    ap.username.value = "";
    ap.first.value = "";
    ap.last.value = "";
    ap.errUser.style.display = "none";
    ap.username.classList.remove("invalid");
    // show
    ap.backdrop.style.display = "block";
    ap.modal.style.display = "block";
    // focus
    setTimeout(()=> ap.username?.focus(), 0);
  }
  function closeAddModal(){
    ap.modal.style.display = "none";
    ap.backdrop.style.display = "none";
  }
  function validateAdd(){
    const ok = norm(ap.username.value).length > 0;
    ap.errUser.style.display = ok ? "none" : "block";
    ap.username.classList.toggle("invalid", !ok);
    return ok;
  }
  function saveAdd(){
    if(!validateAdd()) return;
    const rec = {
      _id: crypto?.randomUUID?.() || (Date.now()+Math.random().toString(16).slice(2)),
      active: !!ap.active.checked,
      user_id: norm(ap.user_id.value),
      username: norm(ap.username.value),
      first_name: norm(ap.first.value),
      last_name: norm(ap.last.value),
    };
    LIST.unshift(rec);
    save();
    render();
    closeAddModal();
  }

  // События модалки
  ap.btnSave?.addEventListener("click", saveAdd);
  ap.btnCancel?.addEventListener("click", closeAddModal);
  ap.modal?.addEventListener("keydown", (e)=>{
    if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) { e.preventDefault(); saveAdd(); }
    if (e.key === "Escape") { e.preventDefault(); closeAddModal(); }
  });
  ap.username?.addEventListener("input", validateAdd);
  ap.backdrop?.addEventListener("click", closeAddModal);

  // Кнопка «Add participant» открывает модалку
  document.getElementById("btnAdd")?.addEventListener("click", openAddModal);

  /* ---------- CSV import ---------- */
  document.getElementById("btnImportCSV")?.addEventListener("click", ()=>{ document.getElementById("fileImportCSV")?.click(); });
  document.getElementById("fileImportCSV")?.addEventListener("change", async (e)=>{
    const f=e.target.files?.[0]; if(!f) return;
    const text = await f.text();
    const rows = parseCSV(text).filter(r=>r.username||r.user_id);
    if(!rows.length){ alert("No valid rows found in CSV."); return; }
    mergeParticipants(rows); e.target.value="";
  });

  function parseCSV(text){
    const lines = text.replace(/\r\n/g,"\n").replace(/\r/g,"\n").split("\n").filter(l=>l.trim().length);
    if(!lines.length) return [];
    const header = lines[0].split(",").map(h=>h.trim().replace(/^"|"$/g,""));
    const idx = {
      user_id: header.findIndex(h=>/user[_ ]?id/i.test(h)),
      username: header.findIndex(h=>/user ?name/i.test(h)),
      first_name: header.findIndex(h=>/first/i.test(h)),
      last_name: header.findIndex(h=>/last/i.test(h)),
      active: header.findIndex(h=>/active/i.test(h))
    };
    return lines.slice(1).map(line=>{
      const cells = line.split(",");
      const val = i => (i>=0 && i<cells.length) ? cells[i].replace(/^"|"$/g,"") : "";
      const a = val(idx.active);
      return {
        _id: crypto?.randomUUID?.() || (Date.now()+Math.random().toString(16).slice(2)),
        active: /^(1|true|yes)$/i.test(a),
        user_id: val(idx.user_id),
        username: val(idx.username),
        first_name: val(idx.first_name),
        last_name: val(idx.last_name)
      };
    });
  }
  function mergeParticipants(imported){
    const byU = new Map(LIST.map(p=>[lower(p.username), p]));
    const byI = new Map(LIST.map(p=>[norm(p.user_id), p]));
    imported.forEach(p=>{
      const u = lower(p.username), id = norm(p.user_id);
      let target = (u && byU.get(u)) || (id && byI.get(id));
      if (target){
        target.first_name = p.first_name || target.first_name;
        target.last_name  = p.last_name  || target.last_name;
        target.user_id    = p.user_id    || target.user_id;
        target.active     = (typeof p.active==='boolean') ? p.active : target.active;
      } else {
        LIST.push(p); if (u) byU.set(u, p); if (id) byI.set(id, p);
      }
    });
    save(); render();
  }

  /* ---------- Save & Close ---------- */
  document.getElementById("btnSaveClose")?.addEventListener("click", ()=>{
    window.close();
    setTimeout(()=>{ try{ if(!window.closed) window.location.href="about:blank"; }catch{} }, 100);
  });

  /* ---------- init ---------- */
  load(); render();
})();