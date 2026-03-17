(function () {
  // Ключ снапшота (должен совпадать с app.js)
  const KEY  = "jtat_skill_snapshot";
  // Фиксированный порядок колонок для вывода
  const COLS = ["Job Role","Industry/Segment","Level","Skills","SMCS Code","SMCS Job","Task"];

  const $   = (s) => document.querySelector(s);
  const esc = (s) => String(s).replace(/[&<>"]/g, c => ({'&':'&','<':'<','>':'>'}[c] || c));

  function loadSnapshot() {
    try { return JSON.parse(localStorage.getItem(KEY) || "null"); }
    catch { return null; }
  }

  function render() {
    const ss   = loadSnapshot();
    const thead= $("#tbl thead");
    const tbody= $("#tbl tbody");
    const meta = $("#meta");
    const note = $("#note");

    // Нет снапшота — показываем подсказку
    if (!ss || !Array.isArray(ss.rows)) {
      if (thead) thead.innerHTML = "";
      if (tbody) tbody.innerHTML = `
        <tr><td class="muted" style="padding:12px">No snapshot found. Go back to the main page and click Skill list again.</td></tr>
      `;
      if (meta) meta.textContent = "—";
      if (note) note.textContent = "";
      return;
    }

    // Метаданные
    if (meta) meta.textContent = `${ss.cascade || "—"} · ${ss.rows.length} rows`;
    if (note) note.textContent = `Snapshot time: ${new Date(ss.ts).toLocaleString()}`;

    // Заголовок таблицы
    if (thead) {
      thead.innerHTML =
        "<tr>" + COLS.map(c => `<th>${esc(c)}</th>`).join("") + "</tr>";
    }

    // Тело таблицы
    if (!ss.rows.length) {
      if (tbody) {
        tbody.innerHTML = `
          <tr><td class="muted" style="padding:12px" colspan="${COLS.length}">
            No data for current filters/search. Return to the main page to adjust filters and re-open Skill list.
          </td></tr>
        `;
      }
    } else {
      if (tbody) {
        tbody.innerHTML = ss.rows.map(r =>
          ("<tr>" + COLS.map(c => `<td>${esc(r[c] ?? "")}</td>`).join("") + "</tr>")
        ).join("");
      }
    }
  }

  // Экспорт в Excel (использует XLSX, если библиотека подключена)
  function exportExcel() {
    const ss = loadSnapshot();
    if (!ss || !Array.isArray(ss.rows)) {
      alert("No snapshot to export. Please refresh or re-open from the main app.");
      return;
    }
    if (typeof XLSX === "undefined") {
      alert("Excel export library (XLSX) is not loaded.");
      return;
    }

    const aoa = [ COLS ];
    ss.rows.forEach(r => aoa.push(COLS.map(c => String(r[c] ?? ""))));

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    // минимальная ширина колонок — для читаемости
    ws['!cols'] = COLS.map(c => ({ wch: Math.max(12, c.length + 2) }));

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Skill list");

    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob  = new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

    const d = new Date().toISOString().slice(0,10);
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `skill_list_${ss.cascade || 'cascade'}_${d}.xlsx`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  }

  // События
  $("#btnRefresh")?.addEventListener("click", render);
  $("#btnDownload")?.addEventListener("click", exportExcel);
  window.addEventListener("storage", (e) => { if (e.key === KEY) render(); });

  // Первый рендер
  render();
})();