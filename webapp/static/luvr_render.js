/** Общий рендер ЛУВР: сетка, правки → yaml ↔ xlsx. */
window.LuvrPanel = (function () {
  function esc(s) {
    return String(s ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/"/g, "&quot;");
  }

  function markClass(mark) {
    if (mark === "1") return "luvr-mark luvr-mark--full";
    if (mark === "0.5") return "luvr-mark luvr-mark--half";
    if (mark) return "luvr-mark luvr-mark--other";
    return "luvr-mark luvr-mark--empty";
  }

  function markLabel(mark) {
    return mark || "·";
  }

  function cycleMark(mark) {
    if (!mark) return "1";
    if (mark === "1") return "0.5";
    if (mark === "0.5") return "";
    return "";
  }

  function renderSummaryTable(month) {
    const people = month.people || [];
    const rows = people
      .map(
        (p) =>
          `<tr>
            <td>${esc(p.num)}</td>
            <td>${esc(p.fio)}</td>
            <td>${esc(p.position || "—")}</td>
            <td>${p.days_present ?? 0}</td>
            <td>${p.days_marked ?? 0}</td>
          </tr>`
      )
      .join("");
    return `<table class="planning-table personnel-table">
      <thead><tr><th>№</th><th>ФИО</th><th>Должность</th><th>Дней</th><th>Отметок</th></tr></thead>
      <tbody>${rows || '<tr><td colspan="5" class="hint-text">Нет данных</td></tr>'}</tbody>
    </table>`;
  }

  function renderPersonLinkSelect(pi, person, personnel) {
    const pid = person.person_id || "";
    const src = person.link_source || (pid ? "auto" : "unmatched");
    const cls =
      src === "unmatched" ? "luvr-link--warn" : src === "manual" ? "luvr-link--manual" : "luvr-link--ok";
    const opts = (personnel || [])
      .map(
        (p) =>
          `<option value="${esc(p.id)}"${p.id === pid ? " selected" : ""}>${esc(p.fio)}</option>`
      )
      .join("");
    return `<select class="luvr-person-link ${cls}" data-person-idx="${pi}" title="person_id из personnel.yaml">
      <option value="">— не связан —</option>${opts}</select>`;
  }

  function monthLinkStats(month) {
    let linked = 0;
    const people = month.people || [];
    for (const p of people) {
      if (p.person_id) linked += 1;
    }
    return { linked, total: people.length, unmatched: people.length - linked };
  }

  function renderGrid(month, editable, personnel) {
    const days = month.days || [];
    const people = month.people || [];
    if (!days.length || !people.length || !("marks" in people[0])) {
      return (
        '<p class="warn-text">Нет сетки по дням. Запустите: <code>python scripts/build_engineer_data.py --luvr</code></p>' +
        renderSummaryTable(month)
      );
    }

    const dayHeaders = days.map((d) => `<th class="luvr-day" title="${esc(d.date)}">${d.day}</th>`).join("");
    const body = people
      .map((p, pi) => {
        const marks = p.marks || [];
        const cells = days
          .map((_, i) => {
            const m = marks[i] ?? "";
            const cls = markClass(m);
            const editAttrs = editable
              ? ` class="${cls} luvr-cell" data-person-idx="${pi}" data-day-idx="${i}" tabindex="0" role="gridcell"`
              : ` class="${cls}"`;
            return `<td${editAttrs}>${esc(markLabel(m))}</td>`;
          })
          .join("");
        return `<tr data-person-idx="${pi}" class="${p.person_id ? "" : "luvr-row--unlinked"}">
          <td class="luvr-sticky luvr-col-num">${esc(p.num)}</td>
          <td class="luvr-sticky luvr-col-fio">${esc(p.fio)}</td>
          <td class="luvr-sticky luvr-col-link">${renderPersonLinkSelect(pi, p, personnel)}</td>
          <td class="luvr-sticky luvr-col-pos" title="${esc(p.position)}">${esc(p.position || "—")}</td>
          ${cells}
          <td class="luvr-sum" data-sum-for="${pi}">${p.days_present ?? 0}</td>
        </tr>`;
      })
      .join("");

    const editHint = editable
      ? '<p class="hint-text luvr-legend">Клик — цикл · → <span class="luvr-mark luvr-mark--full">1</span> → <span class="luvr-mark luvr-mark--half">0.5</span> · клавиши <kbd>1</kbd> / <kbd>.</kbd> / <kbd>Del</kbd></p>'
      : '<p class="hint-text luvr-legend">· — нет отметки · <span class="luvr-mark luvr-mark--full">1</span> полный день · <span class="luvr-mark luvr-mark--half">0.5</span> полдня</p>';

    return `<div class="luvr-grid-wrap">
      <table class="planning-table luvr-grid">
        <thead>
          <tr>
            <th class="luvr-sticky luvr-col-num">№</th>
            <th class="luvr-sticky luvr-col-fio">ФИО</th>
            <th class="luvr-sticky luvr-col-link">Справ.</th>
            <th class="luvr-sticky luvr-col-pos">Должность</th>
            ${dayHeaders}
            <th class="luvr-sum">Σ</th>
          </tr>
        </thead>
        <tbody>${body}</tbody>
      </table>
    </div>
    ${editHint}
    <p id="luvrSaveStatus" class="luvr-save-status" aria-live="polite"></p>`;
  }

  function renderMonthPanel(month, editable, personnel) {
    return `<div class="luvr-month-panel">${renderGrid(month, editable, personnel)}</div>`;
  }

  async function savePersonLink(sheet, personIdx, personId) {
    const res = await fetch("/api/luvr/link", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        sheet,
        person_idx: personIdx,
        person_id: personId,
      }),
    });
    if (!res.ok) {
      let detail = res.statusText;
      try {
        const err = await res.json();
        detail = err.detail || detail;
      } catch (_) {
        /* ignore */
      }
      throw new Error(detail);
    }
    return res.json();
  }

  async function apiAutoLink() {
    const res = await fetch("/api/luvr/auto-link", { method: "POST" });
    if (!res.ok) {
      let detail = res.statusText;
      try {
        const err = await res.json();
        detail = err.detail || detail;
      } catch (_) {
        /* ignore */
      }
      throw new Error(detail);
    }
    return res.json();
  }

  function bindPersonLinks(container, month, sheet, personnel, onLinked) {
    container.querySelectorAll("select.luvr-person-link").forEach((sel) => {
      sel.addEventListener("change", async () => {
        const personIdx = Number(sel.dataset.personIdx);
        const person = (month.people || [])[personIdx];
        if (!person) return;
        const prev = person.person_id || "";
        const next = sel.value;
        if (prev === next) return;
        sel.disabled = true;
        try {
          const result = await savePersonLink(sheet, personIdx, next);
          person.person_id = result.person_id || null;
          person.link_source = result.link_source;
          sel.className = `luvr-person-link ${
            result.link_source === "unmatched"
              ? "luvr-link--warn"
              : result.link_source === "manual"
                ? "luvr-link--manual"
                : "luvr-link--ok"
          }`;
          const row = sel.closest("tr");
          if (row) row.classList.toggle("luvr-row--unlinked", !person.person_id);
          if (onLinked) onLinked();
        } catch (e) {
          sel.value = prev;
          alert(e.message);
        } finally {
          sel.disabled = false;
        }
      });
    });
  }

  async function saveMark(sheet, personIdx, dayIdx, mark) {
    const res = await fetch("/api/luvr/mark", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        sheet,
        person_idx: personIdx,
        day_idx: dayIdx,
        mark,
      }),
    });
    if (!res.ok) {
      let detail = res.statusText;
      try {
        const err = await res.json();
        detail = err.detail || detail;
      } catch (_) {
        /* ignore */
      }
      throw new Error(detail);
    }
    return res.json();
  }

  function applyCellUI(td, mark) {
    td.textContent = markLabel(mark);
    td.className = `${markClass(mark)} luvr-cell`;
  }

  function bindGrid(container, month, sheet, statusEl, onXlsxStale) {
    const people = month.people || [];
    let pending = 0;

    function setStatus(text, kind) {
      if (!statusEl) return;
      statusEl.textContent = text;
      statusEl.className = `luvr-save-status${kind ? ` luvr-save-status--${kind}` : ""}`;
    }

    async function commitCell(td, personIdx, dayIdx, mark) {
      const person = people[personIdx];
      if (!person) return;

      const prev = (person.marks || [])[dayIdx] ?? "";
      if (prev === mark) return;

      person.marks = person.marks || [];
      person.marks[dayIdx] = mark;
      applyCellUI(td, mark);
      td.classList.add("luvr-cell--pending");

      pending += 1;
      setStatus("Сохранение…", "pending");

      try {
        const result = await saveMark(sheet, personIdx, dayIdx, mark);
        person.marks[dayIdx] = result.mark;
        person.days_present = result.days_present;
        person.days_marked = result.days_marked;
        applyCellUI(td, result.mark);
        const sumCell = container.querySelector(`td.luvr-sum[data-sum-for="${personIdx}"]`);
        if (sumCell) sumCell.textContent = String(result.days_present);
        td.classList.remove("luvr-cell--pending");
        td.classList.add("luvr-cell--saved");
        setTimeout(() => td.classList.remove("luvr-cell--saved"), 600);
        if (onXlsxStale) onXlsxStale(Boolean(result.xlsx_stale));
        if (result.xlsx_synced) setStatus("Сохранено в yaml и Excel", "ok");
        else setStatus("Сохранено в yaml (Excel не на диске)", "ok");
      } catch (e) {
        person.marks[dayIdx] = prev;
        applyCellUI(td, prev);
        td.classList.remove("luvr-cell--pending");
        setStatus(`Ошибка: ${e.message}`, "error");
      } finally {
        pending -= 1;
      }
    }

    container.querySelectorAll("td.luvr-cell").forEach((td) => {
      const personIdx = Number(td.dataset.personIdx);
      const dayIdx = Number(td.dataset.dayIdx);

      td.addEventListener("click", () => {
        const cur = (people[personIdx]?.marks || [])[dayIdx] ?? "";
        const next = cycleMark(cur);
        commitCell(td, personIdx, dayIdx, next);
      });

      td.addEventListener("keydown", (ev) => {
        const cur = (people[personIdx]?.marks || [])[dayIdx] ?? "";
        let next = null;
        if (ev.key === "1") next = "1";
        else if (ev.key === "." || ev.key === ",") next = "0.5";
        else if (ev.key === "Delete" || ev.key === "Backspace" || ev.key === "0") next = "";
        else if (ev.key === " " || ev.key === "Enter") {
          ev.preventDefault();
          next = cycleMark(cur);
        }
        if (next === null) return;
        ev.preventDefault();
        commitCell(td, personIdx, dayIdx, next);
      });
    });
  }

  async function apiImportFromXlsx() {
    const res = await fetch("/api/luvr/import-from-xlsx", { method: "POST" });
    if (!res.ok) {
      let detail = res.statusText;
      try {
        const err = await res.json();
        detail = err.detail || detail;
      } catch (_) {
        /* ignore */
      }
      throw new Error(detail);
    }
    return res.json();
  }

  async function apiExportToXlsx(sheet) {
    const res = await fetch("/api/luvr/export-to-xlsx", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(sheet ? { sheet } : {}),
    });
    if (!res.ok) {
      let detail = res.statusText;
      try {
        const err = await res.json();
        detail = err.detail || detail;
      } catch (_) {
        /* ignore */
      }
      throw new Error(detail);
    }
    return res.json();
  }

  function renderSyncBar(xlsxPresent, xlsxStale, linkStats, monthStats) {
    const parts = [];
    if (monthStats) {
      parts.push(
        `<span class="luvr-link-summary">Справочник: ${monthStats.linked}/${monthStats.total} строк${
          monthStats.unmatched ? ` · <span class="luvr-sync-stale">${monthStats.unmatched} без связи</span>` : ""
        }</span>`
      );
    } else if (linkStats) {
      parts.push(
        `<span class="luvr-link-summary">Справочник: ${linkStats.linked}/${linkStats.total} строк</span>`
      );
    }
    if (xlsxPresent) {
      parts.push(
        xlsxStale
          ? '<span class="luvr-sync-stale">Excel не синхронизирован</span>'
          : '<span class="luvr-sync-ok">yaml ↔ Excel</span>'
      );
    }
    const actions = [`<button type="button" class="btn btn-gray btn-sm" id="luvrAutoLink">Авто-связать по ФИО</button>`];
    if (xlsxPresent) {
      actions.push(
        `<button type="button" class="btn btn-gray btn-sm" id="luvrImportXlsx">Загрузить из Excel</button>`,
        `<button type="button" class="btn btn-gray btn-sm" id="luvrExportXlsx">Выгрузить в Excel</button>`
      );
    }
    return `<div class="luvr-sync-bar">
      ${parts.join(" ")}
      <div class="luvr-sync-actions">${actions.join("")}</div>
    </div>`;
  }

  function render(el, data) {
    if (!el) return;

    if (data.cache_ready) {
      let state = { ...data, months: (data.months || []).map((m) => ({ ...m, people: (m.people || []).map((p) => ({ ...p, marks: [...(p.marks || [])] })) })) };
      const months = state.months;
      const editable = Boolean(state.editable && state.grid_ready);
      const defaultSheet = state.default_month || months[months.length - 1]?.sheet || "";
      const sourceLine = state.source
        ? `<strong>${esc(state.source)}</strong>${state.source_kb ? ` · ${state.source_kb} КБ` : ""}`
        : "luvr.yaml";
      const cacheNote =
        state.cache_from_yaml && !state.xlsx_present
          ? `<p class="hint-text">На сервере нет xlsx — правки только в <code>luvr.yaml</code>. Положите <code>ЛУВР.xlsx</code> для синхронизации с Excel.</p>`
          : "";
      const gridNote = state.grid_ready
        ? ""
        : `<p class="warn-text">Сетка по дням устарела — пересоберите кэш: <code>python scripts/build_engineer_data.py --luvr</code></p>`;

      el.innerHTML = `
        <div class="luvr-header">
          <p class="planning-meta">Источник: ${sourceLine} · <code>${esc(state.folder)}/</code></p>
          ${state.contract ? `<p class="planning-meta">Договор: ${esc(state.contract)}</p>` : ""}
          ${cacheNote}
          ${gridNote}
          <div id="luvrSyncBar">${renderSyncBar(state.xlsx_present, state.xlsx_stale, state.link_stats, null)}</div>
          <label class="field-label">Месяц</label>
          <select id="luvrMonth" class="field-input luvr-month-select"></select>
        </div>
        <div id="luvrMonthBody"></div>
        <p id="luvrGlobalStatus" class="luvr-save-status" aria-live="polite"></p>`;

      const sel = el.querySelector("#luvrMonth");
      const body = el.querySelector("#luvrMonthBody");
      const syncBar = el.querySelector("#luvrSyncBar");
      const globalStatus = el.querySelector("#luvrGlobalStatus");

      function refreshSyncBar(month) {
        const ms = month ? monthLinkStats(month) : null;
        if (syncBar) syncBar.innerHTML = renderSyncBar(state.xlsx_present, state.xlsx_stale, state.link_stats, ms);
        bindSyncButtons();
      }

      function setGlobalStatus(text, kind) {
        if (!globalStatus) return;
        globalStatus.textContent = text;
        globalStatus.className = `luvr-save-status${kind ? ` luvr-save-status--${kind}` : ""}`;
      }

      function bindSyncButtons() {
        const importBtn = el.querySelector("#luvrImportXlsx");
        const exportBtn = el.querySelector("#luvrExportXlsx");
        const autoLinkBtn = el.querySelector("#luvrAutoLink");
        if (autoLinkBtn) {
          autoLinkBtn.onclick = async () => {
            autoLinkBtn.disabled = true;
            setGlobalStatus("Авто-связка по ФИО…", "pending");
            try {
              const result = await apiAutoLink();
              const res = await fetch("/api/luvr");
              if (!res.ok) throw new Error(res.statusText);
              window.LuvrPanel.render(el, await res.json());
              setGlobalStatus(
                `Связано: ${result.linked} авто, ${result.manual} вручную, ${result.unmatched} без совпадения`,
                "ok"
              );
            } catch (e) {
              setGlobalStatus(`Авто-связка: ${e.message}`, "error");
              autoLinkBtn.disabled = false;
            }
          };
        }
        if (importBtn) {
          importBtn.onclick = async () => {
            if (!confirm("Загрузить из Excel? Текущий luvr.yaml будет перезаписан данными из xlsx.")) return;
            importBtn.disabled = true;
            setGlobalStatus("Импорт из Excel…", "pending");
            try {
              await apiImportFromXlsx();
              const res = await fetch("/api/luvr");
              if (!res.ok) throw new Error(res.statusText);
              window.LuvrPanel.render(el, await res.json());
            } catch (e) {
              setGlobalStatus(`Импорт: ${e.message}`, "error");
              importBtn.disabled = false;
            }
          };
        }
        if (exportBtn) {
          exportBtn.onclick = async () => {
            exportBtn.disabled = true;
            setGlobalStatus("Выгрузка в Excel…", "pending");
            try {
              const result = await apiExportToXlsx(null);
              state.xlsx_stale = false;
              refreshSyncBar(months.find((m) => m.sheet === sel.value));
              setGlobalStatus(
                `Excel обновлён: ${result.cells_updated} ячеек (${(result.sheets || []).join(", ")})`,
                "ok"
              );
            } catch (e) {
              setGlobalStatus(`Выгрузка: ${e.message}`, "error");
            } finally {
              exportBtn.disabled = false;
            }
          };
        }
      }

      months.forEach((m) => {
        const opt = document.createElement("option");
        opt.value = m.sheet;
        opt.textContent = `${m.sheet}${m.year ? ` ${m.year}` : ""} — ${m.people_count} чел., ${m.days_in_sheet} дн.`;
        if (m.sheet === defaultSheet) opt.selected = true;
        sel.appendChild(opt);
      });

      function showMonth(sheet) {
        const month = months.find((m) => m.sheet === sheet);
        if (!body || !month) return;
        body.innerHTML = renderMonthPanel(month, editable, state.personnel || []);
        refreshSyncBar(month);
        const panel = body.querySelector(".luvr-month-panel");
        bindPersonLinks(panel || body, month, sheet, state.personnel || [], () => refreshSyncBar(month));
        if (editable) {
          const statusEl = body.querySelector("#luvrSaveStatus");
          bindGrid(panel || body, month, sheet, statusEl, (stale) => {
            state.xlsx_stale = stale;
            refreshSyncBar(month);
          });
        }
      }

      bindSyncButtons();
      sel.addEventListener("change", () => showMonth(sel.value));
      showMonth(sel.value);
      return;
    }

    if (data.xlsx_present) {
      el.innerHTML = `<p class="planning-meta">Файл: <strong>${esc(data.source)}</strong> (${data.source_kb} КБ)</p>
        <p class="warn-text">Кэш не собран. Запустите: <code>python scripts/build_engineer_data.py --luvr</code></p>`;
      return;
    }

    el.innerHTML =
      '<p class="hint-text">Нет данных ЛУВР. Выполните <code>git pull</code> или положите <code>ЛУВР.xlsx</code> в <code>data/luvr/</code> и запустите build.</p>';
  }

  return { render };
})();
