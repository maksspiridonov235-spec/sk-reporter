/** Общий рендер ЛУВР: сетка дней × люди; этап 2 — редактирование ячеек → yaml. */
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

  function renderGrid(month, editable) {
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
        return `<tr data-person-idx="${pi}">
          <td class="luvr-sticky luvr-col-num">${esc(p.num)}</td>
          <td class="luvr-sticky luvr-col-fio">${esc(p.fio)}</td>
          <td class="luvr-sticky luvr-col-pos" title="${esc(p.position)}">${esc(p.position || "—")}</td>
          ${cells}
          <td class="luvr-sum" data-sum-for="${pi}">${p.days_present ?? 0}</td>
        </tr>`;
      })
      .join("");

    const editHint = editable
      ? '<p class="hint-text luvr-legend">Клик — цикл · · → <span class="luvr-mark luvr-mark--full">1</span> → <span class="luvr-mark luvr-mark--half">0.5</span> · клавиши <kbd>1</kbd> / <kbd>.</kbd> / <kbd>Del</kbd> · сохранение в <code>luvr.yaml</code></p>'
      : '<p class="hint-text luvr-legend">· — нет отметки · <span class="luvr-mark luvr-mark--full">1</span> полный день · <span class="luvr-mark luvr-mark--half">0.5</span> полдня</p>';

    return `<div class="luvr-grid-wrap">
      <table class="planning-table luvr-grid">
        <thead>
          <tr>
            <th class="luvr-sticky luvr-col-num">№</th>
            <th class="luvr-sticky luvr-col-fio">ФИО</th>
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

  function renderMonthPanel(month, editable) {
    return `<div class="luvr-month-panel">${renderGrid(month, editable)}</div>`;
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

  function bindGrid(container, month, sheet, statusEl) {
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
      } catch (e) {
        person.marks[dayIdx] = prev;
        applyCellUI(td, prev);
        td.classList.remove("luvr-cell--pending");
        setStatus(`Ошибка: ${e.message}`, "error");
        return;
      } finally {
        pending -= 1;
        if (pending === 0) setStatus("Сохранено в luvr.yaml", "ok");
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

  function render(el, data) {
    if (!el) return;

    if (data.cache_ready) {
      const months = data.months || [];
      const editable = Boolean(data.editable && data.grid_ready);
      const defaultSheet = data.default_month || months[months.length - 1]?.sheet || "";
      const sourceLine = data.source
        ? `<strong>${esc(data.source)}</strong>${data.source_kb ? ` · ${data.source_kb} КБ` : ""}`
        : "luvr.yaml";
      const cacheNote = data.cache_from_yaml && !data.xlsx_present
        ? `<p class="hint-text">Данные из <code>luvr.yaml</code>. Обновить из Excel: положите xlsx и <code>python scripts/build_engineer_data.py --luvr</code>.</p>`
        : "";
      const gridNote = data.grid_ready
        ? ""
        : `<p class="warn-text">Сетка по дням устарела — пересоберите кэш: <code>python scripts/build_engineer_data.py --luvr</code></p>`;

      el.innerHTML = `
        <div class="luvr-header">
          <p class="planning-meta">Источник: ${sourceLine} · <code>${esc(data.folder)}/</code></p>
          ${data.contract ? `<p class="planning-meta">Договор: ${esc(data.contract)}</p>` : ""}
          ${cacheNote}
          ${gridNote}
          <label class="field-label">Месяц</label>
          <select id="luvrMonth" class="field-input luvr-month-select"></select>
        </div>
        <div id="luvrMonthBody"></div>`;

      const sel = el.querySelector("#luvrMonth");
      const body = el.querySelector("#luvrMonthBody");
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
        body.innerHTML = renderMonthPanel(month, editable);
        if (editable) {
          const panel = body.querySelector(".luvr-month-panel");
          const statusEl = body.querySelector("#luvrSaveStatus");
          bindGrid(panel || body, month, sheet, statusEl);
        }
      }

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
