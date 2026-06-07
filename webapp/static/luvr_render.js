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

  function renderPersonProjectsSelect(pi, person, projects) {
    const ids = person.project_ids || [];
    const src = person.projects_source || (ids.length ? "auto" : "");
    const cls =
      src === "manual" ? "luvr-proj--manual" : ids.length ? "luvr-proj--ok" : "luvr-proj--empty";
    const opts = (projects || [])
      .map(
        (p) =>
          `<option value="${esc(p.id)}"${ids.includes(p.id) ? " selected" : ""}>${esc(p.title || p.id)}</option>`
      )
      .join("");
    return `<select multiple class="luvr-person-projects ${cls}" data-person-idx="${pi}" title="Проекты за месяц (Ctrl+клик)">${opts || '<option disabled>— нет проектов —</option>'}</select>`;
  }

  function monthProjectStats(month) {
    let withProjects = 0;
    const people = month.people || [];
    for (const p of people) {
      if ((p.project_ids || []).length) withProjects += 1;
    }
    return { withProjects, total: people.length, empty: people.length - withProjects };
  }

  function monthLinkStats(month) {
    let linked = 0;
    const people = month.people || [];
    for (const p of people) {
      if (p.person_id) linked += 1;
    }
    return { linked, total: people.length, unmatched: people.length - linked };
  }

  function renderGrid(month, editable, personnel, projects) {
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
          <td class="luvr-sticky luvr-col-proj">${renderPersonProjectsSelect(pi, p, projects)}</td>
          <td class="luvr-sticky luvr-col-pos" title="${esc(p.position)}">${esc(p.position || "—")}</td>
          ${cells}
          <td class="luvr-sum" data-sum-for="${pi}">${p.days_present ?? 0}</td>
        </tr>`;
      })
      .join("");

    const editHint = editable
      ? '<p class="hint-text luvr-legend">Клик / <kbd>Space</kbd> — цикл · → <span class="luvr-mark luvr-mark--full">1</span> → <span class="luvr-mark luvr-mark--half">0.5</span> · <kbd>Shift</kbd>+клик / протянуть — выделение · <kbd>Ctrl+C</kbd>/<kbd>V</kbd>/<kbd>Z</kbd> · <kbd>Ctrl+F</kbd> — поиск · стрелки / <kbd>Tab</kbd> / <kbd>Enter</kbd></p>'
      : '<p class="hint-text luvr-legend">· — нет отметки · <span class="luvr-mark luvr-mark--full">1</span> полный день · <span class="luvr-mark luvr-mark--half">0.5</span> полдня</p>';

    const searchBar = editable
      ? `<div class="luvr-search-bar">
          <label class="luvr-search-label" for="luvrFioSearch">Поиск по ФИО</label>
          <input type="search" id="luvrFioSearch" class="field-input luvr-fio-search" placeholder="Фамилия или ФИО…" autocomplete="off" spellcheck="false">
          <span id="luvrSearchInfo" class="luvr-search-info hint-text" aria-live="polite"></span>
        </div>`
      : "";

    return `${searchBar}<div class="luvr-grid-wrap">
      <table class="planning-table luvr-grid" role="grid">
        <thead>
          <tr>
            <th class="luvr-sticky luvr-col-num">№</th>
            <th class="luvr-sticky luvr-col-fio">ФИО</th>
            <th class="luvr-sticky luvr-col-link">Справ.</th>
            <th class="luvr-sticky luvr-col-proj">Объект</th>
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

  function renderMonthPanel(month, editable, personnel, projects) {
    return `<div class="luvr-month-panel">${renderGrid(month, editable, personnel, projects)}</div>`;
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

  async function savePersonProjects(sheet, personIdx, projectIds) {
    const res = await fetch("/api/luvr/projects", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        sheet,
        person_idx: personIdx,
        project_ids: projectIds,
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

  async function apiAutoProjects() {
    const res = await fetch("/api/luvr/auto-projects", { method: "POST" });
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
          person.project_ids = result.project_ids || [];
          person.projects_source = result.projects_source || null;
          sel.className = `luvr-person-link ${
            result.link_source === "unmatched"
              ? "luvr-link--warn"
              : result.link_source === "manual"
                ? "luvr-link--manual"
                : "luvr-link--ok"
          }`;
          const row = sel.closest("tr");
          if (row) row.classList.toggle("luvr-row--unlinked", !person.person_id);
          const projSel = row && row.querySelector("select.luvr-person-projects");
          if (projSel) {
            const ids = person.project_ids || [];
            Array.from(projSel.options).forEach((o) => {
              o.selected = ids.includes(o.value);
            });
            projSel.className = `luvr-person-projects ${
              person.projects_source === "manual"
                ? "luvr-proj--manual"
                : ids.length
                  ? "luvr-proj--ok"
                  : "luvr-proj--empty"
            }`;
          }
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

  function bindPersonProjects(container, month, sheet, onChanged) {
    container.querySelectorAll("select.luvr-person-projects").forEach((sel) => {
      sel.addEventListener("change", async () => {
        const personIdx = Number(sel.dataset.personIdx);
        const person = (month.people || [])[personIdx];
        if (!person) return;
        const prev = [...(person.project_ids || [])].sort().join("|");
        const nextIds = Array.from(sel.selectedOptions)
          .map((o) => o.value)
          .filter(Boolean);
        const next = [...nextIds].sort().join("|");
        if (prev === next) return;
        sel.disabled = true;
        try {
          const result = await savePersonProjects(sheet, personIdx, nextIds);
          person.project_ids = result.project_ids || [];
          person.projects_source = result.projects_source || null;
          sel.className = `luvr-person-projects ${
            result.projects_source === "manual"
              ? "luvr-proj--manual"
              : person.project_ids.length
                ? "luvr-proj--ok"
                : "luvr-proj--empty"
          }`;
          if (onChanged) onChanged();
        } catch (e) {
          Array.from(sel.options).forEach((o) => {
            o.selected = (person.project_ids || []).includes(o.value);
          });
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

  async function saveMarksBatch(sheet, updates) {
    const res = await fetch("/api/luvr/marks-batch", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sheet, updates }),
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

  function parsePasteMark(v) {
    const s = String(v ?? "").trim().replace(",", ".");
    if (!s || s === "·" || s === ".") return "";
    if (s === "0") return "";
    if (s === "1") return "1";
    if (s === "0.5" || s === ".5") return "0.5";
    return null;
  }

  function applyCellUI(td, mark) {
    td.textContent = markLabel(mark);
    td.className = `${markClass(mark)} luvr-cell`;
  }

  function bindGrid(container, month, sheet, statusEl, onXlsxStale) {
    const people = month.people || [];
    const dayCount = (month.days || []).length;
    const personCount = people.length;
    let pending = 0;
    let compose = null;
    let selAnchor = null;
    let selEnd = null;
    let dragStart = null;
    let skipClickCycle = false;
    const undoStack = [];
    const UNDO_LIMIT = 50;

    function setStatus(text, kind) {
      if (!statusEl) return;
      statusEl.textContent = text;
      statusEl.className = `luvr-save-status${kind ? ` luvr-save-status--${kind}` : ""}`;
    }

    function cellAt(personIdx, dayIdx) {
      return container.querySelector(
        `td.luvr-cell[data-person-idx="${personIdx}"][data-day-idx="${dayIdx}"]`
      );
    }

    function focusCellAt(personIdx, dayIdx) {
      const td = cellAt(personIdx, dayIdx);
      if (td) td.focus();
      return td;
    }

    function normSel(a, b) {
      return {
        r0: Math.min(a.personIdx, b.personIdx),
        r1: Math.max(a.personIdx, b.personIdx),
        c0: Math.min(a.dayIdx, b.dayIdx),
        c1: Math.max(a.dayIdx, b.dayIdx),
      };
    }

    function paintSelection() {
      container.querySelectorAll("td.luvr-cell--selected").forEach((cell) => {
        cell.classList.remove("luvr-cell--selected");
      });
      if (!selAnchor || !selEnd) return;
      const s = normSel(selAnchor, selEnd);
      for (let pi = s.r0; pi <= s.r1; pi++) {
        for (let di = s.c0; di <= s.c1; di++) {
          const cell = cellAt(pi, di);
          if (cell) cell.classList.add("luvr-cell--selected");
        }
      }
    }

    function setSelection(a, b) {
      selAnchor = a;
      selEnd = b;
      paintSelection();
    }

    function selectionRange() {
      if (!selAnchor || !selEnd) return null;
      return normSel(selAnchor, selEnd);
    }

    function isMultiSelection() {
      const s = selectionRange();
      if (!s) return false;
      return s.r0 !== s.r1 || s.c0 !== s.c1;
    }

    function collectSelectionUpdates(mark) {
      const s = selectionRange();
      if (!s) return [];
      const updates = [];
      for (let pi = s.r0; pi <= s.r1; pi++) {
        for (let di = s.c0; di <= s.c1; di++) {
          const prev = (people[pi].marks || [])[di] ?? "";
          if (prev !== mark) updates.push({ person_idx: pi, day_idx: di, mark });
        }
      }
      return updates;
    }

    function pushUndo(snapshot) {
      if (!snapshot.length) return;
      undoStack.push(snapshot);
      if (undoStack.length > UNDO_LIMIT) undoStack.shift();
    }

    async function commitBatch(updates, opts = {}) {
      const recordUndo = opts.recordUndo !== false;
      if (!updates.length) return;
      const prev = updates.map((u) => ({
        ...u,
        prev: (people[u.person_idx].marks || [])[u.day_idx] ?? "",
      }));

      updates.forEach((u) => {
        people[u.person_idx].marks = people[u.person_idx].marks || [];
        people[u.person_idx].marks[u.day_idx] = u.mark;
        const cell = cellAt(u.person_idx, u.day_idx);
        if (cell) {
          applyCellUI(cell, u.mark);
          cell.classList.add("luvr-cell--pending");
        }
      });

      pending += 1;
      setStatus(`Сохранение ${updates.length} ячеек…`, "pending");

      try {
        const result = await saveMarksBatch(sheet, updates);
        for (const ps of result.people || []) {
          const person = people[ps.person_idx];
          if (!person) continue;
          person.days_present = ps.days_present;
          person.days_marked = ps.days_marked;
          const sumCell = container.querySelector(`td.luvr-sum[data-sum-for="${ps.person_idx}"]`);
          if (sumCell) sumCell.textContent = String(ps.days_present);
        }
        prev.forEach((u) => {
          const cell = cellAt(u.person_idx, u.day_idx);
          if (!cell) return;
          cell.classList.remove("luvr-cell--pending");
          cell.classList.add("luvr-cell--saved");
          setTimeout(() => cell.classList.remove("luvr-cell--saved"), 600);
        });
        if (onXlsxStale) onXlsxStale(Boolean(result.xlsx_stale));
        if (recordUndo) {
          pushUndo(prev.map((u) => ({ person_idx: u.person_idx, day_idx: u.day_idx, mark: u.prev })));
        }
        if (result.xlsx_synced) setStatus(`Сохранено ${result.updated} ячеек (yaml + Excel)`, "ok");
        else setStatus(`Сохранено ${result.updated} ячеек (yaml)`, "ok");
        return true;
      } catch (e) {
        prev.forEach((u) => {
          people[u.person_idx].marks[u.day_idx] = u.prev;
          const cell = cellAt(u.person_idx, u.day_idx);
          if (cell) {
            applyCellUI(cell, u.prev);
            cell.classList.remove("luvr-cell--pending");
          }
        });
        setStatus(`Ошибка: ${e.message}`, "error");
        return false;
      } finally {
        pending -= 1;
      }
    }

    async function undoLast() {
      if (pending > 0) {
        setStatus("Дождитесь сохранения…", "pending");
        return;
      }
      const snapshot = undoStack.pop();
      if (!snapshot?.length) {
        setStatus("Нечего отменять", "ok");
        return;
      }
      const ok = await commitBatch(
        snapshot.map((u) => ({ person_idx: u.person_idx, day_idx: u.day_idx, mark: u.mark })),
        { recordUndo: false }
      );
      if (!ok) undoStack.push(snapshot);
      else setStatus(`Отменено ${snapshot.length} яч.`, "ok");
    }

    async function copySelection() {
      const s = selectionRange();
      if (!s) return;
      const lines = [];
      for (let pi = s.r0; pi <= s.r1; pi++) {
        const row = [];
        for (let di = s.c0; di <= s.c1; di++) {
          row.push((people[pi].marks || [])[di] ?? "");
        }
        lines.push(row.join("\t"));
      }
      try {
        await navigator.clipboard.writeText(lines.join("\n"));
        setStatus("Скопировано в буфер", "ok");
      } catch (_) {
        setStatus("Не удалось скопировать", "error");
      }
    }

    async function pasteClipboard(anchorPerson, anchorDay) {
      let text = "";
      try {
        text = await navigator.clipboard.readText();
      } catch (_) {
        setStatus("Нет доступа к буферу обмена", "error");
        return;
      }
      const rows = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
      while (rows.length && rows[rows.length - 1] === "") rows.pop();
      if (!rows.length) return;

      const updates = [];
      for (let ri = 0; ri < rows.length; ri++) {
        const cols = rows[ri].split("\t");
        for (let ci = 0; ci < cols.length; ci++) {
          const pi = anchorPerson + ri;
          const di = anchorDay + ci;
          if (pi >= personCount || di >= dayCount) continue;
          const mark = parsePasteMark(cols[ci]);
          if (mark === null) continue;
          const prev = (people[pi].marks || [])[di] ?? "";
          if (prev !== mark) updates.push({ person_idx: pi, day_idx: di, mark });
        }
      }
      if (!updates.length) {
        setStatus("Вставка: нет изменений", "ok");
        return;
      }

      const pasteEndPi = Math.min(anchorPerson + rows.length - 1, personCount - 1);
      const pasteEndDi = Math.min(
        anchorDay + Math.max(...rows.map((r) => r.split("\t").length)) - 1,
        dayCount - 1
      );
      await commitBatch(updates);
      setSelection(
        { personIdx: anchorPerson, dayIdx: anchorDay },
        { personIdx: pasteEndPi, dayIdx: pasteEndDi }
      );
    }

    function parseComposeText(text) {
      const s = String(text ?? "").trim().replace(",", ".");
      if (!s || s === "0") return "";
      if (s === "1") return "1";
      if (s === "0.5" || s === ".5") return "0.5";
      return null;
    }

    function showCompose(td, text) {
      td.textContent = text || " ";
      td.classList.add("luvr-cell--compose");
    }

    function endCompose(commit) {
      if (!compose) return Promise.resolve(null);
      const { td, personIdx, dayIdx, text, prevMark } = compose;
      compose = null;
      td.classList.remove("luvr-cell--compose");
      if (!commit) {
        applyCellUI(td, prevMark);
        return Promise.resolve(null);
      }
      const mark = parseComposeText(text);
      if (mark === null) {
        applyCellUI(td, prevMark);
        return Promise.resolve(null);
      }
      return commitCell(td, personIdx, dayIdx, mark);
    }

    function navigate(fromPerson, fromDay, dPerson, dDay, wrapTab) {
      let pi = fromPerson;
      let di = fromDay + dDay;
      if (wrapTab && dDay !== 0) {
        if (di >= dayCount) {
          di = 0;
          pi += 1;
        } else if (di < 0) {
          di = dayCount - 1;
          pi -= 1;
        }
      } else {
        pi += dPerson;
      }
      if (pi < 0 || pi >= personCount || di < 0 || di >= dayCount) return null;
      return focusCellAt(pi, di);
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
        pushUndo([{ person_idx: personIdx, day_idx: dayIdx, mark: prev }]);
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
      td.addEventListener("focus", () => {
        if (compose && compose.td !== td) endCompose(true);
      });
    });

    container.addEventListener("mousedown", (ev) => {
      const td = ev.target.closest("td.luvr-cell");
      if (!td || ev.button !== 0 || ev.shiftKey) return;
      dragStart = {
        personIdx: Number(td.dataset.personIdx),
        dayIdx: Number(td.dataset.dayIdx),
      };
    });

    container.addEventListener("mouseover", (ev) => {
      if (!dragStart || !(ev.buttons & 1)) return;
      const td = ev.target.closest("td.luvr-cell");
      if (!td) return;
      const personIdx = Number(td.dataset.personIdx);
      const dayIdx = Number(td.dataset.dayIdx);
      if (personIdx === dragStart.personIdx && dayIdx === dragStart.dayIdx) return;
      skipClickCycle = true;
      setSelection(dragStart, { personIdx, dayIdx });
    });

    container.addEventListener("mouseup", () => {
      dragStart = null;
    });

    container.querySelectorAll("td.luvr-cell").forEach((td) => {
      const personIdx = Number(td.dataset.personIdx);
      const dayIdx = Number(td.dataset.dayIdx);

      td.addEventListener("click", (ev) => {
        if (compose && compose.td !== td) endCompose(true);
        td.focus();
        if (ev.shiftKey && selAnchor) {
          setSelection(selAnchor, { personIdx, dayIdx });
          return;
        }
        setSelection({ personIdx, dayIdx }, { personIdx, dayIdx });
        selAnchor = { personIdx, dayIdx };
        if (skipClickCycle) {
          skipClickCycle = false;
          return;
        }
        const cur = (people[personIdx]?.marks || [])[dayIdx] ?? "";
        commitCell(td, personIdx, dayIdx, cycleMark(cur));
      });
    });

    container.addEventListener("keydown", async (ev) => {
      const td = ev.target.closest("td.luvr-cell");
      if (!td || !container.contains(td)) return;

      const personIdx = Number(td.dataset.personIdx);
      const dayIdx = Number(td.dataset.dayIdx);
      const cur = (people[personIdx]?.marks || [])[dayIdx] ?? "";

      if ((ev.ctrlKey || ev.metaKey) && ev.key === "f") {
        ev.preventDefault();
        const search = container.querySelector("#luvrFioSearch");
        if (search) {
          search.focus();
          search.select();
        }
        return;
      }
      if ((ev.ctrlKey || ev.metaKey) && ev.key === "z" && !ev.shiftKey) {
        ev.preventDefault();
        await endCompose(true);
        undoLast();
        return;
      }
      if ((ev.ctrlKey || ev.metaKey) && ev.key === "c") {
        ev.preventDefault();
        await endCompose(true);
        copySelection();
        return;
      }
      if ((ev.ctrlKey || ev.metaKey) && ev.key === "v") {
        ev.preventDefault();
        await endCompose(true);
        const s = selectionRange();
        const anchorPerson = s ? s.r0 : personIdx;
        const anchorDay = s ? s.c0 : dayIdx;
        pasteClipboard(anchorPerson, anchorDay);
        return;
      }

      if (compose && compose.td === td) {
        if (ev.key === "Escape") {
          ev.preventDefault();
          endCompose(false);
          return;
        }
        if (ev.key === "Backspace") {
          ev.preventDefault();
          compose.text = compose.text.slice(0, -1);
          showCompose(td, compose.text);
          return;
        }
        if (ev.key.length === 1 && !ev.ctrlKey && !ev.metaKey && !ev.altKey) {
          if (/^[0-9.,]$/.test(ev.key)) {
            ev.preventDefault();
            compose.text += ev.key;
            showCompose(td, compose.text);
            return;
          }
        }
        if (
          ev.key === "Enter" ||
          ev.key === "Tab" ||
          ev.key === "ArrowUp" ||
          ev.key === "ArrowDown" ||
          ev.key === "ArrowLeft" ||
          ev.key === "ArrowRight"
        ) {
          ev.preventDefault();
          await endCompose(true);
          if (ev.key === "Tab") navigate(personIdx, dayIdx, 0, ev.shiftKey ? -1 : 1, true);
          else if (ev.key === "Enter") navigate(personIdx, dayIdx, ev.shiftKey ? -1 : 1, 0, false);
          else if (ev.key === "ArrowUp") navigate(personIdx, dayIdx, -1, 0, false);
          else if (ev.key === "ArrowDown") navigate(personIdx, dayIdx, 1, 0, false);
          else if (ev.key === "ArrowLeft") navigate(personIdx, dayIdx, 0, -1, false);
          else if (ev.key === "ArrowRight") navigate(personIdx, dayIdx, 0, 1, false);
          return;
        }
        return;
      }

      if (ev.key === "Tab") {
        ev.preventDefault();
        await endCompose(true);
        navigate(personIdx, dayIdx, 0, ev.shiftKey ? -1 : 1, true);
        return;
      }
      if (ev.key === "Enter") {
        ev.preventDefault();
        await endCompose(true);
        navigate(personIdx, dayIdx, ev.shiftKey ? -1 : 1, 0, false);
        return;
      }
      if (ev.key === "ArrowUp") {
        ev.preventDefault();
        await endCompose(true);
        navigate(personIdx, dayIdx, -1, 0, false);
        return;
      }
      if (ev.key === "ArrowDown") {
        ev.preventDefault();
        await endCompose(true);
        navigate(personIdx, dayIdx, 1, 0, false);
        return;
      }
      if (ev.key === "ArrowLeft") {
        ev.preventDefault();
        await endCompose(true);
        navigate(personIdx, dayIdx, 0, -1, false);
        return;
      }
      if (ev.key === "ArrowRight") {
        ev.preventDefault();
        await endCompose(true);
        navigate(personIdx, dayIdx, 0, 1, false);
        return;
      }
      if (ev.key === " ") {
        ev.preventDefault();
        await endCompose(true);
        commitCell(td, personIdx, dayIdx, cycleMark(cur));
        return;
      }
      if (ev.key === "Delete" || ev.key === "Backspace") {
        ev.preventDefault();
        await endCompose(true);
        if (isMultiSelection()) {
          await commitBatch(collectSelectionUpdates(""));
        } else {
          commitCell(td, personIdx, dayIdx, "");
        }
        return;
      }
      if (ev.key === "1") {
        ev.preventDefault();
        await endCompose(true);
        commitCell(td, personIdx, dayIdx, "1");
        return;
      }
      if (ev.key === "0") {
        ev.preventDefault();
        compose = { td, personIdx, dayIdx, text: "0", prevMark: cur };
        showCompose(td, "0");
        return;
      }
      if (ev.key === "." || ev.key === ",") {
        ev.preventDefault();
        compose = { td, personIdx, dayIdx, text: ev.key, prevMark: cur };
        showCompose(td, ev.key);
        return;
      }
    });

    bindFioSearch(container, people);
  }

  function normFioSearch(s) {
    return String(s ?? "")
      .toLowerCase()
      .replace(/\s+/g, " ")
      .trim();
  }

  function bindFioSearch(container, people) {
    const input = container.querySelector("#luvrFioSearch");
    const info = container.querySelector("#luvrSearchInfo");
    if (!input) return;

    let hits = [];
    let matchIdx = 0;

    function clearHits() {
      container.querySelectorAll("tr.luvr-row--search-hit").forEach((tr) => {
        tr.classList.remove("luvr-row--search-hit");
      });
    }

    function focusMatch(idx) {
      if (!hits.length) return;
      matchIdx = ((idx % hits.length) + hits.length) % hits.length;
      const pi = hits[matchIdx];
      const tr = container.querySelector(`tr[data-person-idx="${pi}"]`);
      if (tr) {
        tr.scrollIntoView({ block: "nearest", behavior: "smooth" });
        const cell = container.querySelector(
          `td.luvr-cell[data-person-idx="${pi}"][data-day-idx="0"]`
        );
        if (cell) cell.focus();
      }
      info.textContent = `${matchIdx + 1} / ${hits.length}`;
    }

    function applySearch() {
      const q = normFioSearch(input.value);
      clearHits();
      hits = [];
      matchIdx = 0;
      if (!q) {
        info.textContent = "";
        return;
      }
      people.forEach((p, pi) => {
        if (normFioSearch(p.fio).includes(q)) hits.push(pi);
      });
      hits.forEach((pi) => {
        const tr = container.querySelector(`tr[data-person-idx="${pi}"]`);
        if (tr) tr.classList.add("luvr-row--search-hit");
      });
      if (!hits.length) {
        info.textContent = "нет совпадений";
        return;
      }
      info.textContent = `${hits.length} совпад.`;
      focusMatch(0);
    }

    input.addEventListener("input", applySearch);
    input.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter") {
        ev.preventDefault();
        if (hits.length) focusMatch(matchIdx + (ev.shiftKey ? -1 : 1));
        return;
      }
      if (ev.key === "Escape") {
        input.value = "";
        applySearch();
        input.blur();
      }
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

  async function apiBuildAppendix7(sheet) {
    const res = await fetch("/api/luvr/appendix7/build", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sheet }),
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

  function renderSyncBar(
    xlsxPresent,
    xlsxStale,
    linkStats,
    monthStats,
    projectStats,
    monthProjStats,
    appendix7,
    deployment
  ) {
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
    if (monthProjStats) {
      parts.push(
        `<span class="luvr-link-summary">Объекты: ${monthProjStats.withProjects}/${monthProjStats.total}${
          monthProjStats.empty ? ` · ${monthProjStats.empty} без объекта` : ""
        }</span>`
      );
    } else if (projectStats) {
      parts.push(
        `<span class="luvr-link-summary">Объекты: ${projectStats.with_projects}/${projectStats.total}</span>`
      );
    }
    if (xlsxPresent) {
      parts.push(
        xlsxStale
          ? '<span class="luvr-sync-stale">Excel не синхронизирован</span>'
          : '<span class="luvr-sync-ok">yaml ↔ Excel</span>'
      );
    }
    if (appendix7) {
      if (appendix7.template_present) {
        parts.push(
          `<span class="luvr-link-summary">Прил.7: шаблон · ${appendix7.summary_rows || 0} строк ФИО</span>`
        );
      } else if (appendix7.template_error) {
        parts.push(`<span class="luvr-sync-stale" title="${esc(appendix7.template_error)}">Прил.7: нет шаблона</span>`);
      }
    }
    if (deployment) {
      if (deployment.template_present) {
        const ast = deployment.assignments || {};
        const assignHint =
          ast.assignments != null
            ? `${ast.assignments} назнач. · расстановка → <a href="/planning/projects">Проекты</a>`
            : `Расстановка из справочника → <a href="/planning/projects">Проекты</a>`;
        parts.push(`<span class="luvr-link-summary">${assignHint}</span>`);
      } else if (deployment.template_error) {
        parts.push(
          `<span class="luvr-sync-stale" title="${esc(deployment.template_error)}">Расстановка: нет шаблона</span>`
        );
      }
    }
    const actions = [
      `<button type="button" class="btn btn-gray btn-sm" id="luvrAutoLink">Авто-связать по ФИО</button>`,
      `<button type="button" class="btn btn-gray btn-sm" id="luvrAutoProjects">Проекты из справочника</button>`,
    ];
    if (appendix7 && appendix7.template_present) {
      actions.push(
        `<button type="button" class="btn btn-green btn-sm" id="luvrBuildAppendix7">Сформировать Прил.7</button>`
      );
    }
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
          <div id="luvrSyncBar">${renderSyncBar(state.xlsx_present, state.xlsx_stale, state.link_stats, null, state.project_stats, null, state.appendix7, state.deployment)}</div>
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
        const mps = month ? monthProjectStats(month) : null;
        if (syncBar)
          syncBar.innerHTML = renderSyncBar(
            state.xlsx_present,
            state.xlsx_stale,
            state.link_stats,
            ms,
            state.project_stats,
            mps,
            state.appendix7,
            state.deployment
          );
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
        const autoProjectsBtn = el.querySelector("#luvrAutoProjects");
        const appendix7Btn = el.querySelector("#luvrBuildAppendix7");
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
        if (autoProjectsBtn) {
          autoProjectsBtn.onclick = async () => {
            autoProjectsBtn.disabled = true;
            setGlobalStatus("Назначение проектов из справочника…", "pending");
            try {
              const result = await apiAutoProjects();
              const res = await fetch("/api/luvr");
              if (!res.ok) throw new Error(res.statusText);
              window.LuvrPanel.render(el, await res.json());
              setGlobalStatus(
                `Проекты: ${result.auto} авто, ${result.manual} вручную, ${result.empty} без объекта`,
                "ok"
              );
            } catch (e) {
              setGlobalStatus(`Проекты: ${e.message}`, "error");
              autoProjectsBtn.disabled = false;
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
        if (appendix7Btn) {
          appendix7Btn.onclick = async () => {
            const sheet = sel.value;
            appendix7Btn.disabled = true;
            setGlobalStatus("Формирование Прил.7…", "pending");
            try {
              const result = await apiBuildAppendix7(sheet);
              let msg = `Прил.7: ${result.filled_rows} строк, ${result.cells_updated} ячеек`;
              if (result.unmatched_luvr_count) {
                msg += ` · ${result.unmatched_luvr_count} ФИО из ЛУВР не в шаблоне`;
              }
              if (result.unmatched_template_count) {
                msg += ` · ${result.unmatched_template_count} строк шаблона без данных`;
              }
              setGlobalStatus(msg, "ok");
              if (result.download) {
                window.location.href = result.download;
              }
            } catch (e) {
              setGlobalStatus(`Прил.7: ${e.message}`, "error");
            } finally {
              appendix7Btn.disabled = false;
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
        body.innerHTML = renderMonthPanel(month, editable, state.personnel || [], state.projects || []);
        refreshSyncBar(month);
        const panel = body.querySelector(".luvr-month-panel");
        bindPersonLinks(panel || body, month, sheet, state.personnel || [], () => refreshSyncBar(month));
        bindPersonProjects(panel || body, month, sheet, () => refreshSyncBar(month));
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
